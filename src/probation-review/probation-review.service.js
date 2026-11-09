/**
 * src/probation-review/probation-review.service.js
 *
 * =========================================================================================
 * SERVICE: PROBATION REVIEW AUTOMATION
 * =========================================================================================
 * Mục tiêu: Tự động gửi email / tạo bản nháp đánh giá hết hạn thử việc (Probation / Intern / Apprentice).
 *
 * Quy trình tổng quát gồm 3 chặng:
 *   1. Đọc dữ liệu hồ sơ gốc từ Sheet Onboarding (UC03) nạp vào RAM (Map tra cứu siêu tốc).
 *   2. Đọc Sheet điều khiển Contract Tracking (UC05) và bảng lịch sử "Probation Review Logs".
 *   3. Duyệt từng nhân sự trên UC05, kiểm tra điều kiện đến hạn, tạo draft/gửi mail và ghi log.
 */
class ProbationReviewService {
  /**
   * Helper: Phân loại nhãn đánh giá dựa vào Employment Type & Contract Type
   * Ví dụ: Chứa chữ "intern" -> Intern Review, còn lại mặc định là Probation Review.
   */
  static resolveReviewTypeLabel_(empTypeRaw, contractTypeRaw) {
    const raw = `${empTypeRaw || ""} ${contractTypeRaw || ""}`.toLowerCase();
    if (raw.includes("intern")) return "Intern";
    if (raw.includes("apprentice")) return "Apprentice";
    return "Probation";
  }

  // =======================================================================================
  // CHẶNG 1: ĐỌC VÀ LẬP CHỈ MỤC DỮ LIỆU TỪ SHEET ONBOARDING (UC03)
  // =======================================================================================
  /**
   * Nạp toàn bộ dữ liệu nhân sự từ Sheet Onboarding vào 2 cấu trúc Map trong RAM:
   *   - byEmail: Tìm thông tin nhanh bằng Email
   *   - byName:  Tìm thông tin dự phòng bằng Họ và Tên (nếu bên UC05 thiếu email)
   * Kỹ thuật: Dùng getDataRange() & getRichTextValues() để đọc 1 lần toàn bộ bảng (Batch Read).
   */
  static loadOnboardingDataMap_() {
    const oConfig = PROBATION_REVIEW_CONFIG.ONBOARDING;
    const ss = SpreadsheetApp.openById(oConfig.SPREADSHEET_ID);

    // Tìm tab sheet theo TAB_GID được cấu hình
    const sheets = ss.getSheets();
    const sheet = sheets.find((s) => String(s.getSheetId()) === String(oConfig.TAB_GID)) || sheets[0];

    // Đọc đồng thời: Giá trị hiển thị và RichText (chứa link ẩn của Folder OKR)
    const range = sheet.getDataRange();
    const data = range.getValues();
    const richTexts = range.getRichTextValues();

    if (data.length < oConfig.DATA_START_ROW_INDEX) {
      return { byEmail: new Map(), byName: new Map() };
    }

    // Xác định vị trí index các cột động từ dòng tiêu đề (Header Row)
    const headers = data[oConfig.HEADER_ROW_INDEX - 1].map(normalizeHeaderKey);
    const fullNameIdx = headers.indexOf(oConfig.COLS.FULL_NAME);
    const personalEmailIdx = headers.indexOf(oConfig.COLS.PERSONAL_EMAIL);
    const workingEmailIdx = headers.indexOf(oConfig.COLS.WORKING_EMAIL);
    const stopDateIdx = headers.indexOf(oConfig.COLS.STOP_PROBATION_INTERN_DATE);
    const okrFolderIdx = headers.indexOf(oConfig.COLS.LINK_FOLDER_OKRS);
    const titleIdx = headers.indexOf(oConfig.COLS.TITLE);
    const squadIdx = headers.indexOf(oConfig.COLS.SQUAD_UNIT);
    const tribeIdx = headers.indexOf(oConfig.COLS.TRIBE_DIVISION);
    const lmIdx = headers.indexOf(oConfig.COLS.LINE_MANAGER);
    const dlmIdx = headers.indexOf(oConfig.COLS.DOT_LINE_MANAGER);
    const empTypeIdx = headers.indexOf(oConfig.COLS.EMPLOYMENT_TYPE);

    const byEmail = new Map();
    const byName = new Map();
    const startRow = oConfig.DATA_START_ROW_INDEX - 1;

    for (let r = startRow; r < data.length; r++) {
      const row = data[r];
      const fullName = String(row[fullNameIdx] || "").trim();
      if (!fullName) continue;

      const personalEmail = ProbationReviewUtil.cleanEmail(row[personalEmailIdx]);
      const workingEmail = ProbationReviewUtil.cleanEmail(row[workingEmailIdx]);
      const squad = String(row[squadIdx] || "").trim();
      const tribe = String(row[tribeIdx] || "").trim();

      // Bóc tách link Google Drive thư mục OKR (hỗ trợ cả hyperlink ẩn trong chữ lẫn text thô)
      const okrFolderUrl = okrFolderIdx >= 0 ? ProbationReviewUtil.extractCellUrl(row[okrFolderIdx], richTexts[r]?.[okrFolderIdx]) : "";

      // Gộp Squad/Tribe hiển thị (tránh bị lặp lại nếu hai cột trùng tên nhau)
      const squadTribe = squad && tribe && squad.toLowerCase() === tribe.toLowerCase() ? squad : [squad, tribe].filter(Boolean).join(" / ");

      const record = {
        fullName,
        personalEmail,
        workingEmail,
        targetEmail: personalEmail || workingEmail,
        stopProbationDate: row[stopDateIdx],
        stopProbationDisplayDate: ProbationReviewUtil.formatDisplayDate(row[stopDateIdx]),
        okrFolderUrl,
        title: String(row[titleIdx] || "").trim(),
        squadTribe,
        lineManager: ProbationReviewUtil.cleanEmail(row[lmIdx]),
        dotLineManager: ProbationReviewUtil.cleanEmail(row[dlmIdx]),
        employmentType: String(row[empTypeIdx] || "").trim(),
      };

      if (personalEmail) byEmail.set(personalEmail, record);
      if (workingEmail) byEmail.set(workingEmail, record);
      byName.set(fullName.toLowerCase(), record);
    }

    return { byEmail, byName };
  }

  // =======================================================================================
  // CHẶNG 2: QUẢN LÝ LỊCH SỬ GỬI & CHỐNG TRÙNG LẶP (LOG TRACKING)
  // =======================================================================================
  /**
   * Mở hoặc tự khởi tạo tab "Probation Review Logs" trên file Spreadsheet UC05
   */
  static getOrCreateLogSheet_(ss) {
    const LOG_TAB_NAME = "Probation Review Logs";
    let logSheet = ss.getSheetByName(LOG_TAB_NAME);

    // Nếu chưa tồn tại tab Log thì tự động tạo mới kèm dòng Header
    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_TAB_NAME);
      logSheet.appendRow(["Timestamp", "Row UC05", "Full Name", "Target Email", "Review Type", "Status", "Mode", "Subject"]);
      logSheet.getRange("1:1").setFontWeight("bold");
    }
    return logSheet;
  }

  /**
   * Nạp danh sách các lượt gửi thành công trong quá khứ từ tab Log vào Set
   * Khóa định danh dạng: "{RowIndex}_{Email}" -> Đảm bảo kiểm tra tức thì O(1)
   */
  static loadProcessedSetFromLog_(logSheet) {
    const processedSet = new Set();
    const data = logSheet.getDataRange().getValues();

    // Bắt đầu đọc từ dòng thứ 2 (bỏ qua dòng tiêu đề)
    for (let i = 1; i < data.length; i++) {
      const rowNum = data[i][1];
      const email = String(data[i][3] || "")
        .trim()
        .toLowerCase();
      if (rowNum && email) {
        processedSet.add(`${rowNum}_${email}`);
      }
    }
    return processedSet;
  }

  // =======================================================================================
  // CHẶNG 3: LUỒNG QUÉT CHÍNH (MAIN DISPATCHER LOOP)
  // =======================================================================================
  /**
   * Quét bảng UC05, lọc danh sách đến hạn và thực thi tạo Draft hoặc gửi Email
   * @param {"draft"|"send"} mode - "draft": tạo thư nháp trong Gmail; "send": gửi thật
   */
  static runProbationReview(mode = "draft") {
    const tConfig = PROBATION_REVIEW_CONFIG.PROGRESS_TRACKING;
    const timeZone = PROBATION_REVIEW_CONFIG.TIMEZONE;
    const todayKey = ProbationReviewUtil.formatDateKey(new Date(), timeZone);

    // Lấy thông tin người chạy script hiện tại để ký tên cuối email
    const currentUserEmail = Session.getActiveUser().getEmail();
    const senderName = currentUserEmail
      ? currentUserEmail
          .split("@")[0]
          .replace(".", " ")
          .replace(/\b\w/g, (c) => c.toUpperCase())
      : "Talent Development Team";

    // 1. Tải Map Onboarding từ UC03 vào RAM
    const { byEmail, byName } = this.loadOnboardingDataMap_();

    // 2. Mở file UC05 & Chuẩn bị Set tra cứu chống gửi trùng
    const ssUC05 = SpreadsheetApp.openById(tConfig.SPREADSHEET_ID);
    const sheetUC05 = ssUC05.getSheetByName(tConfig.TAB_NAME) || ssUC05.getSheets()[0];
    const logSheet = this.getOrCreateLogSheet_(ssUC05);
    const processedSet = this.loadProcessedSetFromLog_(logSheet);

    const dataUC05 = sheetUC05.getDataRange().getValues();
    if (dataUC05.length < tConfig.DATA_START_ROW_INDEX) {
      return {
        ok: true,
        processedCount: 0,
        message: "Sheet UC05 Contract Tracking không có dữ liệu.",
      };
    }

    // Map vị trí các cột trên Sheet UC05
    const headers = dataUC05[tConfig.HEADER_ROW_INDEX - 1].map(normalizeHeaderKey);
    const nameIdx = headers.indexOf(tConfig.COLS.FULL_NAME);
    const picIdx = headers.indexOf(tConfig.COLS.PIC);
    const contractTypeIdx = headers.indexOf(tConfig.COLS.TYPE_OF_CONTRACT);
    const startProcessIdx = headers.indexOf(tConfig.COLS.HR_OPS_START_PROCESS);
    const consolidationIdx = headers.indexOf(tConfig.COLS.HR_OPS_CONSOLIDATION);
    const candidateEmailIdx = headers.indexOf(tConfig.COLS.CANDIDATE_EMAIL);

    const summary = {
      ok: true,
      mode: mode,
      processedCount: 0,
      processed: [],
      skipped: [],
      errors: [],
    };

    const newLogs = []; // Danh sách các dòng log mới cần ghi sau khi kết thúc vòng lặp
    const startRow = tConfig.DATA_START_ROW_INDEX - 1;

    // 3. Duyệt từng dòng nhân sự trên UC05
    for (let r = startRow; r < dataUC05.length; r++) {
      const row = dataUC05[r];
      const sheetRowNumberUC05 = r + 1;
      const rawName = String(row[nameIdx] || "").trim();
      const rawEmail = String(row[candidateEmailIdx] || "")
        .trim()
        .toLowerCase();
      const rawStartProcess = row[startProcessIdx];
      const rawConsolidation = consolidationIdx >= 0 ? row[consolidationIdx] : null;
      const rawPic = String(row[picIdx] || "").trim();
      const rawContractType = String(row[contractTypeIdx] || "").trim();

      // Bỏ qua dòng trống không có tên
      if (!rawName) continue;

      // Tra cứu chéo sang Onboarding UC03 (ưu tiên khớp Email trước, fallback sang Họ tên)
      let onboardingInfo = null;
      if (rawEmail) onboardingInfo = byEmail.get(rawEmail);
      if (!onboardingInfo && rawName) onboardingInfo = byName.get(rawName.toLowerCase());

      // Xác định email nhận tin: Ưu tiên Personal Email (hoặc fallback sang Working Email)
      const targetEmail = (onboardingInfo && onboardingInfo.personalEmail) || rawEmail || (onboardingInfo ? onboardingInfo.workingEmail : "");

      // LỌC 1: Kiểm tra tính hợp lệ của Email
      if (!targetEmail || !ProbationReviewUtil.isValidEmail(targetEmail)) {
        summary.skipped.push(`Dòng ${sheetRowNumberUC05} (${rawName}): Không tìm thấy Email hợp lệ.`);
        continue;
      }

      // LỌC 2: Kiểm tra chống gửi trùng lặp (Deduplication Check)
      const logKey = `${sheetRowNumberUC05}_${targetEmail.toLowerCase()}`;
      if (processedSet.has(logKey)) {
        continue; // Đã từng xử lý thành công trước đây -> Bỏ qua
      }

      // LỌC 3: Kiểm tra mốc ngày đến hạn
      const stopDate = onboardingInfo ? onboardingInfo.stopProbationDate : null;
      const isDue = ProbationReviewUtil.isDueReview(rawStartProcess, stopDate, PROBATION_REVIEW_CONFIG.DAYS_BEFORE_STOP, todayKey, timeZone);
      if (!isDue) continue; // Chưa tới ngày hoặc đã qua ngày -> Bỏ qua

      // ===================================================================================
      // THỰC THI: SOẠN VÀ GỬI / TẠO DRAFT EMAIL
      // ===================================================================================
      try {
        const reviewTypeName = this.resolveReviewTypeLabel_(onboardingInfo?.employmentType, rawContractType);
        const reviewTypeLower = reviewTypeName.toLowerCase();

        // Xây dựng Tiêu đề Email (Subject):
        // Cấu trúc: [ReviewType] Review - [Employee Name] | [Title] | [Squad/Tribe]
        const titleStr = onboardingInfo?.title || "";
        const squadTribeStr = onboardingInfo?.squadTribe || "";
        const metaSuffix = [titleStr, squadTribeStr].filter(Boolean).join(" | ");
        const subject = `${reviewTypeName} Review - ${rawName}${metaSuffix ? ` | ${metaSuffix}` : ""}`;

        // Xây dựng Danh sách CC:
        // Gồm: Line Manager, Dot Line Manager, people@kyanon.digital, TD in charge (cột PIC)
        const tdEmail = ProbationReviewUtil.cleanEmail(rawPic);
        const ccSet = new Set();
        if (onboardingInfo?.lineManager) ccSet.add(onboardingInfo.lineManager);
        if (onboardingInfo?.dotLineManager) ccSet.add(onboardingInfo.dotLineManager);
        ccSet.add(PROBATION_REVIEW_CONFIG.EMAIL.PEOPLE_TEAM_EMAIL);
        if (tdEmail) ccSet.add(tdEmail);
        ccSet.delete(targetEmail.toLowerCase()); // Tránh tự CC chính người nhận
        const ccList = Array.from(ccSet).join(", ");

        // Chuẩn hóa định dạng ngày hiển thị trong email (DD/MM/YYYY)
        const stopProbationDisplay = onboardingInfo ? onboardingInfo.stopProbationDisplayDate : "[DD/MM/YYYY]";
        const consolidationDisplay = rawConsolidation ? ProbationReviewUtil.formatDisplayDate(rawConsolidation, timeZone) : stopProbationDisplay;

        // Render nội dung từ file template HTML
        const template = HtmlService.createTemplateFromFile("probation-review/probation-review-email");
        template.fullName = rawName;
        template.reviewTypeName = reviewTypeName;
        template.reviewTypeLower = reviewTypeLower;
        template.stopProbationDate = stopProbationDisplay;
        template.consolidationDate = consolidationDisplay;
        template.okrFolderUrl = onboardingInfo ? onboardingInfo.okrFolderUrl : ""; // Link Folder OKRs từ UC03
        template.senderName = senderName;

        const htmlBody = template.evaluate().getContent();
        const mailOptions = {
          htmlBody: htmlBody,
          cc: ccList || undefined,
        };

        // Thực hiện gửi hoặc tạo draft qua GmailApp
        if (mode === "send") {
          GmailApp.sendEmail(targetEmail, subject, "", mailOptions);
        } else {
          GmailApp.createDraft(targetEmail, subject, "", mailOptions);
        }

        // Cập nhật thống kê kết quả
        summary.processedCount++;
        summary.processed.push({
          rowUC05: sheetRowNumberUC05,
          name: rawName,
          type: reviewTypeName,
          email: targetEmail,
          cc: ccList,
          subject: subject,
          stopDate: stopProbationDisplay,
          consolidationDate: consolidationDisplay,
          folderUrl: onboardingInfo?.okrFolderUrl || "",
        });

        // Chuẩn bị dữ liệu ghi log
        newLogs.push([new Date(), sheetRowNumberUC05, rawName, targetEmail, reviewTypeName, "SUCCESS", mode.toUpperCase(), subject]);

        // Cập nhật ngay vào Set tạm thời để tránh lặp trong cùng phiên chạy
        processedSet.add(logKey);
      } catch (err) {
        summary.errors.push(`Dòng ${sheetRowNumberUC05} (${rawName}): ${err.message}`);
      }
    }

    // Ghi toàn bộ lịch sử thành công xuống Sheet Log trong 1 lần ghi duy nhất (Batch Append)
    if (newLogs.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, newLogs.length, newLogs[0].length).setValues(newLogs);
    }

    summary.message = `Đã ${mode === "send" ? "gửi" : "tạo draft"} thành công cho ${summary.processedCount} nhân sự trên UC05.`;
    return summary;
  }
}
