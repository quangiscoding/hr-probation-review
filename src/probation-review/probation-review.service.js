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
 *   3. Duyệt từng nhân sự trên UC05, kiểm tra điều kiện đến hạn, tạo draft/gửi mail, ghi log
 *      và gửi báo cáo tổng kết qua email cho TA/PIC.
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
  /**
   * 1. Đọc Onboarding Sheet (UC03) về RAM Map tra cứu O(1)
   */
  static loadOnboardingDataMap_() {
    const oConfig = PROBATION_REVIEW_CONFIG.ONBOARDING;
    const ss = SpreadsheetApp.openById(oConfig.SPREADSHEET_ID);
    const sheets = ss.getSheets();
    const sheet = sheets.find((s) => String(s.getSheetId()) === String(oConfig.TAB_GID)) || sheets[0];

    const range = sheet.getDataRange();
    const data = range.getValues();
    const richTexts = range.getRichTextValues();

    if (data.length < oConfig.DATA_START_ROW_INDEX) {
      return { byEmail: new Map(), byName: new Map() };
    }

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
    const taIdx = headers.indexOf(oConfig.COLS.TA_IN_CHARGE); // 👈 Bổ sung đọc cột TA In Charge

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
      const okrFolderUrl = okrFolderIdx >= 0 ? ProbationReviewUtil.extractCellUrl(row[okrFolderIdx], richTexts[r]?.[okrFolderIdx]) : "";
      const squadTribe = squad && tribe && squad.toLowerCase() === tribe.toLowerCase() ? squad : [squad, tribe].filter(Boolean).join(" / ");
      const taInCharge = taIdx >= 0 ? ProbationReviewUtil.cleanEmail(row[taIdx]) : ""; // 👈 Lấy email TA In Charge

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
        taInCharge: taInCharge, // 👈 Lưu vào record để CC
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
    if (!logSheet) {
      logSheet = ss.insertSheet(LOG_TAB_NAME);
      logSheet.appendRow(["Timestamp", "Row UC05", "Full Name", "Target Email", "Review Type", "Status", "Mode", "Subject"]);
      logSheet.getRange("1:1").setFontWeight("bold");
    }
    return logSheet;
  }

  /**
   * Nạp danh sách các lượt gửi thành công trong quá khứ từ tab Log vào Set
   * Khóa định danh dạng: "{RowIndex}_{Email}" -> Đảm bảo kiểm tra tức thời O(1)
   */
  static loadProcessedSetFromLog_(logSheet) {
    const processedSet = new Set();
    const data = logSheet.getDataRange().getValues();
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
  // THÔNG BÁO: GỬI EMAIL BÁO CÁO TỔNG HỢP CHO TA / PIC PHỤ TRÁCH
  // =======================================================================================
  /**
   * Gửi email báo cáo tổng kết sau khi phiên cron quét hoàn tất
   * Gọi HTML Template từ file external "probation-review/ta-report-email"
   */
  /**
   * 2. THÔNG BÁO: Gửi email báo cáo tổng hợp chỉ cho người thao tác (PIC) & CC TA In Charge
   */
  static sendSummaryReportToTA_(summary, currentUserEmail, relatedTaEmails = []) {
    if (!currentUserEmail) return;
    if (summary.processedCount === 0 && summary.warnings.length === 0 && summary.errors.length === 0) {
      return;
    }

    const todayStr = ProbationReviewUtil.formatDisplayDate(new Date());
    const subject = `[Hera | Auto-Report] Báo cáo quét Probation Review ngày ${todayStr}`;
    const picAllocCode = currentUserEmail.split("@")[0].trim(); // Tách lấy "quang.nguyenminh"

    try {
      const template = HtmlService.createTemplateFromFile("probation-review/ta-report-email");
      template.picAllocCode = picAllocCode; // 👈 Truyền Alloc Code vào HTML Template
      template.mode = summary.mode;
      template.processedCount = summary.processedCount;
      template.processed = summary.processed;
      template.warnings = summary.warnings;
      template.errors = summary.errors;

      const htmlBody = template.evaluate().getContent();
      const mailOptions = { htmlBody: htmlBody };

      if (relatedTaEmails && relatedTaEmails.length > 0) {
        mailOptions.cc = relatedTaEmails.join(", ");
      }

      GmailApp.sendEmail(currentUserEmail, subject, "", mailOptions);
    } catch (err) {
      Logger.log(`⚠️ Không thể gửi email báo cáo cho TA/PIC: ${err.message}`);
    }
  }

  // =======================================================================================
  // CHẶNG 3: LUỒNG QUÉT CHÍNH (MAIN DISPATCHER LOOP)
  // =======================================================================================
  /**
   * Quét bảng UC05, lọc danh sách đến hạn và thực thi tạo Draft hoặc gửi Email
   * @param {"draft"|"send"} mode - "draft": tạo thư nháp trong Gmail; "send": gửi thật
   */
  /**
   * 3. CHẶNG 3: LUỒNG QUÉT CHÍNH (Lọc đúng PIC thao tác + CC TA In Charge)
   */
  static runProbationReview(mode = "draft") {
    const tConfig = PROBATION_REVIEW_CONFIG.PROGRESS_TRACKING;
    const timeZone = PROBATION_REVIEW_CONFIG.TIMEZONE;
    const todayKey = ProbationReviewUtil.formatDateKey(new Date(), timeZone);

    // 🔴 Lấy username/allocCode người đang bấm nút thao tác
    const currentUserEmail = Session.getActiveUser().getEmail().toLowerCase().trim();
    const currentAllocCode = currentUserEmail.split("@")[0].trim();
    const senderName = currentUserEmail
      ? currentUserEmail
          .split("@")[0]
          .replace(".", " ")
          .replace(/\b\w/g, (c) => c.toUpperCase())
      : "Talent Development Team";

    const { byEmail, byName } = this.loadOnboardingDataMap_();

    const ssUC05 = SpreadsheetApp.openById(tConfig.SPREADSHEET_ID);
    const sheetUC05 = ssUC05.getSheetByName(tConfig.TAB_NAME) || ssUC05.getSheets()[0];
    const logSheet = this.getOrCreateLogSheet_(ssUC05);
    const processedSet = this.loadProcessedSetFromLog_(logSheet);

    const dataUC05 = sheetUC05.getDataRange().getValues();
    if (dataUC05.length < tConfig.DATA_START_ROW_INDEX) {
      return { ok: true, processedCount: 0, message: "Sheet UC05 Contract Tracking không có dữ liệu." };
    }

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
      warnings: [],
      skipped: [],
      errors: [],
    };

    const newLogs = [];
    const relatedTaSet = new Set(); // Tập hợp TA In Charge để CC vào mail báo cáo
    const startRow = tConfig.DATA_START_ROW_INDEX - 1;

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

      if (!rawName) continue;

      // 🔴 BỘ LỌC PIC BẮT BUỘC: Bỏ qua nếu dòng này không thuộc quyền quản lý của người bấm nút
      const picEmail = ProbationReviewUtil.cleanEmail(rawPic);
      const picAllocCode = picEmail ? picEmail.split("@")[0].toLowerCase() : rawPic.toLowerCase().trim();

      if (picAllocCode !== currentAllocCode && picEmail !== currentUserEmail) {
        continue;
      }

      let onboardingInfo = null;
      if (rawEmail) onboardingInfo = byEmail.get(rawEmail);
      if (!onboardingInfo && rawName) onboardingInfo = byName.get(rawName.toLowerCase());

      if (onboardingInfo?.taInCharge) {
        relatedTaSet.add(onboardingInfo.taInCharge);
      }

      const targetEmail = (onboardingInfo && onboardingInfo.personalEmail) || rawEmail || (onboardingInfo ? onboardingInfo.workingEmail : "");

      if (!targetEmail || !ProbationReviewUtil.isValidEmail(targetEmail)) {
        summary.warnings.push({
          row: sheetRowNumberUC05,
          name: rawName,
          reason: "Không tìm thấy Email hợp lệ trên cả file UC05 và File Onboarding.",
        });
        continue;
      }

      const logKey = `${sheetRowNumberUC05}_${targetEmail.toLowerCase()}`;
      if (processedSet.has(logKey)) continue;

      const stopDate = onboardingInfo ? onboardingInfo.stopProbationDate : null;
      const isDue = ProbationReviewUtil.isDueReview(rawStartProcess, stopDate, PROBATION_REVIEW_CONFIG.DAYS_BEFORE_STOP, todayKey, timeZone);
      if (!isDue) continue;

      const okrFolderUrl = onboardingInfo ? onboardingInfo.okrFolderUrl : "";
      if (!okrFolderUrl) {
        summary.warnings.push({
          row: sheetRowNumberUC05,
          name: rawName,
          reason: "Thiếu Link Folder OKRs trên File Onboarding.",
        });
        continue;
      }

      try {
        const reviewTypeName = this.resolveReviewTypeLabel_(onboardingInfo?.employmentType, rawContractType);
        const reviewTypeLower = reviewTypeName.toLowerCase();

        const titleStr = onboardingInfo?.title || "";
        const squadTribeStr = onboardingInfo?.squadTribe || "";
        const metaSuffix = [titleStr, squadTribeStr].filter(Boolean).join(" | ");
        const subject = `${reviewTypeName} Review - ${rawName}${metaSuffix ? ` | ${metaSuffix}` : ""}`;

        // 🔴 CC: Line Manager + Dot Line Manager + People Team + PIC (bạn) + TA In Charge (từ UC03)
        const ccSet = new Set();
        if (onboardingInfo?.lineManager) ccSet.add(onboardingInfo.lineManager);
        if (onboardingInfo?.dotLineManager) ccSet.add(onboardingInfo.dotLineManager);
        ccSet.add(PROBATION_REVIEW_CONFIG.EMAIL.PEOPLE_TEAM_EMAIL);
        if (picEmail) ccSet.add(picEmail);
        if (onboardingInfo?.taInCharge) ccSet.add(onboardingInfo.taInCharge); // 👈 CC thêm TA In Charge

        ccSet.delete(targetEmail.toLowerCase());
        const ccList = Array.from(ccSet).join(", ");

        const stopProbationDisplay = onboardingInfo ? onboardingInfo.stopProbationDisplayDate : "[DD/MM/YYYY]";
        const consolidationDisplay = rawConsolidation ? ProbationReviewUtil.formatDisplayDate(rawConsolidation, timeZone) : stopProbationDisplay;

        const template = HtmlService.createTemplateFromFile("probation-review/probation-review-email");
        template.fullName = rawName;
        template.reviewTypeName = reviewTypeName;
        template.reviewTypeLower = reviewTypeLower;
        template.stopProbationDate = stopProbationDisplay;
        template.consolidationDate = consolidationDisplay;
        template.okrFolderUrl = okrFolderUrl;
        template.senderName = senderName;

        const htmlBody = template.evaluate().getContent();
        const mailOptions = { htmlBody: htmlBody, cc: ccList || undefined };

        if (mode === "send") {
          GmailApp.sendEmail(targetEmail, subject, "", mailOptions);
        } else {
          GmailApp.createDraft(targetEmail, subject, "", mailOptions);
        }

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
          folderUrl: okrFolderUrl,
        });

        newLogs.push([new Date(), sheetRowNumberUC05, rawName, targetEmail, reviewTypeName, "SUCCESS", mode.toUpperCase(), subject]);

        processedSet.add(logKey);
      } catch (err) {
        summary.errors.push(`Dòng ${sheetRowNumberUC05} (${rawName}): ${err.message}`);
      }
    }

    if (newLogs.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, newLogs.length, newLogs[0].length).setValues(newLogs);
    }

    // 🔴 Gửi Auto-Report trực tiếp cho người bấm chạy và CC các TA In Charge liên quan
    this.sendSummaryReportToTA_(summary, currentUserEmail, Array.from(relatedTaSet));

    summary.message = `Đã ${mode === "send" ? "gửi" : "tạo draft"} thành công cho ${summary.processedCount} nhân sự do bạn (${currentAllocCode}) phụ trách.`;
    return summary;
  }
}
