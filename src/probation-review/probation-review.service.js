/**
 * src/probation-review/probation-review.service.js
 * Service: Đọc dữ liệu liên Sheet, lọc đúng hạn và lưu lịch sử vào tab "Probation Review Logs"
 */
class ProbationReviewService {
  static resolveReviewTypeLabel_(empTypeRaw, contractTypeRaw) {
    const raw = `${empTypeRaw || ""} ${contractTypeRaw || ""}`.toLowerCase();
    if (raw.includes("intern")) return "Intern";
    if (raw.includes("apprentice")) return "Apprentice";
    return "Probation";
  }

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

  /**
   * 2. Lấy hoặc tạo Sheet lưu Log trên file UC05
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
   * 3. Tải các bản ghi đã gửi từ tab Log về Set để tra cứu tức thì O(1)
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

  /**
   * 4. Chạy luồng quét đối soát và gửi mail / tạo draft
   * @param {"draft"|"send"} mode
   */
  static runProbationReview(mode = "draft") {
    const tConfig = PROBATION_REVIEW_CONFIG.PROGRESS_TRACKING;
    const timeZone = PROBATION_REVIEW_CONFIG.TIMEZONE;
    const todayKey = ProbationReviewUtil.formatDateKey(new Date(), timeZone);
    const currentUserEmail = Session.getActiveUser().getEmail();
    const senderName = currentUserEmail
      ? currentUserEmail
          .split("@")[0]
          .replace(".", " ")
          .replace(/\b\w/g, (c) => c.toUpperCase())
      : "Talent Development Team";

    // 1. Nạp data Onboarding
    const { byEmail, byName } = this.loadOnboardingDataMap_();

    // 2. Mở file UC05 và lấy danh sách đã chạy từ Log Sheet
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

    const newLogs = [];
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

      let onboardingInfo = null;
      if (rawEmail) onboardingInfo = byEmail.get(rawEmail);
      if (!onboardingInfo && rawName) onboardingInfo = byName.get(rawName.toLowerCase());

      const targetEmail = (onboardingInfo && onboardingInfo.personalEmail) || rawEmail || (onboardingInfo ? onboardingInfo.workingEmail : "");
      if (!targetEmail || !ProbationReviewUtil.isValidEmail(targetEmail)) {
        summary.skipped.push(`Dòng ${sheetRowNumberUC05} (${rawName}): Không tìm thấy Email hợp lệ.`);
        continue;
      }

      // Kiểm tra xem dòng này và email này đã ghi log thành công chưa
      const logKey = `${sheetRowNumberUC05}_${targetEmail.toLowerCase()}`;
      if (processedSet.has(logKey)) {
        continue; // Bỏ qua nếu đã xử lý
      }

      // Kiểm tra mốc ngày
      const stopDate = onboardingInfo ? onboardingInfo.stopProbationDate : null;
      const isDue = ProbationReviewUtil.isDueReview(rawStartProcess, stopDate, PROBATION_REVIEW_CONFIG.DAYS_BEFORE_STOP, todayKey, timeZone);
      if (!isDue) continue;

      try {
        const reviewTypeName = this.resolveReviewTypeLabel_(onboardingInfo?.employmentType, rawContractType);
        const reviewTypeLower = reviewTypeName.toLowerCase();
        const titleStr = onboardingInfo?.title || "";
        const squadTribeStr = onboardingInfo?.squadTribe || "";
        const metaSuffix = [titleStr, squadTribeStr].filter(Boolean).join(" | ");
        const subject = `${reviewTypeName} Review - ${rawName}${metaSuffix ? ` | ${metaSuffix}` : ""}`;

        // Danh sách CC
        const tdEmail = ProbationReviewUtil.cleanEmail(rawPic);
        const ccSet = new Set();
        if (onboardingInfo?.lineManager) ccSet.add(onboardingInfo.lineManager);
        if (onboardingInfo?.dotLineManager) ccSet.add(onboardingInfo.dotLineManager);
        ccSet.add(PROBATION_REVIEW_CONFIG.EMAIL.PEOPLE_TEAM_EMAIL);
        if (tdEmail) ccSet.add(tdEmail);
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
        template.okrFolderUrl = onboardingInfo ? onboardingInfo.okrFolderUrl : "";
        template.senderName = senderName;

        const htmlBody = template.evaluate().getContent();
        const mailOptions = {
          htmlBody: htmlBody,
          cc: ccList || undefined,
        };

        if (mode === "send") {
          GmailApp.sendEmail(targetEmail, subject, "", mailOptions);
        } else {
          GmailApp.createDraft(targetEmail, subject, "", mailOptions);
        }

        // Tăng biến đếm và lưu trạng thái thành công
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

        // Chuẩn bị dòng log
        newLogs.push([new Date(), sheetRowNumberUC05, rawName, targetEmail, reviewTypeName, "SUCCESS", mode.toUpperCase(), subject]);
        processedSet.add(logKey);
      } catch (err) {
        summary.errors.push(`Dòng ${sheetRowNumberUC05} (${rawName}): ${err.message}`);
      }
    }

    // Ghi batch vào tab Log
    if (newLogs.length > 0) {
      logSheet.getRange(logSheet.getLastRow() + 1, 1, newLogs.length, newLogs[0].length).setValues(newLogs);
    }

    summary.message = `Đã ${mode === "send" ? "gửi" : "tạo draft"} thành công cho ${summary.processedCount} nhân sự trên UC05.`;
    return summary;
  }
}
