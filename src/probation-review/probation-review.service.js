/**
 * src/probation-review/probation-review.service.js
 * Service Pattern: Batch reading RAM, đối soát liên Sheet, bốc file OKR
 * Khóa dòng bằng Developer Metadata trực tiếp trên Sheet UC05 (Contract Tracking)
 */

class ProbationReviewService {
  /**
   * Helper: Chuẩn hóa nhãn loại Review từ Employment Type
   * Ví dụ: "Intern" -> "Intern", "Probation" -> "Probation", "Apprentice" -> "Apprentice"
   */
  static resolveReviewTypeLabel_(empTypeRaw, contractTypeRaw) {
    const raw = `${empTypeRaw || ""} ${contractTypeRaw || ""}`.toLowerCase();
    if (raw.includes("intern")) return "Intern";
    if (raw.includes("apprentice")) return "Apprentice";
    return "Probation"; // Mặc định là Probation
  }

  /**
   * 1. Đọc Onboarding Sheet (UC03) về RAM Map tra cứu O(1)
   */
  static loadOnboardingDataMap_() {
    const oConfig = PROBATION_REVIEW_CONFIG.ONBOARDING;
    const ss = SpreadsheetApp.openById(oConfig.SPREADSHEET_ID);

    let sheet = null;
    const sheets = ss.getSheets();
    for (let s of sheets) {
      if (String(s.getSheetId()) === String(oConfig.TAB_GID)) {
        sheet = s;
        break;
      }
    }
    if (!sheet) sheet = sheets[0];

    const data = sheet.getDataRange().getValues();
    if (data.length < oConfig.DATA_START_ROW_INDEX) {
      return { byEmail: new Map(), byName: new Map() };
    }

    const headers = data[oConfig.HEADER_ROW_INDEX - 1].map(normalizeHeaderKey);
    const fullNameIdx = headers.indexOf(oConfig.COLS.FULL_NAME);
    const personalEmailIdx = headers.indexOf(oConfig.COLS.PERSONAL_EMAIL);
    const workingEmailIdx = headers.indexOf(oConfig.COLS.WORKING_EMAIL);
    const stopDateIdx = headers.indexOf(
      oConfig.COLS.STOP_PROBATION_INTERN_DATE,
    );
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
      const personalEmail = ProbationReviewUtil.cleanEmail(
        row[personalEmailIdx],
      );
      const workingEmail = ProbationReviewUtil.cleanEmail(row[workingEmailIdx]);
      const stopProbationDate = row[stopDateIdx];
      const okrFolderUrl = String(row[okrFolderIdx] || "").trim();
      const title = String(row[titleIdx] || "").trim();
      const squad = String(row[squadIdx] || "").trim();
      const tribe = String(row[tribeIdx] || "").trim();
      const lineManager = ProbationReviewUtil.cleanEmail(row[lmIdx]);
      const dotLineManager = ProbationReviewUtil.cleanEmail(row[dlmIdx]);
      const employmentType = String(row[empTypeIdx] || "").trim();

      // Xử lý chống trùng lặp tên nếu squad và tribe giống hệt nhau
      let squadTribe = "";
      if (squad && tribe) {
        squadTribe =
          squad.toLowerCase() === tribe.toLowerCase()
            ? squad
            : `${squad} / ${tribe}`;
      } else {
        squadTribe = squad || tribe || "";
      }

      const record = {
        fullName: fullName,
        personalEmail: personalEmail,
        workingEmail: workingEmail,
        targetEmail: personalEmail || workingEmail,
        stopProbationDate: stopProbationDate,
        stopProbationDisplayDate:
          ProbationReviewUtil.formatDisplayDate(stopProbationDate),
        okrFolderUrl: okrFolderUrl,
        title: title,
        squadTribe: squadTribe,
        lineManager: lineManager,
        dotLineManager: dotLineManager,
        employmentType: employmentType,
      };

      if (personalEmail) byEmail.set(personalEmail, record);
      if (workingEmail) byEmail.set(workingEmail, record);
      if (fullName) byName.set(fullName.toLowerCase(), record);
    }

    return { byEmail, byName };
  }

  /**
   * 2. Bốc file OKR từ Google Drive Folder
   */
  static getOkrAttachmentFromFolder_(folderUrl) {
    if (!folderUrl) return null;
    const folderId = ProbationReviewUtil.extractDriveId(folderUrl);
    if (!folderId) return null;

    try {
      const folder = DriveApp.getFolderById(folderId);
      const files = folder.getFiles();
      let targetFile = null;

      while (files.hasNext()) {
        const file = files.next();
        const name = file.getName().toLowerCase();
        if (name.includes("okr")) {
          targetFile = file;
          break;
        }
        if (!targetFile) targetFile = file;
      }

      if (!targetFile) return null;

      const mime = targetFile.getMimeType();
      if (mime === MimeType.GOOGLE_SHEETS || mime === MimeType.GOOGLE_DOCS) {
        return targetFile
          .getAs(MimeType.PDF)
          .setName(`[OKR]_${targetFile.getName()}.pdf`);
      }
      return targetFile.getBlob();
    } catch (e) {
      Logger.log("⚠️ Không lấy được file OKR: " + e.message);
      return null;
    }
  }

  /**
   * 3. Chạy luồng quét đối soát và gửi mail / tạo draft
   * @param {"draft"|"send"} mode
   */
  static runProbationReview(mode = "draft") {
    const tConfig = PROBATION_REVIEW_CONFIG.PROGRESS_TRACKING;
    const metaPrefix = PROBATION_REVIEW_CONFIG.METADATA_KEY_PREFIX;
    const timeZone = PROBATION_REVIEW_CONFIG.TIMEZONE;
    const todayKey = ProbationReviewUtil.formatDateKey(new Date(), timeZone);
    const currentUserEmail = Session.getActiveUser().getEmail();
    const senderName = currentUserEmail
      ? currentUserEmail
          .split("@")[0]
          .replace(".", " ")
          .replace(/\b\w/g, (c) => c.toUpperCase())
      : "Talent Development Team";

    // 1. Tải dữ liệu Onboarding vào RAM Map
    const { byEmail, byName } = this.loadOnboardingDataMap_();

    // 2. Mở bảng UC05 (Contract Tracking)
    const ssUC05 = SpreadsheetApp.openById(tConfig.SPREADSHEET_ID);
    const sheetUC05 =
      ssUC05.getSheetByName(tConfig.TAB_NAME) || ssUC05.getSheets()[0];
    const dataUC05 = sheetUC05.getDataRange().getValues();

    if (dataUC05.length < tConfig.DATA_START_ROW_INDEX) {
      return {
        ok: true,
        processedCount: 0,
        message: "Sheet UC05 Contract Tracking không có dữ liệu.",
      };
    }

    // 3. Tải Developer Metadata của UC05 về RAM Set O(1)
    const lockedMetadataSet = new Set();
    try {
      const allMetaUC05 = sheetUC05.getDeveloperMetadata();
      allMetaUC05.forEach((m) => {
        const k = m.getKey();
        const v = m.getValue();
        if (k) lockedMetadataSet.add(k);
        if (v) lockedMetadataSet.add(v);
      });
    } catch (errMeta) {
      Logger.log(
        "⚠️ Không tải được Metadata theo lô trên UC05: " + errMeta.message,
      );
    }

    const headers =
      dataUC05[tConfig.HEADER_ROW_INDEX - 1].map(normalizeHeaderKey);
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

    const startRow = tConfig.DATA_START_ROW_INDEX - 1;

    // 4. Quét từng dòng dữ liệu trên UC05
    for (let r = startRow; r < dataUC05.length; r++) {
      const row = dataUC05[r];
      const sheetRowNumberUC05 = r + 1;
      const rawName = String(row[nameIdx] || "").trim();
      const rawEmail = String(row[candidateEmailIdx] || "")
        .trim()
        .toLowerCase();
      const rawStartProcess = row[startProcessIdx];
      const rawConsolidation =
        consolidationIdx >= 0 ? row[consolidationIdx] : null;
      const rawPic = String(row[picIdx] || "").trim();
      const rawContractType = String(row[contractTypeIdx] || "").trim();

      if (!rawName) continue;

      // Tra cứu chéo sang Onboarding: ưu tiên Email nếu có, fallback sang Họ Tên
      let onboardingInfo = null;
      if (rawEmail) onboardingInfo = byEmail.get(rawEmail);
      if (!onboardingInfo && rawName)
        onboardingInfo = byName.get(rawName.toLowerCase());

      // Xác định email gửi: Ưu tiên Personal Email (hoặc Working Email)
      const targetEmail =
        (onboardingInfo && onboardingInfo.personalEmail) ||
        rawEmail ||
        (onboardingInfo ? onboardingInfo.workingEmail : "");

      if (!targetEmail || !ProbationReviewUtil.isValidEmail(targetEmail)) {
        summary.skipped.push(
          `Dòng ${sheetRowNumberUC05} (${rawName}): Không tìm thấy Email hợp lệ.`,
        );
        continue;
      }

      // 🔒 Fast Check Lock Metadata trên RAM
      const lockKey = `${metaPrefix}_${sheetRowNumberUC05}_${targetEmail.toLowerCase()}`;
      if (lockedMetadataSet.has(lockKey)) {
        continue;
      }

      // Kiểm tra mốc thời gian đến hạn (HR Ops - Start Process hoặc trước 15 ngày)
      const stopDate = onboardingInfo ? onboardingInfo.stopProbationDate : null;
      const isDue = ProbationReviewUtil.isDueReview(
        rawStartProcess,
        stopDate,
        PROBATION_REVIEW_CONFIG.DAYS_BEFORE_STOP,
        todayKey,
        timeZone,
      );

      if (!isDue) continue;

      try {
        // Lấy attachment file OKR từ thư mục Drive
        const attachments = [];
        let okrBlob = null;
        if (onboardingInfo && onboardingInfo.okrFolderUrl) {
          okrBlob = this.getOkrAttachmentFromFolder_(
            onboardingInfo.okrFolderUrl,
          );
          if (okrBlob) attachments.push(okrBlob);
        }

        // Đọc Employment Type từ file Onboarding và phân loại đúng nhãn Review
        const reviewTypeName = this.resolveReviewTypeLabel_(
          onboardingInfo?.employmentType,
          rawContractType,
        );
        const reviewTypeLower = reviewTypeName.toLowerCase();

        // Cấu trúc Subject: [ReviewType] Review - [Employee Name] | [Title] | [Squad/Tribe]
        const titleStr = onboardingInfo?.title || "";
        const squadTribeStr = onboardingInfo?.squadTribe || "";
        const metaSuffix = [titleStr, squadTribeStr]
          .filter(Boolean)
          .join(" | ");
        const subject = `${reviewTypeName} Review - ${rawName}${metaSuffix ? ` | ${metaSuffix}` : ""}`;

        // Danh sách CC: Line Manager, Dot Line Manager, people@kyanon.digital, TD in charge
        const tdEmail = ProbationReviewUtil.cleanEmail(rawPic);
        const ccSet = new Set();
        if (onboardingInfo?.lineManager) ccSet.add(onboardingInfo.lineManager);
        if (onboardingInfo?.dotLineManager)
          ccSet.add(onboardingInfo.dotLineManager);
        ccSet.add(PROBATION_REVIEW_CONFIG.EMAIL.PEOPLE_TEAM_EMAIL);
        if (tdEmail) ccSet.add(tdEmail);

        // Loại bỏ chính email người nhận khỏi danh sách CC
        ccSet.delete(targetEmail.toLowerCase());
        const ccList = Array.from(ccSet).join(", ");

        // Format mốc thời gian:
        // DD/MM/YYYY thứ 1: Ngày kết thúc thử việc (từ Onboarding UC03)
        const stopProbationDisplay = onboardingInfo
          ? onboardingInfo.stopProbationDisplayDate
          : "[DD/MM/YYYY]";

        // DD/MM/YYYY thứ 2: Hạn chót từ cột HR Ops Consolidation (từ UC05)
        const consolidationDisplay = rawConsolidation
          ? ProbationReviewUtil.formatDisplayDate(rawConsolidation, timeZone)
          : stopProbationDisplay;

        // Render HTML Email Template
        const template = HtmlService.createTemplateFromFile(
          "probation-review/probation-review-email",
        );
        template.fullName = rawName;
        template.reviewTypeName = reviewTypeName;
        template.reviewTypeLower = reviewTypeLower;
        template.stopProbationDate = stopProbationDisplay;
        template.consolidationDate = consolidationDisplay;
        template.okrFolderUrl = onboardingInfo
          ? onboardingInfo.okrFolderUrl
          : "";
        template.senderName = senderName;

        const htmlBody = template.evaluate().getContent();

        const mailOptions = {
          htmlBody: htmlBody,
          cc: ccList || undefined,
          attachments: attachments,
        };

        if (mode === "send") {
          GmailApp.sendEmail(targetEmail, subject, "", mailOptions);
        } else {
          GmailApp.createDraft(targetEmail, subject, "", mailOptions);
        }

        // 🔒 KHÓA DÒNG BẰNG DEVELOPER METADATA TRÊN SHEET UC05 (CÚ PHÁP ENTIRE ROW CHUẨN)
        sheetUC05
          .getRange(`${sheetRowNumberUC05}:${sheetRowNumberUC05}`)
          .addDeveloperMetadata(lockKey, new Date().toISOString());
        lockedMetadataSet.add(lockKey);

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
          hasOkr: !!okrBlob,
        });
      } catch (err) {
        summary.errors.push(
          `Dòng ${sheetRowNumberUC05} (${rawName}): ${err.message}`,
        );
      }
    }

    summary.message = `Đã ${mode === "send" ? "gửi" : "tạo draft"} thành công cho ${summary.processedCount} nhân sự trên UC05.`;
    return summary;
  }
}
