/**
 * src/probation-review/probation-review.test.js
 * Test Probe: Kiểm tra tính đúng đắn của Config & Mapping trước khi chạy thật
 */

function test_probationReviewProbe() {
  const result = {
    trackingSheetOk: false,
    onboardingSheetOk: false,
    trackingMatchedCols: {},
    onboardingMatchedCols: {},
    sampleTrackingRows: [],
    sampleOnboardingRows: [],
  };

  // 1. KIỂM TRA BẢNG TRACKING (UC05)
  try {
    const tConfig = PROBATION_REVIEW_CONFIG.PROGRESS_TRACKING;
    const tSs = SpreadsheetApp.openById(tConfig.SPREADSHEET_ID);
    const tSheet = tSs.getSheetByName(tConfig.TAB_NAME) || tSs.getSheets()[0];
    const tData = tSheet.getDataRange().getValues();

    if (tData.length >= tConfig.HEADER_ROW_INDEX) {
      result.trackingSheetOk = true;
      const tRawHeaders = tData[tConfig.HEADER_ROW_INDEX - 1];
      const tNormHeaders = tRawHeaders.map(normalizeHeaderKey);

      // Đối chiếu cột
      Object.keys(tConfig.COLS).forEach((key) => {
        const expectedCol = tConfig.COLS[key];
        const foundIdx = tNormHeaders.indexOf(expectedCol);
        result.trackingMatchedCols[key] = {
          expected: expectedCol,
          foundIndex: foundIdx,
          realHeader:
            foundIdx >= 0 ? tRawHeaders[foundIdx] : "❌ KHÔNG TÌM THẤY",
        };
      });

      // Lấy thử 3 dòng dữ liệu đầu tiên
      const startRow = tConfig.DATA_START_ROW_INDEX - 1;
      for (let i = startRow; i < Math.min(startRow + 3, tData.length); i++) {
        result.sampleTrackingRows.push({
          row: i + 1,
          code: tData[i][0],
          name: tData[i][1],
          startProcess: tData[i][7],
          email: tData[i][11],
        });
      }
    }
  } catch (err) {
    result.trackingError = err.message;
  }

  // 2. KIỂM TRA BẢNG ONBOARDING (UC03 DEV)
  try {
    const oConfig = PROBATION_REVIEW_CONFIG.ONBOARDING;
    const oSs = SpreadsheetApp.openById(oConfig.SPREADSHEET_ID);
    let oSheet = null;
    const sheets = oSs.getSheets();
    for (let s of sheets) {
      if (String(s.getSheetId()) === String(oConfig.TAB_GID)) {
        oSheet = s;
        break;
      }
    }
    if (!oSheet) oSheet = sheets[0];

    const oData = oSheet.getDataRange().getValues();
    if (oData.length >= oConfig.HEADER_ROW_INDEX) {
      result.onboardingSheetOk = true;
      const oRawHeaders = oData[oConfig.HEADER_ROW_INDEX - 1];
      const oNormHeaders = oRawHeaders.map(normalizeHeaderKey);

      // Đối chiếu cột
      Object.keys(oConfig.COLS).forEach((key) => {
        const expectedCol = oConfig.COLS[key];
        const foundIdx = oNormHeaders.indexOf(expectedCol);
        result.onboardingMatchedCols[key] = {
          expected: expectedCol,
          foundIndex: foundIdx,
          realHeader:
            foundIdx >= 0 ? oRawHeaders[foundIdx] : "❌ KHÔNG TÌM THẤY",
        };
      });

      // Lấy thử 3 dòng dữ liệu đầu tiên
      const oStartRow = oConfig.DATA_START_ROW_INDEX - 1;
      for (let i = oStartRow; i < Math.min(oStartRow + 3, oData.length); i++) {
        result.sampleOnboardingRows.push({
          row: i + 1,
          name: oData[i][1],
          stopProbation: oData[i][17],
        });
      }
    }
  } catch (err) {
    result.onboardingError = err.message;
  }

  Logger.log("=== KẾT QUẢ KIỂM TRA THỰC TẾ TRÊN 2 SHEET ===");
  Logger.log(JSON.stringify(result, null, 2));
  return result;
}

/**
 * Chạy thử nghiệm tạo Draft và in kết quả ra Logger
 */
function test_probationReviewRun() {
  const result = ProbationReviewService.runProbationReview("draft");
  Logger.log("=== KẾT QUẢ THỰC THI THỬ NGHIỆM ===");
  Logger.log(JSON.stringify(result, null, 2));
}

/**
 * Test độc lập cho ProbationReviewUtil (Unit Test)
 */
function test_checkProbationUtils() {
  const tests = [];

  // 1. Test trích xuất Drive Folder ID
  const testUrls = [
    {
      in: "https://drive.google.com/drive/folders/1a2B3c4D5e6F7g8H9i0JkLmNoPqRsTuVw",
      exp: "1a2B3c4D5e6F7g8H9i0JkLmNoPqRsTuVw",
    },
    {
      in: "https://drive.google.com/open?id=12eYdZzEDttN9Q01yC_p_CVTYISQ6zz12",
      exp: "12eYdZzEDttN9Q01yC_p_CVTYISQ6zz12",
    },
    {
      in: "18LTemcygQH-nxTgev1iGTs5wTSEFblUu",
      exp: "18LTemcygQH-nxTgev1iGTs5wTSEFblUu",
    },
    { in: "invalid_short_url", exp: null },
  ];
  testUrls.forEach((t) => {
    const res = ProbationReviewUtil.extractDriveId(t.in);
    tests.push({
      name: `extractDriveId: ${t.in.slice(0, 30)}...`,
      pass: res === t.exp,
      actual: res,
      expected: t.exp,
    });
  });

  // 2. Test parse các định dạng ngày thực tế từ 2 ảnh Sheet
  // Giả định năm 2026 theo dữ liệu thực tế
  const testDates = [
    { in: "23-Aug", expKey: "2026-08-23", expDisplay: "23/08/2026" },
    { in: "02-Sep", expKey: "2026-09-02", expDisplay: "02/09/2026" },
    { in: "31-Jul-2026", expKey: "2026-07-31", expDisplay: "31/07/2026" },
    { in: "14-Sep-2026", expKey: "2026-09-14", expDisplay: "14/09/2026" },
    {
      in: new Date("2026-08-23T00:00:00"),
      expKey: "2026-08-23",
      expDisplay: "23/08/2026",
    },
  ];
  testDates.forEach((t) => {
    const key = ProbationReviewUtil.formatDateKey(t.in);
    const disp = ProbationReviewUtil.formatDisplayDate(t.in);
    tests.push({
      name: `DateParse (${t.in})`,
      pass: key === t.expKey && disp === t.expDisplay,
      actual: { key, disp },
      expected: { expKey: t.expKey, expDisplay: t.expDisplay },
    });
  });

  // 3. Test logic isDueReview (Trước 15 ngày & Khớp HR Ops)
  // Kịch bản A: Ngày hôm nay là 2026-08-23, cột Start Process là "23-Aug" -> TRUE
  const testA = ProbationReviewUtil.isDueReview(
    "23-Aug",
    null,
    15,
    "2026-08-23",
  );
  tests.push({
    name: "isDue: Trùng HR Ops 23-Aug",
    pass: testA === true,
    actual: testA,
    expected: true,
  });

  // Kịch bản B: Ngày hết hạn thử việc là "31-Jul-2026". Trừ 15 ngày là "2026-07-16".
  // Nếu hôm nay là 2026-07-16 -> TRUE
  const testB = ProbationReviewUtil.isDueReview(
    null,
    "31-Jul-2026",
    15,
    "2026-07-16",
  );
  tests.push({
    name: "isDue: Stop probation 31-Jul-2026 lùi 15 ngày = 16-Jul",
    pass: testB === true,
    actual: testB,
    expected: true,
  });

  // Kịch bản C: Chưa đến hạn (Hôm nay 2026-07-10) -> FALSE
  const testC = ProbationReviewUtil.isDueReview(
    "23-Aug",
    "31-Jul-2026",
    15,
    "2026-07-10",
  );
  tests.push({
    name: "isDue: Chưa đến hạn",
    pass: testC === false,
    actual: testC,
    expected: false,
  });

  // 4. Test clean email
  const emailTest1 = ProbationReviewUtil.cleanEmail("suong.phan");
  const emailTest2 = ProbationReviewUtil.cleanEmail("abc@gmail.com");
  tests.push({
    name: "cleanEmail username",
    pass: emailTest1 === "suong.phan@kyanon.digital",
    actual: emailTest1,
    expected: "suong.phan@kyanon.digital",
  });
  tests.push({
    name: "cleanEmail full",
    pass: emailTest2 === "abc@gmail.com",
    actual: emailTest2,
    expected: "abc@gmail.com",
  });

  Logger.log("=== KẾT QUẢ TEST UTILS ===");
  tests.forEach((t, i) => {
    Logger.log(
      `${t.pass ? "✅ PASS" : "❌ FAIL"} [${i + 1}] ${t.name} -> Actual: ${JSON.stringify(t.actual)}`,
    );
  });

  const totalFailed = tests.filter((t) => !t.pass).length;
  Logger.log(
    `TỔNG KẾT: ${tests.length - totalFailed}/${tests.length} tests passed.`,
  );
  return { total: tests.length, failed: totalFailed, details: tests };
}
