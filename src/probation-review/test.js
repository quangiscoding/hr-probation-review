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
