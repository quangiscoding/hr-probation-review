/**
 * DỌN SẠCH TOÀN BỘ KHÓA DEVELOPER METADATA CỦA PROBATION REVIEW ĐỂ TEST LẠI
 */
function dev_resetAllProbationLocks() {
  const tConfig = PROBATION_REVIEW_CONFIG.PROGRESS_TRACKING;
  const ss = SpreadsheetApp.openById(tConfig.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(tConfig.TAB_NAME) || ss.getSheets()[0];

  const finder = sheet.createDeveloperMetadataFinder();
  const allMeta = finder.find();

  let removedCount = 0;
  allMeta.forEach((meta) => {
    if (meta.getKey().startsWith(PROBATION_REVIEW_CONFIG.METADATA_KEY_PREFIX)) {
      meta.remove();
      removedCount++;
    }
  });

  Logger.log(`🧹 Đã xóa thành công ${removedCount} khóa Developer Metadata của Probation Review trên UC05.`);
  return removedCount;
}
