/**
 * src/main.js - Controller điều phối
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("Probation Catch-up")
    .addItem("Enable reminders and HR form emails", "installTriggers")
    .addItem("Enable reminder drafts", "installDraftTriggers")
    .addItem("Disable automation", "disableTriggers")
    .addItem("Draft reminders now", "draftReminders")
    .addItem("Send reminders now", "sendReminders")
    .addItem("Send HR review emails now", "sendHrEmails")
    .addToUi();
  ui.createMenu("🚀 Probation Review")
    .addItem("1. Chạy thủ công tạo Draft ngay bây giờ", "menuRunDraftNow")
    .addSeparator()
    .addItem("⏰ Bật tự động tạo Draft hàng ngày (Auto 8:00 AM)", "installDailyDraftTrigger")
    .addItem("🛑 Tắt tự động tạo Draft", "uninstallDailyDraftTrigger")
    .addToUi();
}

/**
 * 1. Chạy thủ công ngay trên giao diện
 */
function menuRunDraftNow() {
  const result = ProbationReviewService.runProbationReview("draft");
  SpreadsheetApp.getUi().alert("Kết quả quét", result.message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 2. Hàm chạy ngầm được gọi tự động bởi Time-driven Trigger
 */
function cron_autoDraftProbationReview() {
  Logger.log("⏰ Bắt đầu quét tự động tạo draft Probation Review...");
  const result = ProbationReviewService.runProbationReview("draft");
  Logger.log(`✅ Kết quả: ${result.message}`);
}

/**
 * 3. Bật Trigger tự động chạy ngầm mỗi ngày từ 8h - 9h sáng
 */
function installDailyDraftTrigger() {
  const triggerName = "cron_autoDraftProbationReview";

  // Xóa trigger cũ để tránh chạy trùng lặp
  const triggers = ScriptApp.getProjectTriggers();
  for (let t of triggers) {
    if (t.getHandlerFunction() === triggerName) {
      ScriptApp.deleteTrigger(t);
    }
  }

  // Cài đặt chạy mỗi ngày trong khung giờ từ 8h đến 9h sáng
  ScriptApp.newTrigger(triggerName).timeBased().everyDays(1).atHour(8).create();

  SpreadsheetApp.getUi().alert("Đã bật tự động!", "Hệ thống sẽ tự động quét và tạo bản nháp (Draft) vào khoảng 8h - 9h sáng mỗi ngày.", SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * 4. Tắt Trigger tự động
 */
function uninstallDailyDraftTrigger() {
  const triggerName = "cron_autoDraftProbationReview";
  const triggers = ScriptApp.getProjectTriggers();
  let count = 0;

  for (let t of triggers) {
    if (t.getHandlerFunction() === triggerName) {
      ScriptApp.deleteTrigger(t);
      count++;
    }
  }

  SpreadsheetApp.getUi().alert("Đã tắt tự động", `Đã hủy bỏ ${count} trình kích hoạt tự động.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

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
