/**
 * src/main.js - Controller điều phối
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Probation Catch-up")
    .addItem("Enable reminders and HR form emails", "installTriggers")
    .addItem("Enable reminder drafts", "installDraftTriggers")
    .addItem("Disable automation", "disableTriggers")
    .addItem("Draft reminders now", "draftReminders")
    .addItem("Send reminders now", "sendReminders")
    .addItem("Send HR review emails now", "sendHrEmails")
    .addToUi();
  SpreadsheetApp.getUi()
    .createMenu("🚀 Probation Review")
    .addItem("1. Chạy thủ công tạo Draft ngay bây giờ", "menuRunDraftNow")
    .addSeparator()
    .addItem(
      "⏰ Bật tự động tạo Draft hàng ngày (Auto 8:00 AM)",
      "installDailyDraftTrigger",
    )
    .addItem("🛑 Tắt tự động tạo Draft", "uninstallDailyDraftTrigger")
    .addToUi();
}

/**
 * 1. Chạy thủ công ngay trên giao diện
 */
function menuRunDraftNow() {
  const result = ProbationReviewService.runProbationReview("draft");
  SpreadsheetApp.getUi().alert(
    "Kết quả quét",
    result.message,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
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

  // Xóa trigger cũ để tránh chạy lặp
  const triggers = ScriptApp.getProjectTriggers();
  for (let t of triggers) {
    if (t.getHandlerFunction() === triggerName) {
      ScriptApp.deleteTrigger(t);
    }
  }

  // Cài đặt Giờ: 8h, Phút: xấp xỉ 30 phút
  ScriptApp.newTrigger(triggerName)
    .timeBased()
    .everyDays(1)
    .atHour(13)
    .nearMinute(51) // 👈 Thêm phút tại đây (0, 15, 30, 45,...)
    .create();

  SpreadsheetApp.getUi().alert(
    "Đã bật tự động!",
    "Hệ thống sẽ tự động quét và tạo bản nháp (Draft) vào khoảng 08:30 sáng mỗi ngày.",
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
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

  SpreadsheetApp.getUi().alert(
    "Đã tắt tự động",
    `Đã hủy bỏ ${count} trình kích hoạt tự động.`,
    SpreadsheetApp.getUi().ButtonSet.OK,
  );
}
