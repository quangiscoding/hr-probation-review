/**
 * src/probation-review/probation-review.utils.js
 * Utility Pattern: Các hàm xử lý dữ liệu thuần cho Feature Probation Review
 */

class ProbationReviewUtil {
  /**
   * Trích xuất an toàn Drive/Folder ID từ URL hoặc chuỗi thô
   * Nhận diện /d/, /folders/, tham số ?id=, độ dài 25-50 ký tự, có biên chặn [/?&#]|$
   */
  static extractDriveId(url) {
    if (!url) return null;
    const raw = String(url).trim();
    const match = raw.match(
      /(?:\/(?:d|folders)\/|[?&]id=)([-\w]{25,50})(?:[/?&#]|$)/,
    );
    if (match) return match[1];

    if (/^[-\w]{25,50}$/.test(raw)) {
      return raw;
    }
    return null;
  }

  /**
   * Format Date sang chuỗi yyyy-MM-dd phục vụ đối soát
   */
  static formatDateKey(dateVal, timeZone = "Asia/Ho_Chi_Minh") {
    if (!dateVal) return "";
    if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
      return Utilities.formatDate(dateVal, timeZone, "yyyy-MM-dd");
    }
    const str = String(dateVal).trim();
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
      return str.substring(0, 10);
    }
    const parsed = Date.parse(str);
    if (!isNaN(parsed)) {
      return Utilities.formatDate(new Date(parsed), timeZone, "yyyy-MM-dd");
    }
    return "";
  }

  /**
   * Format Date hiển thị thân thiện trong email (dd/MM/yyyy)
   */
  static formatDisplayDate(dateVal, timeZone = "Asia/Ho_Chi_Minh") {
    if (!dateVal) return "";
    if (dateVal instanceof Date && !isNaN(dateVal.getTime())) {
      return Utilities.formatDate(dateVal, timeZone, "dd/MM/yyyy");
    }
    const str = String(dateVal).trim();
    const parsed = Date.parse(str);
    if (!isNaN(parsed)) {
      return Utilities.formatDate(new Date(parsed), timeZone, "dd/MM/yyyy");
    }
    return str;
  }

  /**
   * Kiểm tra điều kiện đến hạn đánh giá thử việc:
   * 1. Khớp ngày tại cột HR Ops - Start Process
   * 2. Hoặc trước ngày Stop probation đúng targetDays (15 ngày)
   */
  static isDueReview(
    startProcessVal,
    stopProbationDate,
    targetDays = 15,
    todayKey,
    timeZone = "Asia/Ho_Chi_Minh",
  ) {
    if (!todayKey) return false;

    // 1. Kiểm tra ngày tại cột HR Ops - Start Process
    if (startProcessVal) {
      const startKey = this.formatDateKey(startProcessVal, timeZone);
      if (startKey && startKey === todayKey) return true;
    }

    // 2. Kiểm tra trước Stop probation đúng targetDays (15 ngày)
    if (stopProbationDate) {
      let stopDateObj = null;
      if (
        stopProbationDate instanceof Date &&
        !isNaN(stopProbationDate.getTime())
      ) {
        stopDateObj = new Date(stopProbationDate.getTime());
      } else {
        const parsed = Date.parse(String(stopProbationDate).trim());
        if (!isNaN(parsed)) stopDateObj = new Date(parsed);
      }

      if (stopDateObj) {
        stopDateObj.setDate(stopDateObj.getDate() - targetDays);
        const triggerKey = this.formatDateKey(stopDateObj, timeZone);
        if (triggerKey && triggerKey === todayKey) return true;
      }
    }

    return false;
  }

  /**
   * Chuẩn hóa email nhân sự hoặc PIC
   */
  static cleanEmail(raw, defaultDomain = "@kyanon.digital") {
    const val = String(raw || "")
      .trim()
      .toLowerCase();
    if (!val) return "";
    return val.includes("@") ? val : `${val}${defaultDomain}`;
  }

  /**
   * Kiểm tra tính hợp lệ của email
   */
  static isValidEmail(email) {
    return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(email || "").trim());
  }
}
