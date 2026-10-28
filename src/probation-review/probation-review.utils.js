/**
 * src/probation-review/probation-review.utils.js
 * Utility Pattern: Chứa toàn bộ hàm helper xử lý dữ liệu thuần cho Feature Probation Review
 */

class ProbationReviewUtil {
  /**
   * Trích xuất an toàn ID Google Drive / Folder từ URL hoặc chuỗi thô
   * Tuân thủ chuẩn an toàn: nhận diện /d/, /folders/, tham số ?id=, độ dài 25-50 ký tự, có biên chặn.
   * @param {string} url Chuỗi URL hoặc ID Drive
   * @return {string|null} ID hợp lệ hoặc null
   */
  static extractDriveId(url) {
    if (!url) return null;
    const raw = String(url).trim();
    const match = raw.match(
      /(?:\/(?:d|folders)\/|[?&]id=)([-\w]{25,50})(?:[/?&#]|$)/,
    );
    if (match) return match[1];

    // Fallback an toàn nếu dán trực tiếp mã ID thô
    if (/^[-\w]{25,50}$/.test(raw)) {
      return raw;
    }
    return null;
  }

  /**
   * Định dạng đối tượng Date hoặc chuỗi ngày về chuẩn key tra cứu: yyyy-MM-dd
   * @param {Date|string} dateVal Giá trị ngày
   * @param {string} timeZone Múi giờ hệ thống
   * @return {string} Chuỗi yyyy-MM-dd hoặc rỗng nếu không hợp lệ
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
   * Định dạng ngày phục vụ hiển thị trong email: dd/MM/yyyy
   * @param {Date|string} dateVal Giá trị ngày
   * @param {string} timeZone Múi giờ
   * @return {string}
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
   * Kiểm tra điều kiện kích hoạt gửi thông báo đánh giá thử việc:
   * - Điều kiện 1: Trùng ngày tại cột HR Ops - Start Process
   * - Điều kiện 2: Hoặc trước ngày Stop probation đúng targetDays (15 ngày)
   * @param {Date|string} startProcessDate Giá trị cột HR Ops - Start Process
   * @param {Date|string} stopProbationDate Giá trị cột Stop probation/intern date
   * @param {number} targetDays Số ngày đếm lùi (mặc định 15)
   * @param {string} todayKey Chuỗi yyyy-MM-dd của ngày hôm nay
   * @param {string} timeZone
   * @return {boolean}
   */
  static isDueReview(
    startProcessDate,
    stopProbationDate,
    targetDays = 15,
    todayKey,
    timeZone = "Asia/Ho_Chi_Minh",
  ) {
    if (!todayKey) return false;

    // 1. Kiểm tra theo ngày cột HR Ops - Start Process
    if (startProcessDate) {
      const startKey = this.formatDateKey(startProcessDate, timeZone);
      if (startKey && startKey === todayKey) return true;
    }

    // 2. Kiểm tra trước ngày kết thúc thử việc 15 ngày
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
        // Trừ lùi số ngày quy định
        stopDateObj.setDate(stopDateObj.getDate() - targetDays);
        const triggerKey = this.formatDateKey(stopDateObj, timeZone);
        if (triggerKey && triggerKey === todayKey) return true;
      }
    }

    return false;
  }

  /**
   * Chuẩn hóa email nhân sự hoặc PIC
   * Nếu chỉ có username thì tự động gắn domain mặc định
   * @param {string} raw Chuỗi email thô
   * @param {string} defaultDomain Domain mặc định (@kyanon.digital)
   * @return {string}
   */
  static cleanEmail(raw, defaultDomain = "@kyanon.digital") {
    const val = String(raw || "")
      .trim()
      .toLowerCase();
    if (!val) return "";
    return val.includes("@") ? val : `${val}${defaultDomain}`;
  }

  /**
   * Kiểm tra định dạng email hợp lệ
   * @param {string} email
   * @return {boolean}
   */
  static isValidEmail(email) {
    return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(email || "").trim());
  }
}
