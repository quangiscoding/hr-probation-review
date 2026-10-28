/**
 * src/probation-review/probation-review.utils.js
 * Utility Pattern: Các hàm xử lý dữ liệu thuần cho Feature Probation Review
 */

class ProbationReviewUtil {
  /**
   * Bảng ánh xạ tháng viết tắt tiếng Anh (Aug, Sep, Oct, Jul...)
   */
  static get MONTH_MAP() {
    return {
      jan: "01",
      feb: "02",
      mar: "03",
      apr: "04",
      may: "05",
      jun: "06",
      jul: "07",
      aug: "08",
      sep: "09",
      oct: "10",
      nov: "11",
      dec: "12",
    };
  }

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
   * Parse an toàn mọi định dạng ngày trên Google Sheet:
   * - Date object nguyên bản từ Google Sheet
   * - "31-Jul-2026" / "14-Sep-2026" (dạng dd-MMM-yyyy trên Sheet Onboarding UC03)
   * - "23-Aug" / "02-Sep" (dạng dd-MMM trên Sheet UC05 - tự động bù defaultYear)
   * - "yyyy-MM-dd" (ISO)
   * - "dd/MM/yyyy" (Việt Nam)
   */
  static parseToDateObject(val, defaultYear = new Date().getFullYear()) {
    if (!val) return null;

    // 1. Nếu cell vốn đã là Date object
    if (val instanceof Date && !isNaN(val.getTime())) {
      return new Date(val.getTime());
    }

    const str = String(val).trim();
    if (!str) return null;

    // 2. Dạng "31-Jul-2026" (như ảnh Sheet Onboarding UC03)
    const matchFull = str.match(/^(\d{1,2})[-/]([a-zA-Z]{3})[-/](\d{4})$/);
    if (matchFull) {
      const d = matchFull[1].padStart(2, "0");
      const m = this.MONTH_MAP[matchFull[2].toLowerCase()];
      const y = matchFull[3];
      if (m) return new Date(`${y}-${m}-${d}T00:00:00`);
    }

    // 3. Dạng "23-Aug" (như ảnh Sheet UC05 thiếu năm)
    const matchShort = str.match(/^(\d{1,2})[-/]([a-zA-Z]{3})$/);
    if (matchShort) {
      const d = matchShort[1].padStart(2, "0");
      const m = this.MONTH_MAP[matchShort[2].toLowerCase()];
      if (m) return new Date(`${defaultYear}-${m}-${d}T00:00:00`);
    }

    // 4. Dạng "dd/MM/yyyy"
    const matchVn = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (matchVn) {
      const d = matchVn[1].padStart(2, "0");
      const m = matchVn[2].padStart(2, "0");
      const y = matchVn[3];
      return new Date(`${y}-${m}-${d}T00:00:00`);
    }

    // 5. Dạng ISO "yyyy-MM-dd"
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
      return new Date(str.substring(0, 10) + "T00:00:00");
    }

    // 6. Fallback Date.parse
    const parsed = Date.parse(str);
    if (!isNaN(parsed)) return new Date(parsed);

    return null;
  }

  /**
   * Format Date sang chuỗi yyyy-MM-dd phục vụ đối soát
   */
  static formatDateKey(dateVal, timeZone = "Asia/Ho_Chi_Minh") {
    const d = this.parseToDateObject(dateVal);
    if (!d) return "";
    return Utilities.formatDate(d, timeZone, "yyyy-MM-dd");
  }

  /**
   * Format Date hiển thị thân thiện trong email (dd/MM/yyyy)
   */
  static formatDisplayDate(dateVal, timeZone = "Asia/Ho_Chi_Minh") {
    const d = this.parseToDateObject(dateVal);
    if (!d) return String(dateVal || "");
    return Utilities.formatDate(d, timeZone, "dd/MM/yyyy");
  }

  /**
   * Kiểm tra điều kiện đến hạn đánh giá thử việc (Dual-Gate):
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
      const stopDateObj = this.parseToDateObject(stopProbationDate);
      if (stopDateObj) {
        stopDateObj.setDate(stopDateObj.getDate() - targetDays);
        const triggerKey = Utilities.formatDate(
          stopDateObj,
          timeZone,
          "yyyy-MM-dd",
        );
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
