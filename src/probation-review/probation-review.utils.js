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
   */
  static extractDriveId(url) {
    if (!url) return null;
    const raw = String(url).trim();
    const match = raw.match(/(?:\/(?:d|folders)\/|[?&]id=)([-\w]{25,50})(?:[/?&#]|$)/);
    if (match) return match[1];
    if (/^[-\w]{25,50}$/.test(raw)) {
      return raw;
    }
    return null;
  }

  /**
   * Trích xuất an toàn URL Google Drive/Folder từ ô bảng tính
   */
  static extractCellUrl(cellValue, cellRichText) {
    if (cellRichText && typeof cellRichText.getLinkUrl === "function") {
      const url = cellRichText.getLinkUrl();
      if (url && /^https?:\/\//i.test(url.trim())) {
        return url.trim();
      }
    }
    const raw = String(cellValue || "").trim();
    if (!raw) return "";
    if (/^https?:\/\//i.test(raw)) {
      return raw;
    }
    if (/^[-\w]{25,50}$/.test(raw)) {
      return `https://drive.google.com/drive/folders/${raw}`;
    }
    return "";
  }

  /**
   * Parse an toàn mọi định dạng ngày trên Google Sheet
   */
  static parseToDateObject(val, defaultYear = new Date().getFullYear()) {
    if (!val) return null;
    if (val instanceof Date && !isNaN(val.getTime())) {
      return new Date(val.getTime());
    }
    const str = String(val).trim();
    if (!str) return null;

    // Dạng "31-Jul-2026"
    const matchFull = str.match(/^(\d{1,2})[-/]([a-zA-Z]{3})[-/](\d{4})$/);
    if (matchFull) {
      const d = matchFull[1].padStart(2, "0");
      const m = this.MONTH_MAP[matchFull[2].toLowerCase()];
      const y = matchFull[3];
      if (m) return new Date(`${y}-${m}-${d}T00:00:00`);
    }

    // Dạng "23-Aug"
    const matchShort = str.match(/^(\d{1,2})[-/]([a-zA-Z]{3})$/);
    if (matchShort) {
      const d = matchShort[1].padStart(2, "0");
      const m = this.MONTH_MAP[matchShort[2].toLowerCase()];
      if (m) return new Date(`${defaultYear}-${m}-${d}T00:00:00`);
    }

    // Dạng "dd/MM/yyyy"
    const matchVn = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
    if (matchVn) {
      const d = matchVn[1].padStart(2, "0");
      const m = matchVn[2].padStart(2, "0");
      const y = matchVn[3];
      return new Date(`${y}-${m}-${d}T00:00:00`);
    }

    // Dạng "yyyy-MM-dd"
    if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
      return new Date(str.substring(0, 10) + "T00:00:00");
    }

    const parsed = Date.parse(str);
    if (!isNaN(parsed)) return new Date(parsed);
    return null;
  }

  static formatDateKey(dateVal, timeZone = "Asia/Ho_Chi_Minh") {
    const d = this.parseToDateObject(dateVal);
    if (!d) return "";
    return Utilities.formatDate(d, timeZone, "yyyy-MM-dd");
  }

  static formatDisplayDate(dateVal, timeZone = "Asia/Ho_Chi_Minh") {
    const d = this.parseToDateObject(dateVal);
    if (!d) return String(dateVal || "");
    return Utilities.formatDate(d, timeZone, "dd/MM/yyyy");
  }

  /**
   * KIỂM TRA ĐIỀU KIỆN ĐẾN HẠN ĐÁNH GIÁ (Chỉ khớp đúng ngày hôm nay)
   * 1. Nếu có mốc HR Ops - Start Process: ngày đó phải bằng đúng hôm nay.
   * 2. Nếu không có mốc Start Process: Stop probation trừ 15 ngày phải bằng đúng hôm nay.
   */
  static isDueReview(startProcessVal, stopProbationDate, targetDays = 15, todayKey, timeZone = "Asia/Ho_Chi_Minh") {
    if (!todayKey) return false;
    const todayObj = this.parseToDateObject(todayKey);
    if (!todayObj) return false;
    todayObj.setHours(0, 0, 0, 0);

    // Ưu tiên 1: Cột HR Ops - Start Process khớp đúng ngày hôm nay
    if (startProcessVal) {
      const startProcessObj = this.parseToDateObject(startProcessVal);
      if (startProcessObj) {
        startProcessObj.setHours(0, 0, 0, 0);
        return startProcessObj.getTime() === todayObj.getTime();
      }
    }

    // Ưu tiên 2: Mốc Stop probation lùi đúng 15 ngày
    if (stopProbationDate) {
      const stopDateObj = this.parseToDateObject(stopProbationDate);
      if (stopDateObj) {
        stopDateObj.setHours(0, 0, 0, 0);
        const diffDays = Math.round((stopDateObj.getTime() - todayObj.getTime()) / (1000 * 60 * 60 * 24));
        return diffDays === targetDays;
      }
    }

    return false;
  }

  static cleanEmail(raw, defaultDomain = "@kyanon.digital") {
    const val = String(raw || "")
      .trim()
      .toLowerCase();
    if (!val) return "";
    return val.includes("@") ? val : `${val}${defaultDomain}`;
  }

  static isValidEmail(email) {
    return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(email || "").trim());
  }
}
