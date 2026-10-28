/**
 * src/probation-review/probation-review.config.js
 * Cấu hình độc lập cho module Probation Review
 */
function normalizeHeaderKey(str) {
  if (!str) return "";
  return String(str)
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^a-z0-9]/g, "");
}

const PROBATION_REVIEW_CONFIG = {
  DAYS_BEFORE_STOP: 15,
  TIMEZONE: "Asia/Ho_Chi_Minh",
  METADATA_KEY_PREFIX: "PROBATION_REVIEW",

  // 1. SHEET UC05 (Bảng điều khiển: Progress Tracking / Contract Tracking)
  PROGRESS_TRACKING: {
    SPREADSHEET_ID: "1oAzOciyfA_b91aPz14jN-O-crGLyqoegBGFwurbggLA",
    TAB_NAME: "Contract Tracking",
    TAB_GID: "0",
    HEADER_ROW_INDEX: 1,
    DATA_START_ROW_INDEX: 2,
    COLS: {
      CODE: normalizeHeaderKey("Code"),
      FULL_NAME: normalizeHeaderKey("Full name"),
      ONBOARD_DATE: normalizeHeaderKey("Onboard Date"),
      UNIT_SQUAD_NAME: normalizeHeaderKey("Unit / Squad Name"),
      TYPE_OF_CONTRACT: normalizeHeaderKey("Type of contract"),
      PIC: normalizeHeaderKey("PIC"),
      ONE_MONTH_PROBATION_CATCHUP: normalizeHeaderKey(
        "1-month probation catchup",
      ),
      HR_OPS_START_PROCESS: normalizeHeaderKey("HR Ops - Start Process"),
      SL_LM_CONFIRM: normalizeHeaderKey("SL/LM Confirm"),
      HR_OPS_CONSOLIDATION: normalizeHeaderKey("HR Ops Consolidation"),
      HEAD_OF_OP_PROPOSAL: normalizeHeaderKey("Head of O&P Proposal"),
      CANDIDATE_EMAIL: normalizeHeaderKey("Candidate Email"),
    },
  },

  // 2. SHEET ONBOARDING (Bảng hồ sơ gốc: Hera | UC03 Dev)
  ONBOARDING: {
    SPREADSHEET_ID: "1rnowvvOuOJ_5i9zDrz_1-TigNq_EZMY02nnOGbGQgTE",
    TAB_GID: "335034815",
    HEADER_ROW_INDEX: 1,
    DATA_START_ROW_INDEX: 3,
    COLS: {
      FULL_NAME: normalizeHeaderKey("Full Name"),
      ACCENTLESS_FULL_NAME: normalizeHeaderKey("Accentless Full Name"),
      LEVEL: normalizeHeaderKey("Level"),
      SR: normalizeHeaderKey("Sr"),
      EMPLOYMENT_TYPE: normalizeHeaderKey("Employment Type"),
      SPECIALIST: normalizeHeaderKey("Specialist"),
      TITLE: normalizeHeaderKey("Title"),
      SQUAD_UNIT: normalizeHeaderKey("Squad/ Unit"),
      TRIBE_DIVISION: normalizeHeaderKey("Tribe/ Division"),
      LINE_MANAGER: normalizeHeaderKey("Line Manager"),
      DOT_LINE_MANAGER: normalizeHeaderKey("Dot Line Manager"),
      CHAPTER: normalizeHeaderKey("Chapter"),
      CHAPTER_LEAD: normalizeHeaderKey("Chapter Lead"),
      DATE_OF_OFFER: normalizeHeaderKey("Date of Offer"),
      TA_IN_CHARGE: normalizeHeaderKey("TA In Charge"),
      DATE_OF_ONBOARD: normalizeHeaderKey("Date of Onboard"),
      STOP_PROBATION_INTERN_DATE: normalizeHeaderKey(
        "Stop probation/ intern date",
      ),
      OFFER_STATUS: normalizeHeaderKey("Offer Status"),
      DEVICE_REQUEST: normalizeHeaderKey("Device Request from newcomer"),
      ONBOARDING_TYPE: normalizeHeaderKey("Onboarding Type"),
      WORKING_EMAIL: normalizeHeaderKey("Working Email"),
      GOOGLE_CHAT: normalizeHeaderKey("Google Chat"),
      OFFER_LETTER_APPROVAL: normalizeHeaderKey("Offer Letter Approval"),
      OFFER_LETTER_ACCEPTANCE: normalizeHeaderKey("Offer Letter Acceptance"),
      CV: normalizeHeaderKey("CV"),
      ENTRANCE_ASSESSMENT: normalizeHeaderKey("Entrance Assessment"),
      MASTER_CV_FOLDER: normalizeHeaderKey("Master CV (Folder)"),
      PHONE_NUMBER: normalizeHeaderKey("Phone Number"),
      PERSONAL_EMAIL: normalizeHeaderKey("Personal Email"),
      LINK_FOLDER_OKRS: normalizeHeaderKey("Link Folder OKRs"),
      REMARKS: normalizeHeaderKey("Remarks (update by Atlas)"),
      ALLOC_CODE: normalizeHeaderKey("Alloc Code (update by Atlas)"),
    },
  },

  EMAIL: {
    PEOPLE_TEAM_EMAIL: "people@kyanon.digital",
    DEFAULT_DOMAIN: "@kyanon.digital",
  },
};
