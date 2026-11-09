/** Sheet-bound probation catch-up reminders; enable while signed in as HR. */
// Fixed workflow settings. Only the spreadsheet ID and form URL use Script Properties.
const ONBOARD = {
  sheetName: "Contract Tracking",
  headerRow: 1,
  bookingSheet: "Booking Link",
  bookingHeaderRow: 1,
  timeZone: "Asia/Ho_Chi_Minh",
  appointmentTitle: "1-Month Check-in with TD team",
  candidateEmailColumn: "Candidate Email",
  emailDomain: "kyanon.digital",
  reminderHour: 9,
  meetingCheckMinutes: 15,
  reminderDaysMin: 1,
  reminderDaysMax: 3,
  meetingDaysBefore: 3,
  meetingDaysAfter: 7,
};

function onboardConfig_() {
  const properties = PropertiesService.getScriptProperties().getProperties();
  const read = (key) => {
    const name = `ONBOARD_${key}`;
    const value = text_(properties[name]);
    if (!value) throw new Error(`Missing Script Property: ${name}. See the setup guide.`);
    return value;
  };
  const config = {
    ...ONBOARD,
    spreadsheetId: read("SPREADSHEET_ID"),
    hrFormUrl: read("HR_FORM_URL"),
  };
  if (!/^[\w-]{20,}$/.test(config.spreadsheetId)) throw new Error("Invalid ONBOARD_SPREADSHEET_ID: use the spreadsheet ID, not its URL.");
  if (!/^https:\/\/\S+$/.test(config.hrFormUrl)) throw new Error("Invalid ONBOARD_HR_FORM_URL: use an HTTPS URL.");
  return config;
}

const TRACKING = ["Reminder Email Time", "HR Review Email Time"];

// function onOpen() {
//   SpreadsheetApp.getUi()
//     .createMenu("Probation Catch-up")
//     .addItem("Enable", "installTriggers")
//     .addItem("Disable", "disableTriggers")
//     .addItem("Send Reminder", "sendReminders")
//     .addItem("Send Review", "sendHrEmails")
//     .addToUi();
// }

function installTriggers() {
  const config = onboardConfig_();
  validateHr_();
  const spreadsheet = onboardSpreadsheet_(config);
  bookingLinks_(spreadsheet, config);
  disableTriggers();
  ScriptApp.newTrigger("sendReminders").timeBased().atHour(config.reminderHour).everyDays(1).inTimezone(config.timeZone).create();
  ScriptApp.newTrigger("sendHrEmails").timeBased().everyMinutes(config.meetingCheckMinutes).create();
}

function disableTriggers() {
  // Remove the retired draft trigger too when automation is disabled or re-enabled.
  const handlers = ["sendReminders", "sendHrEmails", "draftReminders", "checkOnboardDates", "checkMeetings"];
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (handlers.includes(trigger.getHandlerFunction())) ScriptApp.deleteTrigger(trigger);
  });
}

function sendReminders() {
  runOnboard_("reminder");
}
function sendHrEmails() {
  runOnboard_("meeting");
}

// Keep existing installed triggers working until HR enables the new schedules.
function checkOnboardDates() {
  sendReminders();
}
function checkMeetings() {
  sendHrEmails();
}

function runOnboard_(mode) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return;
  try {
    const config = onboardConfig_();
    const now = new Date();
    const spreadsheet = onboardSpreadsheet_(config);
    const hrEmail = validateHr_();
    const calendars = new Map(); // Resolve each PIC's primary calendar once per run.
    let bookingLinks; // Read the lookup tab once, only when a reminder is due.
    const failures = [];
    spreadsheet.getSheets().forEach((sheet) => {
      if (config.sheetName !== sheet.getName()) return;
      const data = sheet.getDataRange().getValues();
      const header = config.headerRow - 1;
      const candidateHeaders = (data[header] || []).map(text_).map((s) => s.toLowerCase());
      if (!["full name", "1-month probation catchup"].every((name) => candidateHeaders.includes(name))) {
        failures.push(`${sheet.getName()}: row ${config.headerRow} must contain Full name and 1-month probation catchup`);
        return;
      }
      if (!candidateHeaders.includes(config.candidateEmailColumn.toLowerCase())) {
        failures.push(`${sheet.getName()}: add the ${config.candidateEmailColumn} header in row ${config.headerRow}`);
        return;
      }
      if (!candidateHeaders.includes("pic")) {
        failures.push(`${sheet.getName()}: add the PIC header in row ${config.headerRow}`);
        return;
      }
      const headers = data[header].map(text_);
      TRACKING.forEach((name) => {
        if (!headers.some((h) => h.toLowerCase() === name.toLowerCase())) {
          headers.push(name);
          sheet.getRange(header + 1, headers.length).setValue(name);
        }
      });
      const columns = Object.fromEntries(headers.map((h, i) => [h.toLowerCase(), i]));
      for (let r = header + 1; r < data.length; r++) {
        const row = data[r];
        const get = (name) => row[columns[name.toLowerCase()]];
        const set = (name, value) => {
          sheet.getRange(r + 1, columns[name.toLowerCase()] + 1).setValue(value);
          SpreadsheetApp.flush();
        };
        if (!get("1-month probation catchup")) continue;
        const email = text_(get(config.candidateEmailColumn)).toLowerCase();
        if (!email) continue;
        try {
          const day = onboardDay_(get("1-month probation catchup"), config);
          const name = text_(get("Full name"));
          if (!name) throw new Error("Missing candidate name.");
          const timeColumn = mode === "reminder" ? "Reminder Email Time" : "HR Review Email Time";
          if (get(timeColumn)) continue; // One reminder and one review email per row.
          if (!validEmail_(email)) throw new Error("Enter a valid Candidate Email.");
          if (mode === "reminder") {
            const today = dateKey_(now, config);
            if (!reminderDue_(day, today, config)) continue;
            if (text_(get("Candidate Reminder Sent For")) === day) continue; // Already emailed by the previous version.
            const pic = picEmail_(get("PIC"), config);
            if (!bookingLinks) bookingLinks = bookingLinks_(spreadsheet, config);
            const bookingUrl = bookingLinks.get(pic);
            if (!bookingUrl || !/^https:\/\/\S+$/.test(bookingUrl)) {
              throw new Error(`Booking Link: enter an HTTPS Link for ${pic}.`);
            }
            const subject = `Reminder: Please schedule your 1-month probation catch-up - ${day}`;
            const content = `This is a reminder to schedule your 1-month probation catch-up, which is scheduled for ${day}.

Please use the following booking link to choose a time that works best for you:

${bookingUrl}

How to book:

1. Open the link and select an available time slot.
2. Enter your name and use your email address: ${email}.
3. Complete the booking and save your confirmation details.

This 30–45 minute session is an opportunity to discuss your onboarding experience, work progress, challenges, and any questions you may have. Your feedback helps us support your journey at Kyanon.

By default, the meeting will be arranged via Google Meet. The TD team will adjust the format (online or offline) based on your preference and confirm the final meeting details after your booking.`;
            // Candidate, PIC and HR receive the guidance in a single send request.
            // BCC keeps HR's address out of the candidate's recipient list.
            send_({ to: email, cc: pic, bcc: hrEmail, subject, ...emailBody_(name, content) });
            set("Reminder Email Time", new Date());
          } else {
            // Limit matching to this probation catch-up cycle; no name-only matching.
            const start = calendarDay_(day, -config.meetingDaysBefore, config);
            const end = calendarDay_(day, config.meetingDaysAfter + 1, config);
            // Do not repeat review emails already delivered by the previous version.
            if (text_(get("HR Form Status")).toLowerCase() === "sent" || get("HR Follow-up Sent At")) continue;
            const pic = picEmail_(get("PIC"), config);
            if (!calendars.has(pic)) calendars.set(pic, CalendarApp.getCalendarById(pic));
            const calendar = calendars.get(pic);
            if (!calendar) {
              throw new Error(`Cannot access PIC calendar ${pic}. Add it to the running HR account's Google Calendar and grant access to event details.`);
            }
            const matches = calendar
              .getEvents(start, end)
              .filter(
                (event) =>
                  event.getTitle() === config.appointmentTitle &&
                  !event.isAllDayEvent() &&
                  event.getGuestList(true).some((guest) => guest.getEmail().toLowerCase() === email && guest.getGuestStatus() !== CalendarApp.GuestStatus.NO),
              );
            if (matches.length > 1) throw new Error("Multiple check-up meetings for this candidate. Remove or cancel the duplicate booking.");
            if (!matches.length) continue;
            const event = matches[0];
            const meetingStart = event.getStartTime(),
              meetingEnd = event.getEndTime();
            if (meetingEnd.getTime() > now.getTime()) continue;
            const meetingStartText = Utilities.formatDate(meetingStart, config.timeZone, "yyyy-MM-dd HH:mm");
            const meetingEndText = Utilities.formatDate(meetingEnd, config.timeZone, dateKey_(meetingStart, config) === dateKey_(meetingEnd, config) ? "HH:mm" : "yyyy-MM-dd HH:mm");
            const meetingZone = Utilities.formatDate(meetingStart, config.timeZone, "Z").replace(/^([+-]\d{2})(\d{2})$/, "UTC$1:$2");
            send_({
              to: pic,
              subject: `Reminder: Complete the HR review form - ${name} - ${day}`,
              ...emailBody_(
                "TA",
                `Please complete the HR review form for ${name} (${email}) to record the discussion, feedback, and agreed follow-up actions from the 1-month probation catch-up.

The meeting has reached its scheduled end time.
Catch-up date: ${meetingStartText}–${meetingEndText} (${meetingZone})

HR review form:
${config.hrFormUrl}`,
              ),
            });
            set("HR Review Email Time", new Date());
          }
        } catch (error) {
          failures.push(`${sheet.getName()}, row ${r + 1}: ${error.message}`);
        }
      }
    });
    if (failures.length) throw new Error(failures.join("\n"));
  } finally {
    lock.releaseLock();
  }
}

function text_(value) {
  return value == null ? "" : String(value).trim();
}
function validEmail_(value) {
  return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(value);
}
function onboardSpreadsheet_(config) {
  const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
  if (!spreadsheet.getSheets().some((sheet) => sheet.getName() === config.sheetName)) {
    throw new Error(`Candidate tab "${config.sheetName}" was not found in the configured spreadsheet.`);
  }
  return spreadsheet;
}
function bookingLinks_(spreadsheet, config) {
  const sheet = spreadsheet.getSheetByName(config.bookingSheet);
  if (!sheet) throw new Error(`Add the ${config.bookingSheet} tab with TA and Link headers in row ${config.bookingHeaderRow}.`);
  const rows = sheet.getDataRange().getValues();
  const headers = (rows[config.bookingHeaderRow - 1] || []).map((value) => text_(value).toLowerCase());
  const picColumn = headers.indexOf("ta");
  const linkColumn = headers.indexOf("link");
  if (picColumn < 0 || linkColumn < 0) throw new Error(`${config.bookingSheet} row ${config.bookingHeaderRow} must contain TA and Link.`);
  const links = new Map();
  rows.slice(config.bookingHeaderRow).forEach((row, index) => {
    if (!text_(row[picColumn])) return;
    const sourceRow = config.bookingHeaderRow + index + 1;
    let pic;
    try {
      pic = picEmail_(row[picColumn], config);
    } catch (error) {
      throw new Error(`${config.bookingSheet}, row ${sourceRow}, TA "${text_(row[picColumn])}": ${error.message.replace("PIC must be", "TA must be")}`);
    }
    if (links.has(pic)) throw new Error(`${config.bookingSheet}, row ${sourceRow}: duplicate TA ${pic}. Keep one row per TA.`);
    links.set(pic, text_(row[linkColumn]));
  });
  return links;
}
function picEmail_(value, config) {
  const raw = text_(value).toLowerCase();
  const email = raw.includes("@") ? raw : `${raw}@${config.emailDomain}`;
  const parts = email.split("@");
  if (parts.length !== 2 || !/^[a-z0-9]+(?:[._-][a-z0-9]+)*$/.test(parts[0]) || parts[1] !== config.emailDomain) {
    throw new Error(`PIC must be one username or @${config.emailDomain} email.`);
  }
  return email;
}
function dateKey_(date, config) {
  return Utilities.formatDate(date, config.timeZone, "yyyy-MM-dd");
}
function calendarDay_(day, offset, config) {
  const date = new Date(`${day}T00:00:00Z`);
  date.setUTCDate(date.getUTCDate() + offset);
  return Utilities.parseDate(date.toISOString().slice(0, 10), config.timeZone, "yyyy-MM-dd");
}
function onboardDay_(value, config = onboardConfig_()) {
  if (value instanceof Date && !isNaN(value.getTime())) return dateKey_(value, config);
  const day = text_(value);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(day) || isNaN(Date.parse(day)) || new Date(`${day}T00:00:00Z`).toISOString().slice(0, 10) !== day) {
    throw new Error("1-month probation catchup must be a Sheet date or YYYY-MM-DD; ambiguous text dates are rejected.");
  }
  return day;
}
function reminderDue_(day, today, config = onboardConfig_()) {
  const daysUntil = (Date.parse(day) - Date.parse(today)) / 86400000;
  return daysUntil >= config.reminderDaysMin && daysUntil <= config.reminderDaysMax;
}
function validateHr_() {
  const email = text_(Session.getEffectiveUser().getEmail());
  if (!validEmail_(email)) throw new Error("Cannot identify the HR account. HR must authorize and enable the script.");
  return email;
}
function escapeHtml_(value) {
  return String(value).replace(
    /[&<>"']/g,
    (character) =>
      ({
        "&": "&amp;",
        "<": "&lt;",
        ">": "&gt;",
        '"': "&quot;",
        "'": "&#39;",
      })[character],
  );
}
function emailBody_(name, content) {
  const recipient = text_(name).toUpperCase();
  const htmlContent = escapeHtml_(content)
    .replace(/https:\/\/[^\s]+/g, (url) => `<a href="${url}">${url}</a>`)
    .split(/\n\n+/)
    .map((paragraph) => {
      const lines = paragraph.split("\n");
      if (lines.every((line) => /^\d+\. /.test(line))) {
        return "<ol>" + lines.map((line) => "<li>" + line.replace(/^\d+\. /, "") + "</li>").join("") + "</ol>";
      }
      return "<p>" + paragraph.replace(/\n/g, "<br>") + "</p>";
    })
    .join("");
  return {
    body: `Dear ${recipient},\n\n${content}\n\nBest regards,\n\nTalent Development Team`,
    htmlBody: `<p>Dear <strong>${escapeHtml_(recipient)}</strong>,</p>${htmlContent}<p>Best regards,</p><p>Talent Development Team</p>`,
  };
}
function send_(message) {
  const recipients = [message.to, message.cc, message.bcc].filter(Boolean).reduce((count, emails) => count + emails.split(",").length, 0);
  if (MailApp.getRemainingDailyQuota() < recipients) throw new Error("Insufficient daily email quota.");
  MailApp.sendEmail(withSenderSignature_(message));
}

// Execution-local cache, keyed by sender so different accounts never share a signature.
const EMAIL_SIGNATURES_ = new Map();
function withSenderSignature_(message) {
  const sender = String(Session.getEffectiveUser().getEmail() || "")
    .trim()
    .toLowerCase();
  if (!sender) throw new Error("Cannot identify the sender for the Gmail signature.");
  if (!EMAIL_SIGNATURES_.has(sender)) {
    try {
      const settings = Gmail.Users.Settings.SendAs.get("me", sender, { fields: "signature" });
      EMAIL_SIGNATURES_.set(sender, String(settings.signature || ""));
    } catch (error) {
      throw new Error("Cannot read the Gmail signature. Update appsscript.json and run checkEmailSignature to authorize Gmail settings. " + error.message);
    }
  }
  const signature = EMAIL_SIGNATURES_.get(sender);
  if (!signature.trim()) return message;
  const plain = signature
    .replace(/<br\s*\/?\s*>/gi, "\n")
    .replace(/<\/(?:p|div|tr)>/gi, "\n")
    .replace(/<[^>]*>/g, "")
    .replace(
      /&(nbsp|amp|lt|gt|quot|#39);/g,
      (entity) =>
        ({
          "&nbsp;": " ",
          "&amp;": "&",
          "&lt;": "<",
          "&gt;": ">",
          "&quot;": '"',
          "&#39;": "'",
        })[entity],
    )
    .trim();
  return { ...message, body: message.body + (plain ? "\n\n" + plain : ""), htmlBody: message.htmlBody + '<div class="gmail_signature" data-smartmail="gmail_signature">' + signature + "</div>" };
}

// Run once from the editor after updating the manifest. Does not send any email.
function checkEmailSignature() {
  const preview = withSenderSignature_({ body: "", htmlBody: "" });
  if (!preview.htmlBody) throw new Error("No Gmail signature found. Set a default signature for new emails in Gmail Settings, then run this function again.");
  console.info("Gmail signature is available. No email was sent.");
}
