/** Sheet-bound probation catch-up reminders; enable while signed in as HR. */
const ONBOARD = {
  spreadsheetId: "1oAzOciyfA_b91aPz14jN-O-crGLyqoegBGFwurbggLA",
  sheetNames: ["Contract Tracking"], // Only this candidate tab is processed.
  timeZone: "Asia/Ho_Chi_Minh",
  bookingSheet: "Booking Link", // Row 1 headers: PIC and Link.
  appointmentTitle: "1-Month Check-in with TD team", // Exact booked-event title.
  hrFormUrl: "https://forms.gle/js3uUfqeBgyTte3a6",
  candidateEmailColumn: "Candidate Email", // Email used to book the check-up.
};

const TRACKING = [
  "Reminder Email Time",
  "HR Review Email Time",
  "Reminder Draft Time",
];

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
}

function installTriggers() {
  if (!ONBOARD.appointmentTitle || !ONBOARD.hrFormUrl) {
    throw new Error("Configure appointmentTitle and hrFormUrl first.");
  }
  validateHr_();
  const calendar = CalendarApp.getDefaultCalendar();
  if (!calendar) throw new Error("Cannot access the HR booking calendar.");
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet)
    throw new Error("Enable this script from the candidate spreadsheet.");
  bookingLinks_(spreadsheet);
  PropertiesService.getUserProperties().setProperty(
    "spreadsheetId",
    spreadsheet.getId(),
  );
  disableTriggers();
  ScriptApp.newTrigger("sendReminders")
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .inTimezone(ONBOARD.timeZone)
    .create();
  ScriptApp.newTrigger("sendHrEmails").timeBased().everyMinutes(15).create();
}

function installDraftTriggers() {
  validateHr_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet)
    throw new Error("Enable this script from the candidate spreadsheet.");
  bookingLinks_(spreadsheet);
  PropertiesService.getUserProperties().setProperty(
    "spreadsheetId",
    spreadsheet.getId(),
  );
  disableTriggers();
  ScriptApp.newTrigger("draftReminders")
    .timeBased()
    .atHour(9)
    .everyDays(1)
    .inTimezone(ONBOARD.timeZone)
    .create();
}

function disableTriggers() {
  const handlers = [
    "sendReminders",
    "sendHrEmails",
    "draftReminders",
    "checkOnboardDates",
    "checkMeetings",
  ];
  ScriptApp.getProjectTriggers().forEach((trigger) => {
    if (handlers.includes(trigger.getHandlerFunction()))
      ScriptApp.deleteTrigger(trigger);
  });
}

function sendReminders() {
  runOnboard_("reminder");
}
function sendHrEmails() {
  runOnboard_("meeting");
}
function draftReminders() {
  runOnboard_("draft");
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
    const now = new Date();
    const spreadsheetId =
      PropertiesService.getUserProperties().getProperty("spreadsheetId") ||
      ONBOARD.spreadsheetId;
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const calendar =
      mode === "meeting" ? CalendarApp.getDefaultCalendar() : null;
    if (mode === "meeting" && (!calendar || !ONBOARD.appointmentTitle)) {
      throw new Error(
        "Configure an accessible HR calendar and exact appointmentTitle.",
      );
    }
    const hrEmail = validateHr_();
    let bookingLinks; // Read the lookup tab once, only when a reminder/draft is due.
    const failures = [];
    spreadsheet.getSheets().forEach((sheet) => {
      if (
        ONBOARD.sheetNames.length &&
        !ONBOARD.sheetNames.includes(sheet.getName())
      )
        return;
      const data = sheet.getDataRange().getValues();
      const header = 1; // Headers start at B2; candidate rows start at B3.
      const candidateHeaders = (data[header] || [])
        .slice(1)
        .map(text_)
        .map((s) => s.toLowerCase());
      if (
        !["full name", "1-month probation catchup"].every((name) =>
          candidateHeaders.includes(name),
        )
      ) {
        if (ONBOARD.sheetNames.includes(sheet.getName()))
          failures.push(
            `${sheet.getName()}: row 2, from column B, must contain Full name and 1-month probation catchup`,
          );
        return;
      }
      if (
        ONBOARD.candidateEmailColumn &&
        !candidateHeaders.includes(ONBOARD.candidateEmailColumn.toLowerCase())
      ) {
        failures.push(
          `${sheet.getName()}: add the ${ONBOARD.candidateEmailColumn} header in row 2`,
        );
        return;
      }
      if (
        (mode === "reminder" || mode === "draft") &&
        !candidateHeaders.includes("pic")
      ) {
        failures.push(`${sheet.getName()}: add the PIC header in row 2`);
        return;
      }
      const headers = data[header].map(text_);
      TRACKING.forEach((name) => {
        if (!headers.some((h) => h.toLowerCase() === name.toLowerCase())) {
          headers.push(name);
          sheet.getRange(header + 1, headers.length).setValue(name);
        }
      });
      const columns = Object.fromEntries(
        headers.map((h, i) => [h.toLowerCase(), i]).slice(1),
      );
      for (let r = header + 1; r < data.length; r++) {
        const row = data[r];
        const get = (name) => row[columns[name.toLowerCase()]];
        const set = (name, value) => {
          sheet
            .getRange(r + 1, columns[name.toLowerCase()] + 1)
            .setValue(value);
          SpreadsheetApp.flush();
        };
        if (!get("1-month probation catchup")) continue;
        try {
          const day = onboardDay_(get("1-month probation catchup"));
          const name = text_(get("Full name"));
          const email = ONBOARD.candidateEmailColumn
            ? text_(get(ONBOARD.candidateEmailColumn)).toLowerCase()
            : "";
          if (!name) throw new Error("Missing candidate name.");
          const timeColumn =
            mode === "reminder"
              ? "Reminder Email Time"
              : mode === "draft"
                ? "Reminder Draft Time"
                : "HR Review Email Time";
          if (get(timeColumn)) continue; // One reminder/draft and one review email per row.
          if (!validEmail_(email))
            throw new Error("Enter a valid Candidate Email.");
          if (mode === "reminder" || mode === "draft") {
            const today = dateKey_(now);
            if (!reminderDue_(day, today)) continue;
            if (text_(get("Candidate Reminder Sent For")) === day) continue; // Already emailed by the previous version.
            const pic = picEmail_(get("PIC"));
            if (!bookingLinks) bookingLinks = bookingLinks_(spreadsheet);
            const bookingUrl = bookingLinks.get(pic);
            if (!bookingUrl || !/^https:\/\/\S+$/.test(bookingUrl)) {
              throw new Error(`Booking Link: enter an HTTPS Link for ${pic}.`);
            }
            const draftSubject = `Please schedule your 1-month probation catch-up - ${day}`;
            const draftBody = `Hello ${name},\n\nYour 1-month probation catch-up date is ${day}, in three days. Please schedule your catch-up with HR using this booking page:\n${bookingUrl}\n\nHow to book:\n1. Open the link and choose a suitable available date and time.\n2. Enter your name and use this email address: ${email}\n3. Complete the booking and keep your confirmation for the meeting details.\n\nPlease use the email address above so HR can match your appointment to your candidate record. If you have already booked your catch-up, keep your existing appointment; you do not need to book again.\n\nThank you,\nHR Team`;
            if (mode === "draft") {
              draftEmail_({
                to: email,
                cc: pic,
                bcc: hrEmail,
                subject: draftSubject,
                body: draftBody,
              });
              set("Reminder Draft Time", new Date());
            } else {
              // Candidate, PIC and HR receive the guidance in a single send request.
              // BCC keeps HR's address out of the candidate's recipient list.
              send_({
                to: email,
                cc: pic,
                bcc: hrEmail,
                subject: draftSubject,
                body: draftBody,
              });
              set("Reminder Email Time", new Date());
            }
          } else {
            // Limit matching to this probation catch-up cycle; no name-only matching.
            const start = new Date(`${day}T00:00:00+07:00`);
            start.setUTCDate(start.getUTCDate() - 3);
            const end = new Date(`${day}T00:00:00+07:00`);
            end.setUTCDate(end.getUTCDate() + 8);
            // Do not repeat review emails already delivered by the previous version.
            if (
              text_(get("HR Form Status")).toLowerCase() === "sent" ||
              get("HR Follow-up Sent At")
            )
              continue;
            const matches = calendar
              .getEvents(start, end)
              .filter(
                (event) =>
                  event.getTitle() === ONBOARD.appointmentTitle &&
                  !event.isAllDayEvent() &&
                  event
                    .getGuestList()
                    .some(
                      (guest) =>
                        guest.getEmail().toLowerCase() === email &&
                        guest.getGuestStatus() !== CalendarApp.GuestStatus.NO,
                    ),
              );
            if (matches.length > 1)
              throw new Error(
                "Multiple check-up meetings for this candidate. Remove or cancel the duplicate booking.",
              );
            if (!matches.length) continue;
            const event = matches[0];
            if (event.getEndTime().getTime() > now.getTime()) continue;
            send_({
              to: hrEmail,
              subject: `HR form required: ${name} - probation catch-up ${day}`,
              body: `Hello,\n\nThe 1-month probation catch-up meeting for ${name} (${email}) has reached its scheduled end time.\nCatch-up date: ${day}\n\nPlease complete the HR form:\n${ONBOARD.hrFormUrl}\n\nThank you.`,
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
function bookingLinks_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(ONBOARD.bookingSheet);
  if (!sheet)
    throw new Error(
      "Add the Booking Link tab with PIC and Link headers in row 1.",
    );
  const rows = sheet.getDataRange().getValues();
  const headers = (rows[0] || []).map((value) => text_(value).toLowerCase());
  const picColumn = headers.indexOf("pic");
  const linkColumn = headers.indexOf("link");
  if (picColumn < 0 || linkColumn < 0)
    throw new Error("Booking Link row 1 must contain PIC and Link.");
  const links = new Map();
  rows.slice(1).forEach((row) => {
    if (!text_(row[picColumn])) return;
    const pic = picEmail_(row[picColumn]);
    if (links.has(pic))
      throw new Error(
        `Booking Link: duplicate PIC ${pic}. Keep one row per PIC.`,
      );
    links.set(pic, text_(row[linkColumn]));
  });
  return links;
}
function picEmail_(value) {
  const raw = text_(value).toLowerCase();
  const email = raw.includes("@") ? raw : `${raw}@kyanon.digital`;
  if (!/^[a-z0-9]+(?:[._-][a-z0-9]+)*@kyanon\.digital$/.test(email)) {
    throw new Error(
      "PIC must be one Kyanon username or @kyanon.digital email.",
    );
  }
  return email;
}
function dateKey_(date) {
  return Utilities.formatDate(date, ONBOARD.timeZone, "yyyy-MM-dd");
}
function onboardDay_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return dateKey_(value);
  const day = text_(value);
  if (
    !/^\d{4}-\d{2}-\d{2}$/.test(day) ||
    isNaN(Date.parse(day)) ||
    new Date(`${day}T00:00:00Z`).toISOString().slice(0, 10) !== day
  ) {
    throw new Error(
      "1-month probation catchup must be a Sheet date or YYYY-MM-DD; ambiguous text dates are rejected.",
    );
  }
  return day;
}
function reminderDue_(day, today) {
  return (Date.parse(day) - Date.parse(today)) / 86400000 === 3;
}
function validateHr_() {
  const email = text_(Session.getEffectiveUser().getEmail());
  if (!validEmail_(email))
    throw new Error(
      "Cannot identify the HR account. HR must authorize and enable the script.",
    );
  if (!/^https:\/\/\S+$/.test(ONBOARD.hrFormUrl))
    throw new Error("Set an HTTPS hrFormUrl.");
  return email;
}
function draftEmail_(message) {
  const to = message.to ? message.to.split(",").filter(Boolean).join(",") : "";
  const cc = message.cc ? message.cc.split(",").filter(Boolean).join(",") : "";
  const bcc = message.bcc
    ? message.bcc.split(",").filter(Boolean).join(",")
    : "";
  GmailApp.createDraft(to, message.subject, message.body, { cc: cc, bcc: bcc });
}

function send_(message) {
  const recipients = [message.to, message.cc, message.bcc]
    .filter(Boolean)
    .reduce((count, emails) => count + emails.split(",").length, 0);
  if (MailApp.getRemainingDailyQuota() < recipients)
    throw new Error("Insufficient daily email quota.");
  MailApp.sendEmail(message);
}
