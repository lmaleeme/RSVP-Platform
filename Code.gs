/**
 * RSVP Platform - Google Apps Script backend
 *
 * 1. Create a new Google Sheet (or use an existing bound spreadsheet).
 * 2. Paste this file into Extensions > Apps Script as Code.gs.
 * 3. Run setupRSVPPlatform() once and approve permissions.
 * 4. Deploy as a Web app:
 *      Execute as: Me
 *      Who has access: Anyone
 * 5. Copy the Web app URL into CONFIG.SHEET_ENDPOINT_URL in the RSVP HTML.
 *
 * Public settings are managed through the Admin Portal (?page=admin).
 * Responses are saved in the "RSVP Responses" sheet.
 * Set an admin password once by running setRSVPAdminPassword('your-strong-password') from the editor.
 */

const RSVP_CONFIG_SHEET = 'Config';
const RSVP_RESPONSES_SHEET = 'RSVP Responses';
const RSVP_SPREADSHEET_PROPERTY = 'RSVP_SPREADSHEET_ID';
const RSVP_DEFAULT_ARTWORK_URL = 'https://lmaleeme.github.io/RSVP-Platform/sipnsee.png';

const RSVP_DEFAULT_SETTINGS = [
  ['campaignTitle', "You're invited", 'Main invitation heading'],
  ['introText', 'We would be delighted if you could join us.', 'Short line beneath the heading'],
  ['invitationMessage', 'We look forward to celebrating with you. Please let us know if you can attend.', 'Main invitation message shown in the message panel'],
  ['rsvpByDate', '', 'Last date to RSVP, format YYYY-MM-DD'],
  ['logoUrl', '', 'Optional replacement invitation image URL'],
  ['eventDate', '2026-11-24', 'Event date YYYY-MM-DD'],
  ['eventStartTime', '17:30', 'Event start time HH:MM'],
  ['eventEndTime', '20:00', 'Event end time HH:MM'],
  ['eventVenue', 'The Art Quarter, Michelangelo Towers', 'Calendar event location'],
  ['backgroundMediaUrl', '', 'Optional background image or video URL'],
  ['backgroundMediaType', 'none', 'Use image, video, or none'],
  ['primaryColor', '#6B4E71', 'Button and accent colour'],
  ['accentColor', '#E8B54D', 'Highlight colour'],
  ['backgroundColor', '#1B1B3A', 'Page background colour'],
  ['textColor', '#F5F1E8', 'Main light text colour']
];

const RSVP_HEADERS = [
  'RSVP Reference',
  'First Name',
  'Surname',
  'Mobile Number',
  'Attendance',
  'Comments',
  'Submitted At',
  'Received At'
];

/**
 * Run once to create/configure the spreadsheet tabs.
 * If this is a spreadsheet-bound script, it uses that spreadsheet.
 * If standalone, it creates a new spreadsheet and stores its ID.
 */
function setupRSVPPlatform() {
  const props = PropertiesService.getScriptProperties();
  let ss;
  const savedId = props.getProperty(RSVP_SPREADSHEET_PROPERTY);

  if (savedId) {
    ss = SpreadsheetApp.openById(savedId);
  } else {
    ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      ss = SpreadsheetApp.create('RSVP Platform');
    }
    props.setProperty(RSVP_SPREADSHEET_PROPERTY, ss.getId());
  }

  let config = ss.getSheetByName(RSVP_CONFIG_SHEET);
  if (!config) config = ss.insertSheet(RSVP_CONFIG_SHEET);
  if (config.getLastRow() === 0) {
    config.getRange(1, 1, 1, 3).setValues([['Setting', 'Value', 'Description']]);
    config.getRange(2, 1, RSVP_DEFAULT_SETTINGS.length, 3)
      .setValues(RSVP_DEFAULT_SETTINGS);
    config.setFrozenRows(1);
    config.autoResizeColumns(1, 3);
  }

  let responses = ss.getSheetByName(RSVP_RESPONSES_SHEET);
  if (!responses) responses = ss.insertSheet(RSVP_RESPONSES_SHEET);
  if (responses.getLastRow() === 0) {
    responses.getRange(1, 1, 1, RSVP_HEADERS.length).setValues([RSVP_HEADERS]);
    responses.setFrozenRows(1);
    responses.getRange(1, 1, 1, RSVP_HEADERS.length).setFontWeight('bold');
    responses.autoResizeColumns(1, RSVP_HEADERS.length);
  }

  Logger.log('RSVP spreadsheet: ' + ss.getUrl());
  Logger.log('Spreadsheet ID saved in Script Properties.');
  return ss.getUrl();
}

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  if (params.page === 'index' || params.page === 'rsvp') {
    // Do not read SpreadsheetApp here. Return the public page immediately;
    // the browser fetches editable settings asynchronously after first paint.
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Sip & See | RSVP')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (params.page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin')
      .setTitle('RSVP Admin Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (params.action === 'getPublicConfig') return jsonOutput_(getPublicConfig_());
  return jsonOutput_({
    status: 'ok',
    message: 'RSVP platform endpoint is running.',
    adminPortal: 'Open this web app URL with ?page=admin',
    actions: ['getPublicConfig', 'submitRSVP']
  });
}

function safeJsonForHtml_(value) {
  return JSON.stringify(value || {})
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026')
    .replace(/\u2028/g, '\\u2028')
    .replace(/\u2029/g, '\\u2029');
}

function doPost(e) {
  try {
    const raw = e && e.postData && e.postData.contents
      ? e.postData.contents
      : '{}';
    const payload = JSON.parse(raw);

    if (payload.action === 'getPublicConfig') {
      return jsonOutput_(getPublicConfig_());
    }

    return jsonOutput_(saveRSVP_(payload));
  } catch (err) {
    return jsonOutput_({
      status: 'error',
      message: err && err.message ? err.message : 'Unable to process request.'
    });
  }
}

/**
 * FIRST-TIME ADMIN PASSWORD SETUP:
 * Replace the placeholder below with a unique password of at least 12 characters,
 * run configureRSVPAdminPassword() once from the Apps Script editor, then remove
 * the password literal from this function and save the project.
 */
function configureRSVPAdminPassword() {
  const password = 'REPLACE_WITH_A_UNIQUE_PASSWORD_12_CHARS_OR_LONGER';
  if (password.indexOf('REPLACE_WITH_') === 0) {
    throw new Error('Edit configureRSVPAdminPassword() and set your own password first.');
  }
  return setRSVPAdminPassword(password);
}

function setRSVPAdminPassword(password) {
  const value = String(password || '');
  if (value.length < 12) throw new Error('Use an admin password of at least 12 characters.');
  PropertiesService.getScriptProperties().setProperty('RSVP_ADMIN_PASSWORD_HASH', hashAdminPassword_(value));
  return 'RSVP admin password configured.';
}

function adminLogin(password) {
  const props = PropertiesService.getScriptProperties();
  let stored = props.getProperty('RSVP_ADMIN_PASSWORD_HASH');

  // Backward compatibility: an earlier setup instruction used ADMIN_PASSWORD.
  // On first successful lookup, migrate that value to a SHA-256 hash and
  // remove the plaintext Script Property.
  if (!stored) {
    const legacyPassword = props.getProperty('ADMIN_PASSWORD');
    if (legacyPassword) {
      stored = hashAdminPassword_(legacyPassword);
      props.setProperty('RSVP_ADMIN_PASSWORD_HASH', stored);
      props.deleteProperty('ADMIN_PASSWORD');
    }
  }

  if (!stored) {
    throw new Error('Admin password is not configured. Set Script Property ADMIN_PASSWORD, or run configureRSVPAdminPassword() once.');
  }
  return constantTimeEquals_(stored, hashAdminPassword_(String(password || '')));
}

function adminGetSettings(password) {
  requireAdmin_(password);
  const theme = getPublicConfig_().theme;
  return theme;
}

function adminSaveSettings(settings, password) {
  requireAdmin_(password);
  if (!settings || typeof settings !== 'object') throw new Error('No settings were supplied.');
  const allowed = RSVP_DEFAULT_SETTINGS.map(row => row[0]);
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(RSVP_CONFIG_SHEET);
  if (!sheet) throw new Error('Config sheet not found. Run setupRSVPPlatform().');
  const values = sheet.getDataRange().getValues();
  const rowByKey = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) rowByKey[key] = i + 1;
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    allowed.forEach(key => {
      if (settings[key] === undefined) return;
      let value = String(settings[key] == null ? '' : settings[key]).trim();
      if (key.toLowerCase().includes('color') && value && !/^#[0-9a-f]{6}$/i.test(value)) {
        throw new Error('Invalid colour value for ' + key + '.');
      }
      if (key === 'backgroundMediaType' && ['none','image','video'].indexOf(value) === -1) {
        throw new Error('Background type must be none, image or video.');
      }
      if ((key === 'logoUrl' || key === 'backgroundMediaUrl') && value && !/^https:\/\//i.test(value)) {
        throw new Error('Media URLs must use HTTPS.');
      }
      if ((key === 'rsvpByDate' || key === 'eventDate') && value && !/^\d{4}-\d{2}-\d{2}$/.test(value)) {
        throw new Error('RSVP deadline must be a valid date.');
      }
      if ((key === 'eventStartTime' || key === 'eventEndTime') && value && !/^([01]\d|2[0-3]):[0-5]\d$/.test(value)) throw new Error('Invalid time for ' + key);
      if (value.length > 5000) throw new Error(key + ' is too long.');
      if (rowByKey[key]) sheet.getRange(rowByKey[key], 2).setValue(value);
      else { sheet.appendRow([key, value, 'Managed from RSVP Admin Portal']); rowByKey[key] = sheet.getLastRow(); }
    });
  } finally { lock.releaseLock(); }
  return {status:'ok', publishedAt:String(Date.now())};
}

function requireAdmin_(password) {
  if (!adminLogin(password)) throw new Error('Admin authentication failed. Please sign in again.');
}

function hashAdminPassword_(value) {
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, value, Utilities.Charset.UTF_8);
  return digest.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}

function constantTimeEquals_(a, b) {
  a = String(a); b = String(b);
  if (a.length !== b.length) return false;
  let diff = 0;
  for (let i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

/**
 * Returns the public-facing invitation settings.
 * No private spreadsheet data or RSVP responses are exposed.
 */
// Public client-callable wrappers. Apps Script does not expose functions
// whose names end in an underscore to google.script.run.
function getPublicConfig() {
  return getPublicConfig_();
}

function submitRSVP(payload) {
  return saveRSVP_(payload);
}

function getPublicConfig_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(RSVP_CONFIG_SHEET);
  if (!sheet) throw new Error('Config sheet not found. Run setupRSVPPlatform().');

  const values = sheet.getDataRange().getValues();
  const settings = {};
  for (let i = 1; i < values.length; i++) {
    const key = String(values[i][0] || '').trim();
    if (key) settings[key] = values[i][1] == null ? '' : String(values[i][1]).trim();
  }

  return {
    status: 'ok',
    publishedAt: String(Date.now()),
    theme: {
      campaignTitle: settings.campaignTitle || "You're invited",
      introText: settings.introText || '',
      invitationMessage: settings.invitationMessage || '',
      rsvpByDate: settings.rsvpByDate || '',
      logoUrl: settings.logoUrl || RSVP_DEFAULT_ARTWORK_URL,
      eventDate: settings.eventDate || '2026-11-24',
      eventStartTime: settings.eventStartTime || '17:30',
      eventEndTime: settings.eventEndTime || '20:00',
      eventVenue: settings.eventVenue || 'The Art Quarter, Michelangelo Towers',
      backgroundMediaUrl: settings.backgroundMediaUrl || '',
      backgroundMediaType: settings.backgroundMediaType || 'none',
      primaryColor: settings.primaryColor || '',
      accentColor: settings.accentColor || '',
      backgroundColor: settings.backgroundColor || '',
      textColor: settings.textColor || ''
    }
  };
}

/**
 * Validates and appends one RSVP response.
 */
function saveRSVP_(payload) {
  const firstName = cleanText_(payload.firstName, 100);
  const lastName = cleanText_(payload.lastName, 100);
  const mobile = cleanText_(payload.mobile, 40);
  const attendance = cleanText_(payload.attendance, 40);
  const comments = cleanText_(payload.comments, 2000);
  const reference = cleanText_(payload.rsvpReference, 80) || createReference_();

  if (!firstName || !lastName) {
    throw new Error('First name and surname are required.');
  }
  if (['Yes', 'Yes with Partner', 'Not attending'].indexOf(attendance) === -1) {
    throw new Error('Please provide a valid RSVP response.');
  }

  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(RSVP_RESPONSES_SHEET);
  if (!sheet) throw new Error('RSVP Responses sheet not found. Run setupRSVPPlatform().');

  const submittedAt = cleanText_(payload.submittedAt, 60);
  const receivedAt = new Date();

  const row = [
    safeSheetText_(reference),
    safeSheetText_(firstName),
    safeSheetText_(lastName),
    safeSheetText_(mobile),
    safeSheetText_(attendance),
    safeSheetText_(comments),
    safeSheetText_(submittedAt),
    receivedAt
  ];

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet.appendRow(row);
  } finally {
    lock.releaseLock();
  }

  return {
    status: 'ok',
    message: 'RSVP received.',
    rsvpReference: reference
  };
}

function getSpreadsheet_() {
  const id = PropertiesService.getScriptProperties()
    .getProperty(RSVP_SPREADSHEET_PROPERTY);
  if (!id) {
    throw new Error('Spreadsheet is not configured. Run setupRSVPPlatform() first.');
  }
  return SpreadsheetApp.openById(id);
}

function cleanText_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength);
}

/**
 * Prevent user-entered values beginning with formula characters from
 * being interpreted as spreadsheet formulas.
 */
function safeSheetText_(value) {
  const text = String(value == null ? '' : value);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function createReference_() {
  return 'RSVP-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
}

function jsonOutput_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
