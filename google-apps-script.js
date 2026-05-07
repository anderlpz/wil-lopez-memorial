/**
 * ═══════════════════════════════════════════════════════════════════════════════
 *  Wil Lopez Memorial — RSVP Google Apps Script Backend
 *  Wilbum "Wil" Lopez · December 31, 1949 – April 27, 2026
 *  United States Marine Corps · Vietnam Veteran · Semper Fidelis
 * ═══════════════════════════════════════════════════════════════════════════════
 *
 *  SETUP INSTRUCTIONS
 *  ──────────────────────────────────────────────────────────────────────────────
 *  1. Go to: https://script.google.com
 *  2. Click "New project"
 *  3. Name it: "Wil Lopez Memorial RSVP"
 *  4. Delete all existing code and paste this entire file
 *  5. Set DASHBOARD_KEY below to any secret string you choose (e.g. a random
 *     word or short phrase). The dashboard page uses this same key.
 *  6. Click "Deploy" → "New deployment"
 *  7. Settings:
 *       Type:       Web app
 *       Execute as: Me (your Google account)
 *       Who has access: Anyone
 *  8. Click "Deploy" — authorize permissions when prompted
 *  9. Copy the Web app URL (looks like: https://script.google.com/macros/s/AKfycb.../exec)
 * 10. In rsvp.html, replace 'YOUR_GOOGLE_APPS_SCRIPT_URL_HERE' with that URL
 * 11. In dashboard.html, use the same URL with ?action=dashboard&key=YOUR_KEY
 *
 *  The script Wil automatically create a Google Sheet called "Wil Lopez RSVPs"
 *  in your Google Drive, and write each submission as a new row.
 *
 *  NOTE: After editing this file, you must deploy a NEW version (not update
 *  the existing one) for changes to take effect in production.
 * ═══════════════════════════════════════════════════════════════════════════════
 */

/* ── Configuration ───────────────────────────────────────────────────────────── */
var SPREADSHEET_NAME = 'Wil Lopez Memorial RSVPs';
var SHEET_NAME       = 'RSVPs';

/**
 * Dashboard read key — change this before deploying.
 * Anyone who knows this string can read all RSVPs via ?action=dashboard&key=…
 * Not cryptographic security, just keeps random hits from reading the data.
 */
var DASHBOARD_KEY = 'CHANGE_ME_BEFORE_DEPLOYING';

/* Column headers — must match the order in appendRsvpRow() below */
var HEADERS = [
  'Submitted (PT)',
  'Name',
  'Email',
  'Guests',
  'Events Attending',
  'Message',
  'Submitted At (UTC)'
];


/* ═══════════════════════════════════════════════════════════════════════════════
   doPost — Main entry point for RSVP submissions
   ───────────────────────────────────────────────────────────────────────────────
   Called when the HTML form POSTs to this web app URL.
   Content-Type: text/plain (body is a JSON string — avoids CORS preflight).
═══════════════════════════════════════════════════════════════════════════════ */
function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    /* Parse JSON body */
    if (!e || !e.postData || !e.postData.contents) {
      throw new Error('No POST data received.');
    }

    var data = JSON.parse(e.postData.contents);

    /* ── Anti-spam: honeypot field ─────────────────────────────────────────
       The form has a hidden "website" field that real users never fill in.
       Bots that blindly fill every field Wil populate it — we silently
       succeed so the bot thinks its submission worked. */
    if (data.website) {
      output.setContent(JSON.stringify({ success: true }));
      return output;
    }

    /* ── Anti-spam: submission speed check ─────────────────────────────────
       Legitimate users take at least a few seconds to read and fill the form.
       If _loadTime was sent and the round-trip is under 3 seconds, it's a bot. */
    if (data._loadTime && (Date.now() - Number(data._loadTime)) < 3000) {
      output.setContent(JSON.stringify({ success: true }));
      return output;
    }

    /* ── Required field validation ─────────────────────────────────────────
       Name, email, and guest count must be present and non-empty. */
    if (!data.name || !String(data.name).trim()) {
      throw new Error('Name is required.');
    }
    if (!data.email || !String(data.email).trim()) {
      throw new Error('Email address is required.');
    }
    if (data.guests === undefined || data.guests === null || data.guests === '') {
      throw new Error('Guest count is required.');
    }

    /* ── Anti-spam: per-email rate limit (1 submission per hour) ───────────
       Scan existing rows for a matching email submitted in the last 60 min.
       Uses the UTC ISO timestamp stored in column 7 (index 6). */
    // Check if this email already has an RSVP (update vs new)
    var emailLower = String(data.email).trim().toLowerCase();
    var isUpdate = false;
    try {
      var files = DriveApp.getFilesByName(SPREADSHEET_NAME);
      if (files.hasNext()) {
        var ss = SpreadsheetApp.open(files.next());
        var checkSheet = ss.getSheetByName(SHEET_NAME);
        if (checkSheet && checkSheet.getLastRow() > 1) {
          var emails = checkSheet.getRange(2, 3, checkSheet.getLastRow() - 1, 1).getValues();
          for (var i = 0; i < emails.length; i++) {
            if (String(emails[i][0]).toLowerCase().trim() === emailLower) {
              isUpdate = true;
              break;
            }
          }
        }
      }
    } catch(e) { /* Sheet doesn't exist yet — treat as new */ }

    // Only rate-limit NEW submissions (not updates)
    if (!isUpdate && isRateLimited_(emailLower)) {
      throw new Error(
        'It looks like we already received an RSVP from this email address ' +
        'recently. If you need to make a change, please reach out directly. ' +
        'Thank you for your patience.'
      );
    }

    /* Write to sheet */
    var sheet = getOrCreateSheet();
    appendRsvpRow(sheet, data);

    /* Optional: send a confirmation email to the submitter */
    /* sendConfirmationEmail(data); */

    output.setContent(JSON.stringify({ success: true }));

  } catch (err) {
    /* Log for debugging in Apps Script execution log */
    Logger.log('RSVP Error: ' + err.toString());
    Logger.log('Raw body: ' + (e && e.postData ? e.postData.contents : 'none'));

    output.setContent(JSON.stringify({
      success: false,
      error: err.message || err.toString()
    }));
  }

  return output;
}


/* ═══════════════════════════════════════════════════════════════════════════════
   doGet — Health check / info endpoint + dashboard data endpoint
   ───────────────────────────────────────────────────────────────────────────────
   GET  /exec                                → health check (public)
   GET  /exec?action=dashboard&key=SECRET    → full RSVP data (key-protected)
═══════════════════════════════════════════════════════════════════════════════ */
function doGet(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  var params = (e && e.parameter) ? e.parameter : {};

  /* ── Dashboard endpoint ──────────────────────────────────────────────────── */
  if (params.action === 'dashboard') {

    /* Key check — wrong or missing key gets a terse 401-style response */
    if (!params.key || params.key !== DASHBOARD_KEY) {
      output.setContent(JSON.stringify({ error: 'unauthorized' }));
      return output;
    }

    try {
      var sheet = getOrCreateSheet();
      var lastRow = sheet.getLastRow();

      /* No data rows yet (only header row, or sheet is empty) */
      if (lastRow <= 1) {
        output.setContent(JSON.stringify({
          summary: {
            total_rsvps:       0,
            total_guests:      0,
            ceremony_count:    0,
            celebration_count: 0
          },
          rsvps: []
        }));
        return output;
      }

      /* Read all data rows (skip row 1 = headers) */
      var numRows  = lastRow - 1;
      var rawRows  = sheet.getRange(2, 1, numRows, HEADERS.length).getValues();

      var rsvps             = [];
      var total_guests      = 0;
      var ceremony_count    = 0;
      var celebration_count = 0;

      for (var i = 0; i < rawRows.length; i++) {
        var row    = rawRows[i];
        var events = String(row[4] || '');

        /* Map raw row to a readable object using column positions */
        var rsvp = {
          submitted_pt: String(row[0] || ''),
          name:         String(row[1] || ''),
          email:        String(row[2] || ''),
          guests:       Number(row[3]) || 0,
          events:       events,
          message:      String(row[5] || ''),
          submitted_at: row[6] instanceof Date
                          ? row[6].toISOString()
                          : String(row[6] || '')
        };

        rsvps.push(rsvp);
        total_guests += rsvp.guests;

        /* Count attendees by event type (case-insensitive keyword match) */
        var eventsLower = events.toLowerCase();
        if (eventsLower.indexOf('ceremony') !== -1) {
          ceremony_count++;
        }
        if (eventsLower.indexOf('celebration') !== -1 ||
            eventsLower.indexOf('reception')   !== -1) {
          celebration_count++;
        }
      }

      output.setContent(JSON.stringify({
        summary: {
          total_rsvps:       rsvps.length,
          total_guests:      total_guests,
          ceremony_count:    ceremony_count,
          celebration_count: celebration_count
        },
        rsvps: rsvps
      }));

    } catch (err) {
      Logger.log('Dashboard Error: ' + err.toString());
      output.setContent(JSON.stringify({ error: err.message || err.toString() }));
    }

    return output;
  }

  /* ── Default: health check ───────────────────────────────────────────────── */
  output.setContent(JSON.stringify({
    status:  'ok',
    service: 'Wil Lopez Memorial RSVP',
    message: 'Semper Fidelis. POST your RSVP to this endpoint.'
  }));
  return output;
}


/* ═══════════════════════════════════════════════════════════════════════════════
   isRateLimited_ — Return true if the email submitted within the last hour
   ───────────────────────────────────────────────────────────────────────────────
   Scans existing sheet rows. Uses the UTC ISO timestamp in column 7 (index 6).
   Returns false (not rate-limited) if the sheet doesn't exist yet or is empty.
═══════════════════════════════════════════════════════════════════════════════ */
function isRateLimited_(emailLower) {
  try {
    var files = DriveApp.getFilesByName(SPREADSHEET_NAME);
    if (!files.hasNext()) return false;                  /* no sheet yet */

    var ss    = SpreadsheetApp.openById(files.next().getId());
    var sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return false;

    var lastRow = sheet.getLastRow();
    if (lastRow <= 1) return false;                      /* only headers */

    var oneHourAgo = Date.now() - (60 * 60 * 1000);
    var rows       = sheet.getRange(2, 1, lastRow - 1, HEADERS.length).getValues();

    for (var i = 0; i < rows.length; i++) {
      var rowEmail = String(rows[i][2] || '').trim().toLowerCase();
      if (rowEmail !== emailLower) continue;

      /* Column 7 (index 6) holds the UTC ISO timestamp */
      var rawTs   = rows[i][6];
      var tsMs    = rawTs instanceof Date
                      ? rawTs.getTime()
                      : new Date(String(rawTs)).getTime();

      if (!isNaN(tsMs) && tsMs > oneHourAgo) return true;
    }
  } catch (_) {
    /* Non-fatal — if we can't read the sheet, allow the submission through */
  }

  return false;
}


/* ═══════════════════════════════════════════════════════════════════════════════
   getOrCreateSheet — Find or create the RSVP spreadsheet and worksheet
═══════════════════════════════════════════════════════════════════════════════ */
function getOrCreateSheet() {
  var ss   = null;
  var file = null;

  /* Try to find an existing spreadsheet by name */
  var files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    file = files.next();
    ss   = SpreadsheetApp.openById(file.getId());
  } else {
    /* Create a new one */
    ss = SpreadsheetApp.create(SPREADSHEET_NAME);
    formatSpreadsheet_(ss);
  }

  /* Get or create the RSVPs tab */
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    /* Remove the default "Sheet1" if it exists and is empty */
    var defaultSheet = ss.getSheetByName('Sheet1');
    if (defaultSheet && defaultSheet.getLastRow() === 0) {
      ss.deleteSheet(defaultSheet);
    }
    writeHeaders_(sheet);
  } else if (sheet.getLastRow() === 0) {
    /* Sheet exists but is empty — write headers */
    writeHeaders_(sheet);
  }

  return sheet;
}


/* ═══════════════════════════════════════════════════════════════════════════════
   appendRsvpRow — Write one RSVP submission as a new row
═══════════════════════════════════════════════════════════════════════════════ */
function appendRsvpRow(sheet, data) {
  var ptTime = Utilities.formatDate(
    new Date(),
    'America/Los_Angeles',
    'yyyy-MM-dd HH:mm:ss'
  ) + ' PT';

  var events = '';
  if (Array.isArray(data.events) && data.events.length > 0) {
    events = data.events.join(', ');
  } else if (typeof data.events === 'string') {
    events = data.events;
  } else {
    events = '(none selected)';
  }

// Upsert: check if this email already has an RSVP, update if so
  var existingRow = -1;
  if (sheet.getLastRow() > 1) {
    var emails = sheet.getRange(2, 3, sheet.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < emails.length; i++) {
      if (String(emails[i][0]).toLowerCase().trim() === String(data.email).toLowerCase().trim()) {
        existingRow = i + 2; // +2 because data starts at row 2 (1-indexed, skip header)
        break;
      }
    }
  }

  var rowData = [
    ptTime,                           /* Submitted (PT)       */
    data.name    || '',               /* Name                 */
    data.email   || '',               /* Email                */
    data.guests  || 1,                /* Guests               */
    events,                           /* Events Attending     */
    data.message || '',               /* Message              */
    new Date().toISOString()          /* UTC timestamp        */
  ];

  if (existingRow > 0) {
    // Update existing row
    sheet.getRange(existingRow, 1, 1, 7).setValues([rowData]);
  } else {
    // New submission
    sheet.appendRow(rowData);
  }

  /* Auto-resize columns for readability */
  try {
    sheet.autoResizeColumns(1, HEADERS.length);
  } catch (_) { /* non-fatal */ }
}


/* ═══════════════════════════════════════════════════════════════════════════════
   writeHeaders_ — Write styled header row to a fresh sheet
═══════════════════════════════════════════════════════════════════════════════ */
function writeHeaders_(sheet) {
  var headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
  headerRange.setValues([HEADERS]);

  /* Style the header row */
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#1a1714');
  headerRange.setFontColor('#c4a97a');
  headerRange.setFontFamily('Arial');
  headerRange.setFontSize(10);

  /* Freeze the header row */
  sheet.setFrozenRows(1);

  /* Set a reasonable default column width */
  sheet.setColumnWidth(1, 170);  /* Submitted (PT)   */
  sheet.setColumnWidth(2, 180);  /* Name             */
  sheet.setColumnWidth(3, 220);  /* Email            */
  sheet.setColumnWidth(4, 60);   /* Guests           */
  sheet.setColumnWidth(5, 340);  /* Events           */
  sheet.setColumnWidth(6, 320);  /* Message          */
  sheet.setColumnWidth(7, 200);  /* UTC ISO          */
}


/* ═══════════════════════════════════════════════════════════════════════════════
   formatSpreadsheet_ — Apply basic formatting to a brand-new spreadsheet
═══════════════════════════════════════════════════════════════════════════════ */
function formatSpreadsheet_(ss) {
  try {
    /* Set the spreadsheet tab color */
    var sheet = ss.getSheets()[0];
    sheet.setName(SHEET_NAME);
    sheet.setTabColor('#c4a97a');
    writeHeaders_(sheet);
  } catch (_) { /* non-fatal */ }
}


/* ═══════════════════════════════════════════════════════════════════════════════
   sendConfirmationEmail — (OPTIONAL) Send a thank-you email to the RSVP
   ───────────────────────────────────────────────────────────────────────────────
   Uncomment the call in doPost() to enable this.
   Requires the Gmail scope in your Apps Script project.
═══════════════════════════════════════════════════════════════════════════════ */
/*
function sendConfirmationEmail(data) {
  if (!data.email) return;

  var subject = 'Thank You for Your RSVP — Wilbum "Wil" Lopez';
  var body = [
    'Dear ' + data.name + ',',
    '',
    'Thank you for letting us know you Wil be with us to celebrate Wil\'s life.',
    'Your RSVP has been received.',
    '',
    'We look forward to honoring Wil alongside all who loved him.',
    '',
    'Semper Fidelis.',
    '',
    '— The Lopez Family'
  ].join('\n');

  GmailApp.sendEmail(data.email, subject, body, {
    name: 'The Lopez Family'
  });
}
*/


/* ═══════════════════════════════════════════════════════════════════════════════
   testSubmission — Manual test function (run from Apps Script editor)
   ───────────────────────────────────────────────────────────────────────────────
   Click ▶ Run with this function selected to test without the HTML form.
═══════════════════════════════════════════════════════════════════════════════ */
function testSubmission() {
  var fakePost = {
    postData: {
      contents: JSON.stringify({
        name:        'Test Guest',
        email:       'test@example.com',
        guests:      2,
        events:      ['Marine Corps Ceremony', 'Memorial Service', 'Reception'],
        message:     'Wil was a wonderful man. We are honored to celebrate his life.',
        submittedAt: new Date().toISOString()
      })
    }
  };

  var result = doPost(fakePost);
  Logger.log('Test result: ' + result.getContent());

  /* Open the spreadsheet so you can verify the row */
  var files = DriveApp.getFilesByName(SPREADSHEET_NAME);
  if (files.hasNext()) {
    var url = SpreadsheetApp.openById(files.next().getId()).getUrl();
    Logger.log('Spreadsheet URL: ' + url);
  }
}
