// ════════════════════════════════════════════════════════════════════
// JSR Immigration Ltd. — Unified Intake Apps Script
// Handles: pgwp | visitor_visa | visitor_record | study_extension | work_permit
//
// HOW TO DEPLOY:
//   1. Open script.google.com → New project
//   2. Paste this entire file into Code.gs
//   3. Update SPREADSHEET_ID, DRIVE_FOLDER_ID, and RCIC_EMAIL below
//   4. Deploy → New deployment → Web app
//      Execute as: Me | Who has access: Anyone
//   5. Copy the web app URL into each index.html's APPS_SCRIPT_URL constant
// ════════════════════════════════════════════════════════════════════

// ── CONFIG ──────────────────────────────────────────────────────────
// Replace these three values with your own:
const SPREADSHEET_ID  = 'YOUR_SPREADSHEET_ID_HERE';   // Google Sheet ID (from its URL)
const DRIVE_FOLDER_ID = 'YOUR_DRIVE_FOLDER_ID_HERE';  // Parent "Open Files" folder ID
const RCIC_EMAIL      = 'info@jsrimmigration.com';    // Your email (receives draft)

// Maximum characters of the client's situation text included inline in the email body.
// The full text is always stored in the PDF and the Google Sheet.
const MAX_EMAIL_SITUATION_LENGTH = 800;

// ════════════════════════════════════════════════════════════════════
// ENTRY POINT
// ════════════════════════════════════════════════════════════════════
function doPost(e) {
  // Guard against un-configured deployment
  if (SPREADSHEET_ID === 'YOUR_SPREADSHEET_ID_HERE' || DRIVE_FOLDER_ID === 'YOUR_DRIVE_FOLDER_ID_HERE') {
    return jsonResponse('error', 'Apps Script not configured: please set SPREADSHEET_ID and DRIVE_FOLDER_ID in Code.gs before deploying.');
  }

  try {
    const payload = JSON.parse(e.parameter.payload);
    const formType = payload.form_type || 'unknown';

    switch (formType) {
      case 'pgwp':             return handlePgwp(payload);
      case 'visitor_visa':     return handleVisitorVisa(payload);
      case 'visitor_record':   return handleVisitorRecord(payload);
      case 'study_extension':  return handleStudyExtension(payload);
      case 'work_permit':      return handleWorkPermit(payload);
      default:
        return jsonResponse('error', 'Unknown form_type: ' + formType);
    }
  } catch (err) {
    return jsonResponse('error', err.message);
  }
}

// ════════════════════════════════════════════════════════════════════
// WORK PERMIT HANDLER
// ════════════════════════════════════════════════════════════════════
function handleWorkPermit(d) {
  const name      = d.full_name        || 'Unknown';
  const email     = d.email            || '';
  const timestamp = d.submission_timestamp || new Date().toISOString();
  const dateStr   = timestamp.slice(0, 10);
  const category  = d.wp_category      || '';
  const situation = d.situation_explanation || '';

  // 1. Create client folder in Drive
  const parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const folderName   = 'WP — ' + name + ' — ' + dateStr;
  const clientFolder = parentFolder.createFolder(folderName);
  const folderUrl    = clientFolder.getUrl();

  // 2. Build human-readable summary
  const summary = buildWorkPermitSummary(d);

  // 3. Create PDF in folder
  const pdfBlob = buildPdf('Work Permit Intake — ' + name, summary);
  clientFolder.createFile(pdfBlob);

  // 4. Log row to "Work Permit" sheet tab
  logWorkPermitRow(d, folderUrl);

  // 5. Draft email
  draftWorkPermitEmail(name, email, category, situation, summary, folderUrl, dateStr);

  return jsonResponse('ok', 'Work permit intake saved.');
}

// ════════════════════════════════════════════════════════════════════
// SHEET LOGGING — Work Permit
// ════════════════════════════════════════════════════════════════════
function logWorkPermitRow(d, folderUrl) {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tab   = getOrCreateSheet(ss, 'Work Permit');
  const now   = new Date();

  // Write headers on row 1 if sheet is empty
  if (tab.getLastRow() === 0) {
    const headers = [
      'Submitted At',
      'Full Name',
      'DOB',
      'Country of Birth',
      'Citizenship',
      'UCI',
      'Current Status',
      'Status Valid To',
      'Passport Number',
      'Passport Country',
      'Passport Expiry',
      'Marital Status',
      'Phone',
      'Email',
      'Alt Phone',
      'Address',
      'City',
      'Province',
      'Postal Code',
      'First Entry Date',
      'First Entry Place',
      'Recent Entry Date',
      'Recent Entry Place',
      'WP Category',
      'WP Category (Other)',
      'Employer',
      'Job Title',
      'NOC Code',
      'LMIA Number',
      'Job Offer',
      'Admissibility: Criminal',
      'Admissibility: Refusal',
      'Admissibility: Prev Canada',
      'Admissibility: Removal',
      'Admissibility: Overstay',
      'Admissibility: TB/Medical',
      'Admissibility: Pending App',
      'Admissibility: Misrep',
      'Admissibility: Security',
      'Situation / Explanation',
      'Signed Name',
      'Signed Date',
      'Drive Folder URL',
    ];
    tab.appendRow(headers);
    tab.getRange(1, 1, 1, headers.length)
       .setFontWeight('bold')
       .setBackground('#1a3a5c')
       .setFontColor('#ffffff');
    tab.setFrozenRows(1);
  }

  const row = [
    now,
    d.full_name           || '',
    d.dob                 || '',
    d.country_birth       || '',
    d.citizenship         || '',
    d.uci                 || '',
    d.current_status      || '',
    d.status_to           || '',
    d.passport_number     || '',
    d.passport_country    || '',
    d.passport_expiry     || '',
    d.marital_status      || '',
    d.phone               || '',
    d.email               || '',
    d.alt_phone           || '',
    d.address_street      || '',
    d.address_city        || '',
    d.address_province    || '',
    d.postal_code         || '',
    d.first_entry_date    || '',
    d.first_entry_place   || '',
    d.recent_entry_date   || '',
    d.recent_entry_place  || '',
    d.wp_category         || '',
    d.wp_category_other   || '',
    d.wp_employer         || '',
    d.wp_job_title        || '',
    d.wp_noc              || '',
    d.wp_lmia_number      || '',
    d.job_offer           || '',
    d.aq_criminal         || '',
    d.aq_refusal          || '',
    d.aq_prev_canada      || '',
    d.aq_removal          || '',
    d.aq_overstay         || '',
    d.aq_tb               || '',
    d.aq_pending          || '',
    d.aq_misrep           || '',
    d.aq_security         || '',
    d.situation_explanation || '',
    d.sig_name            || '',
    d.sig_date            || '',
    folderUrl,
  ];
  tab.appendRow(row);
}

// ════════════════════════════════════════════════════════════════════
// SUMMARY TEXT — Work Permit
// ════════════════════════════════════════════════════════════════════
function buildWorkPermitSummary(d) {
  const lines = [];

  lines.push('WORK PERMIT INTAKE — JSR IMMIGRATION LTD.');
  lines.push('Submitted: ' + (d.submission_timestamp || ''));
  lines.push('');

  lines.push('--- PERSONAL INFORMATION ---');
  lines.push('Full name:          ' + (d.full_name       || ''));
  lines.push('Date of birth:      ' + (d.dob             || ''));
  lines.push('Country of birth:   ' + (d.country_birth   || ''));
  lines.push('Citizenship:        ' + (d.citizenship     || ''));
  lines.push('UCI:                ' + (d.uci             || ''));
  lines.push('');

  lines.push('--- IMMIGRATION STATUS ---');
  lines.push('Current status:     ' + (d.current_status || ''));
  lines.push('Valid from:         ' + (d.status_from    || ''));
  lines.push('Valid to:           ' + (d.status_to      || ''));
  lines.push('');

  lines.push('--- PASSPORT ---');
  lines.push('Number:             ' + (d.passport_number  || ''));
  lines.push('Country of issue:   ' + (d.passport_country || ''));
  lines.push('Issue date:         ' + (d.passport_issued  || ''));
  lines.push('Expiry date:        ' + (d.passport_expiry  || ''));
  lines.push('');

  lines.push('--- MARITAL STATUS ---');
  lines.push('Status:             ' + (d.marital_status    || ''));
  lines.push('Partner first name: ' + (d.partner_first     || ''));
  lines.push('Partner last name:  ' + (d.partner_last      || ''));
  lines.push('Partner DOB:        ' + (d.partner_dob       || ''));
  lines.push('Date of marriage:   ' + (d.date_of_marriage  || ''));
  lines.push('Date of separation: ' + (d.date_of_separation|| ''));
  lines.push('');

  lines.push('--- CONTACT & ADDRESS ---');
  lines.push('Phone:              ' + (d.phone           || ''));
  lines.push('Email:              ' + (d.email           || ''));
  lines.push('Alt phone:          ' + (d.alt_phone       || ''));
  lines.push('Address:            ' + (d.address_street  || ''));
  lines.push('City:               ' + (d.address_city    || ''));
  lines.push('Province:           ' + (d.address_province|| ''));
  lines.push('Postal code:        ' + (d.postal_code     || ''));
  lines.push('');

  lines.push('--- ENTRY TO CANADA ---');
  lines.push('First entry date:   ' + (d.first_entry_date  || ''));
  lines.push('First entry place:  ' + (d.first_entry_place || ''));
  lines.push('Recent entry date:  ' + (d.recent_entry_date || ''));
  lines.push('Recent entry place: ' + (d.recent_entry_place|| ''));
  lines.push('');

  lines.push('--- WORK PERMIT CATEGORY ---');
  lines.push('Category:           ' + (d.wp_category       || ''));
  if (d.wp_category_other) {
    lines.push('Category (other):   ' + d.wp_category_other);
  }
  lines.push('Employer:           ' + (d.wp_employer    || ''));
  lines.push('Job title:          ' + (d.wp_job_title   || ''));
  lines.push('NOC code:           ' + (d.wp_noc         || ''));
  lines.push('LMIA number:        ' + (d.wp_lmia_number || ''));
  lines.push('Job offer:          ' + (d.job_offer      || ''));
  lines.push('');

  lines.push('--- EMPLOYMENT HISTORY ---');
  let jobIdx = 1;
  while (d['emp_name_' + jobIdx] || d['emp_title_' + jobIdx]) {
    lines.push('Job ' + jobIdx + ':');
    lines.push('  Employer:   ' + (d['emp_name_'  + jobIdx] || ''));
    lines.push('  Title:      ' + (d['emp_title_' + jobIdx] || ''));
    lines.push('  Type:       ' + (d['emp_type_'  + jobIdx] || ''));
    lines.push('  Start:      ' + (d['emp_start_' + jobIdx] || ''));
    lines.push('  End:        ' + (d['emp_end_'   + jobIdx] || ''));
    lines.push('  Address:    ' + (d['emp_addr_'  + jobIdx] || ''));
    lines.push('  Reason:     ' + (d['emp_reason_'+ jobIdx] || ''));
    jobIdx++;
  }
  lines.push('');

  lines.push('--- ADMISSIBILITY ---');
  const aqKeys = [
    ['aq_criminal',    'Criminal record'],
    ['aq_refusal',     'Refusal'],
    ['aq_prev_canada', 'Previous Canada application'],
    ['aq_removal',     'Removal order'],
    ['aq_overstay',    'Overstay'],
    ['aq_tb',          'TB / Medical'],
    ['aq_pending',     'Pending application'],
    ['aq_misrep',      'Misrepresentation'],
    ['aq_security',    'Security / terrorism'],
  ];
  aqKeys.forEach(([key, label]) => {
    const ans    = d[key]           || '';
    const detail = d[key + '_detail'] || '';
    lines.push(label + ': ' + ans + (detail ? ' — ' + detail : ''));
  });
  lines.push('');

  lines.push('--- YOUR SITUATION & GOALS ---');
  lines.push(d.situation_explanation || '(not provided)');
  lines.push('');

  lines.push('--- DECLARATION ---');
  lines.push('Signed name: ' + (d.sig_name || ''));
  lines.push('Signed date: ' + (d.sig_date || ''));
  lines.push('');
  lines.push('JSR Immigration Ltd. · RCIC #R712841 · info@jsrimmigration.com');

  return lines.join('\n');
}

// ════════════════════════════════════════════════════════════════════
// EMAIL DRAFT — Work Permit
// ════════════════════════════════════════════════════════════════════
function draftWorkPermitEmail(name, clientEmail, category, situation, summary, folderUrl, dateStr) {
  const subject  = 'WP Intake — ' + name + ' — ' + dateStr;
  const truncSit = situation.length > MAX_EMAIL_SITUATION_LENGTH
    ? situation.slice(0, MAX_EMAIL_SITUATION_LENGTH) + '…'
    : situation;
  const body =
    'New Work Permit intake received.\n\n' +
    'Client:       ' + name        + '\n' +
    'Email:        ' + clientEmail + '\n' +
    'Category:     ' + category    + '\n' +
    'Date:         ' + dateStr     + '\n' +
    'Drive folder: ' + folderUrl   + '\n\n' +
    '--- CLIENT SITUATION ---\n' +
    truncSit + '\n\n' +
    '--- FULL INTAKE SUMMARY ---\n' +
    summary;

  GmailApp.createDraft(RCIC_EMAIL, subject, body);
}

// ════════════════════════════════════════════════════════════════════
// PGWP HANDLER  (unchanged logic — included for completeness)
// ════════════════════════════════════════════════════════════════════
function handlePgwp(d) {
  const name    = d.full_name || 'Unknown';
  const email   = d.email    || '';
  const dateStr = (d.submission_timestamp || new Date().toISOString()).slice(0, 10);

  const parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const clientFolder = parentFolder.createFolder('PGWP — ' + name + ' — ' + dateStr);
  const folderUrl    = clientFolder.getUrl();

  const summary = buildPgwpSummary(d);
  clientFolder.createFile(buildPdf('PGWP Intake — ' + name, summary));
  logPgwpRow(d, folderUrl);
  draftGenericEmail('PGWP', name, email, summary, folderUrl, dateStr);

  return jsonResponse('ok', 'PGWP intake saved.');
}

function buildPgwpSummary(d) {
  const lines = [];
  lines.push('PGWP INTAKE — JSR IMMIGRATION LTD.');
  lines.push('Submitted: ' + (d.submission_timestamp || ''));
  lines.push('');
  lines.push('Full name:        ' + (d.full_name       || ''));
  lines.push('DOB:              ' + (d.dob             || ''));
  lines.push('Country of birth: ' + (d.country_birth   || ''));
  lines.push('Citizenship:      ' + (d.citizenship     || ''));
  lines.push('UCI:              ' + (d.uci             || ''));
  lines.push('');
  lines.push('Status:           ' + (d.current_status  || ''));
  lines.push('Status valid to:  ' + (d.status_to       || ''));
  lines.push('');
  lines.push('Passport number:  ' + (d.passport_number  || ''));
  lines.push('Passport country: ' + (d.passport_country || ''));
  lines.push('Passport expiry:  ' + (d.passport_expiry  || ''));
  lines.push('');
  lines.push('Phone: ' + (d.phone || '') + '  Email: ' + (d.email || ''));
  lines.push('Address: ' + [d.address_street, d.address_city, d.address_province, d.postal_code].filter(Boolean).join(', '));
  lines.push('');
  // Education
  let eduIdx = 1;
  while (d['edu_name_' + eduIdx]) {
    lines.push('Education ' + eduIdx + ': ' + d['edu_name_' + eduIdx] + ' — ' + d['edu_program_' + eduIdx] + ' (' + d['edu_credential_' + eduIdx] + ')');
    lines.push('  ' + (d['edu_start_' + eduIdx] || '') + ' to ' + (d['edu_end_' + eduIdx] || ''));
    eduIdx++;
  }
  lines.push('');
  lines.push('Signed: ' + (d.sig_name || '') + '  Date: ' + (d.sig_date || ''));
  lines.push('');
  lines.push('JSR Immigration Ltd. · RCIC #R712841');
  return lines.join('\n');
}

function logPgwpRow(d, folderUrl) {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tab = getOrCreateSheet(ss, 'PGWP');
  if (tab.getLastRow() === 0) {
    const headers = ['Submitted At','Full Name','DOB','Country of Birth','Citizenship','UCI',
      'Current Status','Status Valid To','Passport Number','Passport Country','Passport Expiry',
      'Marital Status','Phone','Email','Address','City','Province','Postal Code',
      'First Entry Date','First Entry Place','Recent Entry Date','Recent Entry Place',
      'Education (summary)','Employment (summary)','Signed Name','Signed Date','Drive Folder URL'];
    tab.appendRow(headers);
    tab.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    tab.setFrozenRows(1);
  }
  const eduSummary = buildRepeatSummary(d, 'edu_name_', 5);
  const empSummary = buildRepeatSummary(d, 'emp_name_', 5);
  tab.appendRow([
    new Date(), d.full_name||'', d.dob||'', d.country_birth||'', d.citizenship||'', d.uci||'',
    d.current_status||'', d.status_to||'', d.passport_number||'', d.passport_country||'', d.passport_expiry||'',
    d.marital_status||'', d.phone||'', d.email||'', d.address_street||'', d.address_city||'', d.address_province||'', d.postal_code||'',
    d.first_entry_date||'', d.first_entry_place||'', d.recent_entry_date||'', d.recent_entry_place||'',
    eduSummary, empSummary, d.sig_name||'', d.sig_date||'', folderUrl,
  ]);
}

// ════════════════════════════════════════════════════════════════════
// VISITOR VISA HANDLER
// ════════════════════════════════════════════════════════════════════
function handleVisitorVisa(d) {
  const name    = d.full_name || 'Unknown';
  const email   = d.email    || '';
  const dateStr = (d.submission_timestamp || new Date().toISOString()).slice(0, 10);

  const parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const clientFolder = parentFolder.createFolder('TRV — ' + name + ' — ' + dateStr);
  const folderUrl    = clientFolder.getUrl();

  const summary = buildGenericSummary('VISITOR VISA (TRV) INTAKE', d);
  clientFolder.createFile(buildPdf('Visitor Visa Intake — ' + name, summary));
  logGenericRow(d, folderUrl, 'Visitor Visa', [
    'Submitted At','Full Name','DOB','Citizenship','Phone','Email','Address','City','Province','Postal Code',
    'Passport Number','Passport Expiry','Purpose of Visit','Travel History','Funds Available','Signed Name','Signed Date','Drive Folder URL'
  ], [
    new Date(), d.full_name||'', d.dob||'', d.citizenship||'', d.phone||'', d.email||'',
    d.address_street||'', d.address_city||'', d.address_province||'', d.postal_code||'',
    d.passport_number||'', d.passport_expiry||'', d.visit_purpose||'', d.travel_history||'', d.funds_available||'',
    d.sig_name||'', d.sig_date||'', folderUrl,
  ]);
  draftGenericEmail('Visitor Visa (TRV)', name, email, summary, folderUrl, dateStr);
  return jsonResponse('ok', 'Visitor visa intake saved.');
}

// ════════════════════════════════════════════════════════════════════
// VISITOR RECORD HANDLER
// ════════════════════════════════════════════════════════════════════
function handleVisitorRecord(d) {
  const name    = d.full_name || 'Unknown';
  const email   = d.email    || '';
  const dateStr = (d.submission_timestamp || new Date().toISOString()).slice(0, 10);

  const parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const clientFolder = parentFolder.createFolder('VR — ' + name + ' — ' + dateStr);
  const folderUrl    = clientFolder.getUrl();

  const summary = buildGenericSummary('VISITOR RECORD INTAKE', d);
  clientFolder.createFile(buildPdf('Visitor Record Intake — ' + name, summary));
  logGenericRow(d, folderUrl, 'Visitor Record', [
    'Submitted At','Full Name','DOB','Citizenship','Phone','Email','Address','City','Province','Postal Code',
    'Current Status','Status Valid To','Passport Number','Passport Expiry','Signed Name','Signed Date','Drive Folder URL'
  ], [
    new Date(), d.full_name||'', d.dob||'', d.citizenship||'', d.phone||'', d.email||'',
    d.address_street||'', d.address_city||'', d.address_province||'', d.postal_code||'',
    d.current_status||'', d.status_to||'', d.passport_number||'', d.passport_expiry||'',
    d.sig_name||'', d.sig_date||'', folderUrl,
  ]);
  draftGenericEmail('Visitor Record', name, email, summary, folderUrl, dateStr);
  return jsonResponse('ok', 'Visitor record intake saved.');
}

// ════════════════════════════════════════════════════════════════════
// STUDY EXTENSION HANDLER
// ════════════════════════════════════════════════════════════════════
function handleStudyExtension(d) {
  const name    = d.full_name || 'Unknown';
  const email   = d.email    || '';
  const dateStr = (d.submission_timestamp || new Date().toISOString()).slice(0, 10);

  const parentFolder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const clientFolder = parentFolder.createFolder('SP-Ext — ' + name + ' — ' + dateStr);
  const folderUrl    = clientFolder.getUrl();

  const summary = buildGenericSummary('STUDY PERMIT EXTENSION INTAKE', d);
  clientFolder.createFile(buildPdf('Study Extension Intake — ' + name, summary));
  logGenericRow(d, folderUrl, 'Study Extension', [
    'Submitted At','Full Name','DOB','Citizenship','Phone','Email','Address','City','Province','Postal Code',
    'Current Status','Status Valid To','Passport Number','Passport Expiry',
    'Institution','DLI','Program','Course Start','Course End','Tuition Fees','Funds Available',
    'Signed Name','Signed Date','Drive Folder URL'
  ], [
    new Date(), d.full_name||'', d.dob||'', d.citizenship||'', d.phone||'', d.email||'',
    d.address_street||'', d.address_city||'', d.address_province||'', d.postal_code||'',
    d.current_status||'', d.status_to||'', d.passport_number||'', d.passport_expiry||'',
    d.current_institution||'', d.dli_number||'', d.current_program||'',
    d.course_start||'', d.course_end||'', d.tuition_fees||'', d.funds_available||'',
    d.sig_name||'', d.sig_date||'', folderUrl,
  ]);
  draftGenericEmail('Study Permit Extension', name, email, summary, folderUrl, dateStr);
  return jsonResponse('ok', 'Study extension intake saved.');
}

// ════════════════════════════════════════════════════════════════════
// SHARED HELPERS
// ════════════════════════════════════════════════════════════════════

/**
 * Returns an existing sheet by name, or creates one at the end.
 */
function getOrCreateSheet(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

/**
 * Builds a plain-text PDF blob from a title and body string.
 */
function buildPdf(title, body) {
  const html   = '<html><head><meta charset="UTF-8"><style>body{font-family:monospace;font-size:11pt;white-space:pre-wrap;padding:24px;}</style></head><body>' +
                 escapeHtml(title + '\n\n' + body) +
                 '</body></html>';
  const blob   = Utilities.newBlob(html, 'text/html', title + '.html');
  const pdfBlob = blob.getAs('application/pdf');
  pdfBlob.setName(title + '.pdf');
  return pdfBlob;
}

function escapeHtml(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

/**
 * Generic summary builder used for simpler form types.
 */
function buildGenericSummary(title, d) {
  const lines = [title, 'Submitted: ' + (d.submission_timestamp || ''), ''];
  Object.keys(d).forEach(k => {
    if (k === 'signature_image' || k === 'form_type') return;
    lines.push(k.replace(/_/g, ' ') + ': ' + (d[k] || ''));
  });
  lines.push('');
  lines.push('JSR Immigration Ltd. · RCIC #R712841');
  return lines.join('\n');
}

/**
 * Generic row logger — creates headers on first write.
 */
function logGenericRow(d, folderUrl, sheetName, headers, values) {
  const ss  = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tab = getOrCreateSheet(ss, sheetName);
  if (tab.getLastRow() === 0) {
    tab.appendRow(headers);
    tab.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#1a3a5c').setFontColor('#ffffff');
    tab.setFrozenRows(1);
  }
  tab.appendRow(values);
}

/**
 * Generic email drafter.
 */
function draftGenericEmail(formLabel, name, clientEmail, summary, folderUrl, dateStr) {
  const subject = formLabel + ' Intake — ' + name + ' — ' + dateStr;
  const body =
    'New ' + formLabel + ' intake received.\n\n' +
    'Client:       ' + name        + '\n' +
    'Email:        ' + clientEmail + '\n' +
    'Date:         ' + dateStr     + '\n' +
    'Drive folder: ' + folderUrl   + '\n\n' +
    '--- FULL INTAKE SUMMARY ---\n' +
    summary;
  GmailApp.createDraft(RCIC_EMAIL, subject, body);
}

/**
 * Builds a comma-separated list of up to `max` repeater values.
 */
function buildRepeatSummary(d, prefix, max) {
  const vals = [];
  for (let i = 1; i <= max; i++) {
    const v = d[prefix + i];
    if (v) vals.push(v);
  }
  return vals.join(', ');
}

/**
 * Returns a JSON ContentService response.
 */
function jsonResponse(status, message) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: status, message: message }))
    .setMimeType(ContentService.MimeType.JSON);
}
