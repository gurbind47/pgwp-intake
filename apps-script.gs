/**
 * ═══════════════════════════════════════════════════════════════════
 * JSR Immigration Ltd. — Multi-form Intake → Sheet + PDF + Gmail Draft
 * Google Apps Script Web App  ·  one project, two intake forms
 * ═══════════════════════════════════════════════════════════════════
 *
 * Supported form types (sent in the POST payload as `form_type`):
 *   • "pgwp"             — Post-Graduate Work Permit intake    → tab "Responses"
 *   • "visitor_record"   — Visitor Record extension intake      → tab "Visitor Records"
 *   • "study_extension"  — Study Permit extension intake        → tab "Study Extensions"
 *   • "visitor_visa"     — Visitor Visa (TRV / Super Visa)       → tab "Visitor Visas"
 *   • "visitor_visa_inside_canada" — Visitor Visa (Inside Canada) → tab "Visitor Visas (Inside Canada)"
 *
 * On every submission this script:
 *   1. Appends a row to the matching tab in the Google Sheet.
 *   2. Creates a folder inside the parent Drive folder
 *        – PGWP             → "DRAFT - <Client Name>"
 *        – Visitor          → "DRAFT - VR - <Client Name>"
 *        – Study Extension  → "DRAFT - SPE - <Client Name>"
 *        – Visitor Visa     → "DRAFT - TRV - <Client Name>"
 *        – TRV Inside Canada→ "DRAFT - TRV-IC - <Client Name>"
 *   3. Generates a printable PDF of the submission inside that folder.
 *   4. Saves the drawn signature as a PNG inside that folder (if any).
 *   5. Drafts a Gmail email (in info@jsrimmigration.com) addressed to
 *      the client with the PDF attached, ready for review and send.
 *
 * SETUP (do once, in order):
 *   1. Paste this entire file into Extensions → Apps Script. Save (⌘S).
 *   2. Function dropdown → "setupAllSheets"  → ▶ Run → authorize.
 *      → "Responses" + "Visitor Records" + "Study Extensions" + "Visitor Visas" + "Visitor Visas (Inside Canada)" tabs appear with header rows.
 *   3. Function dropdown → "testPipeline"                → ▶ Run → PGWP folder + PDF + Gmail draft created.
 *   4. Function dropdown → "testVisitorPipeline"         → ▶ Run → VR folder + PDF + Gmail draft created.
 *   5. Function dropdown → "testStudyExtensionPipeline"  → ▶ Run → SPE folder + PDF + Gmail draft created.
 *   6. Function dropdown → "testVisitorVisaPipeline"     → ▶ Run → TRV folder + PDF + Gmail draft created.
 *   7. Function dropdown → "testVisitorVisaIcPipeline"   → ▶ Run → TRV-IC folder + PDF + Gmail draft created.
 *   8. Deploy → New deployment → ⚙ → Web app
 *      Execute as: Me   ·   Who has access: Anyone   →   Deploy.
 *   9. Copy Web App URL → paste into all active intake index.html files (APPS_SCRIPT_URL).
 *
 * RE-DEPLOYING after edits:
 *      Deploy → Manage deployments → ✏ → Version: New version → Deploy.
 *      The URL stays the same. If you skip this, the live web app keeps
 *      running the OLD code.
 * ═══════════════════════════════════════════════════════════════════
 */

// ── GLOBAL CONFIG ───────────────────────────────────────────────────
const SPREADSHEET_ID         = '1mQbwitzTiMWQPN1lTwYSXyw2pT9HK8CplC5O3Ur8mMQ';
const DRAFT_PARENT_FOLDER_ID = '1fWnLXQmMfAfVF-H0jou-uK9UQTEWLjiI';   // shared by both forms
const FIRM_NAME              = 'JSR Immigration Ltd.';
const FIRM_RCIC              = 'RCIC #R712841';
const FIRM_EMAIL             = 'info@jsrimmigration.com';

// ════════════════════════════════════════════════════════════════════
// HEADERS — one array per form (column order must match buildXxxRow)
// ════════════════════════════════════════════════════════════════════
const PGWP_HEADERS = [
  'Submitted',
  'Full Name','UCI','DOB','Country of Birth','Citizenship',
  'Current Status','Status From','Status To',
  'Passport #','Passport Country','Passport Issued','Passport Expires',
  'Marital Status','Partner Name','Partner DOB','Marriage Date','Separation Date',
  'Phone','Email','Alt Phone','Street','City','Province','Postal',
  'First Entry Date','First Entry Port','Recent Entry Date','Recent Entry Port',
  'Education History','Employment History',
  'AQ Criminal','AQ Criminal Detail',
  'AQ Refusal','AQ Refusal Detail',
  'AQ Prior Canada','AQ Prior Canada Detail',
  'AQ Removal','AQ Removal Detail',
  'AQ Overstay','AQ Overstay Detail',
  'AQ Medical','AQ Medical Detail',
  'AQ Pending','AQ Pending Detail',
  'AQ Misrep','AQ Misrep Detail',
  'AQ Security','AQ Security Detail',
  'AQ Study Permit','AQ Study Permit Detail',
  'Work Permit Type',
  'Program (Application)',
  'Background — What They Did',
  'Reason for Applying',
  'Work Permit (Own Words)',
  'Signature','Signature Date','Consent Accuracy','Consent Auth','Drawn Signature',
  'Client Folder URL','Client PDF URL'
];

const VISITOR_HEADERS = [
  'Submitted','Form Type',
  'Full Name','UCI','DOB','Country of Birth','Citizenship',
  'Phone','Email','Alt Phone','Street','City','Province','Postal',
  'Marital Status','Partner Name','Partner DOB','Marriage Date','Separation Date',
  'Current Status','Status To','Status Expired (Restoration?)',
  'Inviter Present','Inviter Name','Inviter Address','Inviter Email','Inviter Phone',
  'Inviter Employed','Inviter Job Title','Inviter Bank Balance',
  'Original Entry Date','Original Entry Place','Recent Entry Date','Recent Entry Place',
  'Highest Education','Program','School Name','School Address','Edu Start','Edu End','Additional Education',
  'Has Work Experience','Occupation','Company','Work Start','Work End','Work Address','Additional Jobs',
  'AQ Criminal','AQ Criminal Detail',
  'AQ Refusal','AQ Refusal Detail',
  'AQ Removal','AQ Removal Detail',
  'AQ Overstay','AQ Overstay Detail',
  'AQ Medical','AQ Medical Detail',
  'AQ Canadian Status (Q33)','AQ Canadian Status Detail',
  'Doc Financial','Doc Passport','Doc Status','Doc Photo','Doc Edu/Emp',
  'Doc Inviter','Doc Medical','Doc Marriage',
  'Signature','Signature Date','Consent Accuracy','Consent Auth','Drawn Signature',
  'Client Folder URL','Client PDF URL'
];

const STUDY_EXTENSION_HEADERS = [
  'Submitted','Form Type',
  'Full Name','UCI','DOB','Country of Birth','Citizenship',
  'Phone','Email','Alt Phone','Street','City','Province','Postal',
  'Marital Status','Partner Name','Partner DOB','Marriage Date','Separation Date',
  'Previously Married','Previous Spouse Name',
  'Current Status','Status To','Status Expired (Restoration?)',
  'Original Entry Date','Original Entry Place','Purpose of First Entry',
  'Recent Entry Date','Recent Entry Place',
  'Current Institution','DLI Number','Program (Current)','Credential',
  'LOA Received','Course Start','Course End',
  'Tuition Fees (CAD)','Funds Available (CAD)','Who Pays Fees','Payer Details',
  'Highest Past Education','Past Program','Past School','Past School Address',
  'Past Edu Start','Past Edu End','Additional Education',
  'AQ Prior Applied','AQ Prior Applied Detail',
  'AQ Refusal','AQ Refusal Detail',
  'AQ Criminal','AQ Criminal Detail',
  'AQ Political','AQ Political Detail',
  'AQ Military','AQ Military Detail',
  'AQ Medical','AQ Medical Detail',
  'Explain Yes Answers',
  'Doc Passport','Doc Current Status','Doc Marriage','Doc LOA',
  'Doc Parents','Doc Transcripts','Doc Funds','Doc Photo','Doc Medical',
  'Signature','Signature Date','Consent Accuracy','Consent Auth','Drawn Signature',
  'Client Folder URL','Client PDF URL'
];

const VISITOR_VISA_HEADERS = [
  'Submitted','Form Type','Application Type',
  'Full Name','UCI','DOB','Sex','Native Language',
  'Country of Birth','City of Birth','Citizenship','Country of Residence',
  'Passport #','Passport Country','Passport Issued','Passport Expires',
  'Prior Passport','Prior Passport Detail',
  'Phone','Email','Alt Phone','Street','City','Province/State','Home Country','Postal',
  'Marital Status','Partner Name','Partner DOB','Marriage Date','Separation Date',
  'Previously Married','Previous Spouse',
  'Visit Purpose','Visit Purpose Detail','Intended Arrival','Intended Departure',
  'Cities to Visit','Prior Canada','Prior Canada Detail',
  'Has Host','Host Name','Host Relationship','Host Status',
  'Host Address','Host Email','Host Phone','Host Occupation','Host Household Size','Host Income',
  'Who Pays','Payer Details','Applicant Occupation','Applicant Employer','Education History','Employment History',
  'Annual Income','Savings Amount','Funds Source',
  'Father Name','Father DOB','Father Country','Father Occupation',
  'Mother Name','Mother DOB','Mother Country','Mother Occupation',
  'Children','Siblings','Family in Canada','Canada Family List',
  'Has Travel History','Travel History','Has Prior Refusal','Prior Refusals',
  'AQ Criminal','AQ Criminal Detail',
  'AQ Refusal','AQ Refusal Detail',
  'AQ Overstay','AQ Overstay Detail',
  'AQ Medical','AQ Medical Detail',
  'AQ Military','AQ Military Detail',
  'AQ Misrepresentation','AQ Misrepresentation Detail',
  'SV Host LICO Ack','SV Medical Insurance Ack','SV Invitation Letter Ack','SV Host Relation',
  'Doc Passport','Doc Photo','Doc Funds','Doc Itinerary','Doc Employment','Doc Ties',
  'Doc Invitation','Doc Host Status','Doc Host Income','Doc Medical Insurance','Doc Marriage',
  'Signature','Signature Date','Consent Accuracy','Consent Auth','Drawn Signature',
  'Client Folder URL','Client PDF URL'
];

const VISITOR_VISA_IC_HEADERS = [
  'Submitted','Form Type',
  'Full Name','UCI','DOB','Sex','Country of Birth','Citizenship',
  'Phone','Email','Street','City','Province','Postal',
  'Marital Status','Partner Name','Partner DOB','Marriage Date','Separation Date',
  'Previously Married','Previous Spouse',
  'Current Status','Status From','Status To','Status Document Number','DLI or Employer',
  'Address History','Employment History','Education History',
  'Savings Amount','Funds Source','Who Pays','Payer Details',
  'AQ Criminal','AQ Criminal Detail',
  'AQ Refusal','AQ Refusal Detail',
  'AQ Overstay','AQ Overstay Detail',
  'AQ Medical','AQ Medical Detail',
  'AQ Military','AQ Military Detail',
  'AQ Misrepresentation','AQ Misrepresentation Detail',
  'Doc Passport','Doc Photo','Doc Funds','Doc Itinerary','Doc Employment','Doc Ties',
  'Doc Invitation','Doc Host Status','Doc Host Income','Doc Medical Insurance','Doc Marriage',
  'Signature','Signature Date','Consent Accuracy','Consent Auth','Drawn Signature',
  'Client Folder URL','Client PDF URL'
];

// ════════════════════════════════════════════════════════════════════
// FORM_CONFIGS — one entry per supported form_type
// Each entry is everything doPost / appendRow / generateClientArtifacts
// need to know to handle that form. Add a third form by adding another
// entry below — no other code changes required.
// ════════════════════════════════════════════════════════════════════
const FORM_CONFIGS = {
  pgwp: {
    label:           'PGWP Intake',
    sheetName:       'Responses',
    headers:         PGWP_HEADERS,
    headerColor:     '#1a3a5c',
    folderPrefix:    'DRAFT',
    pdfFilePrefix:   'PGWP Intake',
    buildRow:        function(d, fUrl, pUrl) { return buildPgwpRow(d, fUrl, pUrl); },
    buildPdfHtml:    function(d)             { return buildPgwpPdfHtml(d); },
    buildEmailHtml:  function(d, url)        { return buildPgwpEmailHtml(d, url); },
    buildEmailPlain: function(d, url)        { return buildPgwpEmailPlain(d, url); },
    emailSubject:    function(name)          { return 'Please confirm your Work Permit intake details — ' + name; },
    sampleData:      function()              { return sampleDataPgwp(); },
    aqColumns:       [32, 34, 36, 38, 40, 42, 44, 46, 48, 50]
  },
  visitor_record: {
    label:           'Visitor Record Extension',
    sheetName:       'Visitor Records',
    headers:         VISITOR_HEADERS,
    headerColor:     '#1a3a5c',
    folderPrefix:    'DRAFT - VR',
    pdfFilePrefix:   'Visitor Record Intake',
    buildRow:        function(d, fUrl, pUrl) { return buildVisitorRow(d, fUrl, pUrl); },
    buildPdfHtml:    function(d)             { return buildVisitorPdfHtml(d); },
    buildEmailHtml:  function(d, url)        { return buildVisitorEmailHtml(d, url); },
    buildEmailPlain: function(d, url)        { return buildVisitorEmailPlain(d, url); },
    emailSubject:    function(name)          { return 'Please confirm your Visitor Record extension intake — ' + name; },
    sampleData:      function()              { return sampleDataVisitor(); },
    // 1-based indices of "AQ ... " answer columns (Yes triggers amber highlight)
    aqColumns:       [49, 51, 53, 55, 57, 59]
  },
  study_extension: {
    label:           'Study Permit Extension',
    sheetName:       'Study Extensions',
    headers:         STUDY_EXTENSION_HEADERS,
    headerColor:     '#1a3a5c',
    folderPrefix:    'DRAFT - SPE',
    pdfFilePrefix:   'Study Permit Extension Intake',
    buildRow:        function(d, fUrl, pUrl) { return buildStudyExtensionRow(d, fUrl, pUrl); },
    buildPdfHtml:    function(d)             { return buildStudyExtensionPdfHtml(d); },
    buildEmailHtml:  function(d, url)        { return buildStudyExtensionEmailHtml(d, url); },
    buildEmailPlain: function(d, url)        { return buildStudyExtensionEmailPlain(d, url); },
    emailSubject:    function(name)          { return 'Please confirm your Study Permit extension intake — ' + name; },
    sampleData:      function()              { return sampleDataStudyExtension(); },
    // 1-based indices of admissibility answer columns that warrant amber highlight on "Yes"
    // Skipping AQ Prior Applied (col 48) — a "Yes" there is routine history, not a flag.
    aqColumns:       [50, 52, 54, 56, 58]
  },
  visitor_visa: {
    label:           'Visitor Visa (TRV) Intake',
    sheetName:       'Visitor Visas',
    headers:         VISITOR_VISA_HEADERS,
    headerColor:     '#1a3a5c',
    folderPrefix:    'DRAFT - TRV',
    pdfFilePrefix:   'Visitor Visa Intake',
    buildRow:        function(d, fUrl, pUrl) { return buildVisitorVisaRow(d, fUrl, pUrl); },
    buildPdfHtml:    function(d)             { return buildVisitorVisaPdfHtml(d); },
    buildEmailHtml:  function(d, url)        { return buildVisitorVisaEmailHtml(d, url); },
    buildEmailPlain: function(d, url)        { return buildVisitorVisaEmailPlain(d, url); },
    emailSubject:    function(name)          { return 'Please confirm your Visitor Visa intake — ' + name; },
    sampleData:      function()              { return sampleDataVisitorVisa(); },
    // 1-based indices of the 6 AQ answer columns (Criminal, Refusal, Overstay, Medical, Military, Misrep)
    aqColumns:       [76, 78, 80, 82, 84, 86]
  },
  visitor_visa_inside_canada: {
    label:           'Visitor Visa (Inside Canada) Intake',
    sheetName:       'Visitor Visas (Inside Canada)',
    headers:         VISITOR_VISA_IC_HEADERS,
    headerColor:     '#1a3a5c',
    folderPrefix:    'DRAFT - TRV-IC',
    pdfFilePrefix:   'Visitor Visa Inside Canada Intake',
    buildRow:        function(d, fUrl, pUrl) { return buildVisitorVisaIcRow(d, fUrl, pUrl); },
    buildPdfHtml:    function(d)             { return buildVisitorVisaIcPdfHtml(d); },
    buildEmailHtml:  function(d, url)        { return buildVisitorVisaIcEmailHtml(d, url); },
    buildEmailPlain: function(d, url)        { return buildVisitorVisaIcEmailPlain(d, url); },
    emailSubject:    function(name)          { return 'Please confirm your TRV (Inside Canada) intake — ' + name; },
    sampleData:      function()              { return sampleDataVisitorVisaIc(); },
    aqColumns:       [34, 36, 38, 40, 42, 44]
  }
};

function getConfig(formType) {
  return FORM_CONFIGS[formType] || FORM_CONFIGS.pgwp;   // backward-compat default
}

// ════════════════════════════════════════════════════════════════════
// ENTRY POINTS
// ════════════════════════════════════════════════════════════════════
function doGet() {
  return json({ status: 'ok', app: 'JSR Multi-Form Intake' });
}

function doPost(e) {
  try {
    let raw = '';
    if (e && e.parameter && e.parameter.payload) {
      raw = e.parameter.payload;
    } else if (e && e.postData && e.postData.contents) {
      raw = e.postData.contents;
    }
    const data = raw ? JSON.parse(raw) : {};
    const cfg  = getConfig(data.form_type);

    // 1. Drive folder + PDF + Gmail draft (do BEFORE row append so URLs go in row)
    let folderUrl = '', pdfUrl = '';
    try {
      const result = generateClientArtifacts(data, cfg);
      folderUrl = result.folderUrl;
      pdfUrl    = result.pdfUrl;
    } catch (artifactErr) {
      Logger.log('[' + cfg.label + '] artifact step failed (sheet row will still be written): ' + artifactErr);
    }

    // 2. Always append to sheet (even if artifact step failed)
    appendRow(data, cfg, folderUrl, pdfUrl);

    return json({ status: 'ok', form_type: data.form_type || 'pgwp', folderUrl: folderUrl, pdfUrl: pdfUrl });
  } catch (err) {
    Logger.log('doPost error: ' + err);
    return json({ status: 'error', message: String(err) });
  }
}

// ════════════════════════════════════════════════════════════════════
// MANUAL EDITOR FUNCTIONS  (run from Apps Script editor)
// ════════════════════════════════════════════════════════════════════
function setupSheet()                { setupSheetForConfig(getConfig('pgwp')); }
function setupVisitorSheet()         { setupSheetForConfig(getConfig('visitor_record')); }
function setupStudyExtensionSheet()  { setupSheetForConfig(getConfig('study_extension')); }
function setupVisitorVisaSheet()     { setupSheetForConfig(getConfig('visitor_visa')); }
function setupVisitorVisaIcSheet()   { setupSheetForConfig(getConfig('visitor_visa_inside_canada')); }
function setupAllSheets()            { setupSheet(); setupVisitorSheet(); setupStudyExtensionSheet(); setupVisitorVisaSheet(); setupVisitorVisaIcSheet(); }

/**
 * Safely adds the 5 new Application Background columns to the existing
 * "Responses" sheet WITHOUT clearing any existing data.
 * Run this ONCE after pasting the updated apps-script code.
 */
function addWorkPermitColumnsToSheet() {
  const ss      = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet   = ss.getSheetByName('Responses');
  if (!sheet) { Logger.log('ERROR: "Responses" sheet not found.'); return; }
  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const sigIdx  = headers.indexOf('Signature'); // 0-based
  if (sigIdx < 0) { Logger.log('ERROR: "Signature" column not found — sheet may already be updated, or headers differ.'); return; }
  const newCols = [
    'Work Permit Type', 'Program (Application)',
    'Background — What They Did', 'Reason for Applying', 'Work Permit (Own Words)'
  ];
  sheet.insertColumnsBefore(sigIdx + 1, newCols.length); // sigIdx is 0-based; sheet cols are 1-based
  sheet.getRange(1, sigIdx + 1, 1, newCols.length)
    .setValues([newCols])
    .setBackground('#1a3a5c').setFontColor('#ffffff').setFontWeight('bold');
  sheet.autoResizeColumns(sigIdx + 1, newCols.length);
  Logger.log('✓ 5 new columns added before "Signature" column in "Responses" sheet.');
}

function testWrite()                 { appendRow(sampleDataPgwp(),           getConfig('pgwp'),            '', ''); Logger.log('✓ PGWP test row written.'); }
function testVisitorWrite()          { appendRow(sampleDataVisitor(),        getConfig('visitor_record'),  '', ''); Logger.log('✓ Visitor test row written.'); }
function testStudyExtensionWrite()   { appendRow(sampleDataStudyExtension(), getConfig('study_extension'), '', ''); Logger.log('✓ Study Extension test row written.'); }
function testVisitorVisaWrite()      { appendRow(sampleDataVisitorVisa(),    getConfig('visitor_visa'),    '', ''); Logger.log('✓ Visitor Visa test row written.'); }
function testVisitorVisaIcWrite()    { appendRow(sampleDataVisitorVisaIc(),  getConfig('visitor_visa_inside_canada'), '', ''); Logger.log('✓ Visitor Visa Inside Canada test row written.'); }

function testPipeline()               { runPipelineFor('pgwp'); }
function testVisitorPipeline()        { runPipelineFor('visitor_record'); }
function testStudyExtensionPipeline() { runPipelineFor('study_extension'); }
function testVisitorVisaPipeline()    { runPipelineFor('visitor_visa'); }
function testVisitorVisaIcPipeline()  { runPipelineFor('visitor_visa_inside_canada'); }

function runPipelineFor(formType) {
  const cfg    = getConfig(formType);
  const data   = cfg.sampleData();
  const result = generateClientArtifacts(data, cfg);
  appendRow(data, cfg, result.folderUrl, result.pdfUrl);
  Logger.log('✓ [' + cfg.label + '] Folder: ' + result.folderUrl);
  Logger.log('✓ [' + cfg.label + '] PDF:    ' + result.pdfUrl);
  Logger.log('✓ [' + cfg.label + '] Draft created in ' + FIRM_EMAIL + ' Drafts folder.');
}

function setupSheetForConfig(cfg) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) sheet = ss.insertSheet(cfg.sheetName);
  sheet.clear();
  sheet.appendRow(cfg.headers);
  sheet.getRange(1, 1, 1, cfg.headers.length)
    .setBackground(cfg.headerColor).setFontColor('#ffffff')
    .setFontWeight('bold').setFontSize(10);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, cfg.headers.length);
  Logger.log('✓ Sheet "' + cfg.sheetName + '" ready with ' + cfg.headers.length + ' columns.');
}

// ════════════════════════════════════════════════════════════════════
// SAMPLE DATA  (used by test functions)
// ════════════════════════════════════════════════════════════════════
function sampleDataPgwp() {
  return {
    form_type: 'pgwp',
    submission_timestamp: new Date().toISOString(),
    full_name: 'TEST CLIENT — delete me',
    uci: '1234-5678', dob: '1998-05-12',
    country_birth: 'India', citizenship: 'India',
    current_status: 'Study Permit', status_from: '2022-09-01', status_to: '2026-08-31',
    passport_number: 'X1234567', passport_country: 'India',
    passport_issued: '2020-01-01', passport_expiry: '2030-01-01',
    marital_status: 'Single',
    phone: '647-555-0123', email: 'test.client@example.com',
    alt_phone: '', address_street: '123 Main Street', address_city: 'Brampton',
    address_province: 'Ontario', postal_code: 'L6T 1A1',
    first_entry_date: '2022-08-25', first_entry_place: 'Toronto Pearson (YYZ)',
    recent_entry_date: '2024-01-10', recent_entry_place: 'Toronto Pearson (YYZ)',
    edu_name_1: 'Humber College', edu_program_1: 'Business Administration',
    edu_credential_1: 'Diploma', edu_dli_1: 'O19359011355',
    edu_start_1: '2022-09-01', edu_end_1: '2024-12-15', edu_grad_1: '2024-12-20',
    edu_address_1: '205 Humber College Blvd, Toronto',
    emp_name_1: 'Tim Hortons', emp_title_1: 'Team Member', emp_type_1: 'Part-time',
    emp_start_1: '2023-01-15', emp_end_1: '2024-12-15',
    emp_addr_1: '500 Queen St, Brampton', emp_reason_1: 'Graduated',
    aq_criminal: 'No', aq_refusal: 'No', aq_prev_canada: 'No', aq_removal: 'No',
    aq_overstay: 'No', aq_tb: 'No', aq_pending: 'No', aq_misrep: 'No',
    aq_security: 'No', aq_sp_valid: 'Yes',
    work_permit_type: 'PGWP',
    bg_program: 'Post-Graduate Work Permit (PGWP) stream',
    bg_what_did: 'Completed a 2-year Diploma in Business Administration at Humber College (Sept 2022 – Dec 2024). Worked part-time at Tim Hortons throughout studies.',
    bg_why_apply: 'I want to gain Canadian work experience in my field of study. A PGWP will allow me to work full-time and help build toward permanent residency.',
    bg_which_wp: 'I want to apply for an open work permit under the Post-Graduate Work Permit program. I graduated from Humber College in December 2024 and my study permit has expired.',
    sig_name: 'Test Client', sig_date: '2026-04-16',
    consent1: true, consent2: true,
    signature_image: ''
  };
}

function sampleDataVisitor() {
  return {
    form_type: 'visitor_record',
    submission_timestamp: new Date().toISOString(),
    full_name: 'TEST VR CLIENT — delete me',
    uci: '8765-4321', dob: '1988-03-04',
    country_birth: 'India', citizenship: 'India',
    phone: '647-555-9876', email: 'test.vr@example.com', alt_phone: '',
    address_street: '789 Visitor Ave', address_city: 'Mississauga',
    address_province: 'Ontario', postal_code: 'L5B 2C9',
    marital_status: 'Married', partner_first: 'Jane', partner_last: 'Doe',
    partner_dob: '1990-07-15', date_of_marriage: '2015-06-20', date_of_separation: '',
    current_status: 'Visitor (TRV)', status_to: '2026-05-31',
    status_expired: 'No',
    inviter_present: 'Yes', inviter_name: 'Harpreet Singh',
    inviter_address: '12 Maple Lane, Brampton, ON L6X 1A1',
    inviter_email: 'harpreet@example.com', inviter_phone: '647-555-2222',
    inviter_employed: 'Yes', inviter_job_title: 'Software Engineer',
    inviter_bank_balance: '6500',
    original_entry_date: '2024-09-10', original_entry_place: 'Toronto Pearson (YYZ)',
    recent_entry_date: '', recent_entry_place: '',
    highest_education: 'Bachelor', program_name: 'BCom',
    school_name: 'Delhi University', school_address: 'New Delhi, India',
    edu_start_date: '2008-07-01', edu_end_date: '2011-06-30',
    additional_education: '',
    has_work_experience: 'Yes', occupation: 'Accountant',
    company: 'ABC Pvt Ltd', work_start: '2012-01-01', work_end: '2024-08-31',
    work_address: 'New Delhi, India', additional_jobs: '',
    aq_criminal: 'No',  aq_criminal_detail: '',
    aq_refusal: 'No',   aq_refusal_detail: '',
    aq_removal: 'No',   aq_removal_detail: '',
    aq_overstay: 'No',  aq_overstay_detail: '',
    aq_medical: 'No',   aq_medical_detail: '',
    aq_cdn_status: 'Yes', aq_cdn_status_detail: 'Currently on TRV expiring 2026-05-31',
    doc_financial: true, doc_passport: true, doc_status: true, doc_photo: true,
    doc_edu_emp: true, doc_inviter: true, doc_medical: false, doc_marriage: true,
    sig_name: 'Test VR Client', sig_date: '2026-04-17',
    consent1: true, consent2: true, signature_image: ''
  };
}

function sampleDataStudyExtension() {
  return {
    form_type: 'study_extension',
    submission_timestamp: new Date().toISOString(),
    full_name: 'TEST SPE CLIENT — delete me',
    uci: '5555-1212', dob: '2001-11-20',
    country_birth: 'India', citizenship: 'India',
    phone: '647-555-7777', email: 'test.spe@example.com', alt_phone: '',
    address_street: '45 Campus Drive', address_city: 'North York',
    address_province: 'Ontario', postal_code: 'M2J 2X5',
    marital_status: 'Single', partner_first: '', partner_last: '',
    partner_dob: '', date_of_marriage: '', date_of_separation: '',
    previously_married: 'No', previous_spouse: '',
    current_status: 'Study Permit', status_to: '2026-08-31',
    status_expired: 'No',
    original_entry_date: '2023-08-15', original_entry_place: 'Toronto Pearson (YYZ)',
    entry_purpose: 'study',
    recent_entry_date: '', recent_entry_place: '',
    current_institution: 'Seneca Polytechnic',
    dli_number: 'O19395492322',
    current_program: 'Diploma in Computer Programming',
    credential: 'Diploma',
    loa_received: 'Yes',
    course_start: '2026-09-01', course_end: '2028-06-30',
    tuition_fees: '17500', funds_available: '20000',
    who_pays: 'Parents', payer_details: 'Father — Jaswinder Singh, Farmer, India',
    highest_past_education: 'Grade 12 / High School',
    past_program: '', past_school: 'DAV Public School',
    past_school_address: 'Jalandhar, India',
    past_edu_start: '2017-04-01', past_edu_end: '2023-03-30',
    additional_education: '',
    aq_prior_applied: 'Yes', aq_prior_applied_detail: 'Initial study permit approved in 2023.',
    aq_refusal: 'No',  aq_refusal_detail: '',
    aq_criminal: 'No', aq_criminal_detail: '',
    aq_political: 'No', aq_political_detail: '',
    aq_military: 'No',  aq_military_detail: '',
    aq_medical: 'No',   aq_medical_detail: '',
    explain_yes: '',
    doc_passport: true, doc_current_status: true, doc_marriage: false, doc_loa: true,
    doc_parents: true,  doc_transcripts: true,    doc_funds: true,     doc_photo: true, doc_medical: false,
    sig_name: 'Test SPE Client', sig_date: '2026-04-17',
    consent1: true, consent2: true, signature_image: ''
  };
}

function sampleDataVisitorVisa() {
  return {
    form_type: 'visitor_visa',
    submission_timestamp: new Date().toISOString(),
    super_visa: 'No',
    full_name: 'TEST TRV CLIENT — delete me',
    uci: '', dob: '1980-06-15', sex: 'Male', native_language: 'Punjabi',
    country_birth: 'India', city_birth: 'Jalandhar', citizenship: 'India',
    country_residence: 'India',
    passport_number: 'P9876543', passport_country: 'India',
    passport_issued: '2021-01-01', passport_expiry: '2031-01-01',
    prior_passport: 'No', prior_passport_detail: '',
    phone: '+91-98765-43210', email: 'test.trv@example.com', alt_phone: '',
    address_street: '12 Green Park', address_city: 'Jalandhar',
    address_province: 'Punjab', address_country: 'India', postal_code: '144001',
    marital_status: 'Married',
    partner_first: 'Jasbir', partner_last: 'Kaur',
    partner_dob: '1982-09-12', date_of_marriage: '2005-05-20', date_of_separation: '',
    previously_married: 'No', previous_spouse: '',
    visit_purpose: 'Family visit',
    visit_purpose_detail: 'Visiting my brother in Brampton for 2 months. Plan to attend my nephew\u2019s wedding and tour Niagara Falls.',
    intended_arrival: '2026-06-15', intended_departure: '2026-08-15',
    cities_to_visit: 'Brampton, Toronto, Niagara Falls',
    prior_canada: 'No', prior_canada_detail: '',
    has_host: 'Yes',
    host_name: 'Harpreet Singh', host_relationship: 'Brother', host_status: 'Canadian Citizen',
    host_address: '12 Maple Lane, Brampton, ON L6X 1A1, Canada',
    host_email: 'harpreet@example.com', host_phone: '647-555-2222',
    host_occupation: 'Software Engineer', host_household_size: '4', host_income: '95000',
    who_pays: 'Self',
    payer_details: '',
    applicant_occupation: 'Business Owner', applicant_employer: 'Singh Traders (self-employed)',
    _count_education: 1,
    education_list: [
      {
        level: 'Bachelor',
        program: 'Bachelor of Commerce',
        school: 'Guru Nanak Dev University',
        school_address: 'Amritsar, India',
        start: '1998-07-01',
        end: '2001-05-30',
        present: '',
        country: 'India'
      }
    ],
    _count_employment: 1,
    employment_list: [
      {
        occupation: 'Business Owner',
        company: 'Singh Traders',
        company_address: 'Jalandhar, India',
        start: '2005-04-01',
        end: '',
        current: 'Yes',
        country: 'India',
        duties: 'Retail management, supplier coordination, and accounts oversight.'
      }
    ],
    annual_income: '₹18,00,000', savings_amount: 'CAD 15,000',
    funds_source: 'Business income and fixed deposits accumulated over 15 years.',
    father_name: 'Surjit Singh', father_dob: '1950-01-10',
    father_country: 'India', father_occupation: 'Retired Farmer',
    mother_name: 'Balbir Kaur', mother_dob: '1952-03-22',
    mother_country: 'India', mother_occupation: 'Homemaker',
    children_list: [
      { name: 'Manpreet Singh', dob: '2008-07-14', relationship: 'Son', country: 'India', accompanying: 'No' }
    ],
    siblings_list: [
      { name: 'Harpreet Singh', dob: '1978-04-02', relationship: 'Brother', country: 'Canada', accompanying: 'N/A' }
    ],
    family_in_canada: 'Yes',
    canada_family_list: [
      { name: 'Harpreet Singh', relationship: 'Brother', status: 'Canadian Citizen', city: 'Brampton, ON' }
    ],
    has_travel: 'Yes',
    travel_list: [
      { country: 'United Arab Emirates', purpose: 'Tourism', start: '2019-03-01', end: '2019-03-10' },
      { country: 'Thailand', purpose: 'Tourism', start: '2022-11-05', end: '2022-11-18' }
    ],
    has_refusal: 'No', refusal_list: [],
    aq_criminal: 'No',  aq_criminal_detail: '',
    aq_refusal: 'No',   aq_refusal_detail: '',
    aq_overstay: 'No',  aq_overstay_detail: '',
    aq_medical: 'No',   aq_medical_detail: '',
    aq_military: 'No',  aq_military_detail: '',
    aq_misrep: 'No',    aq_misrep_detail: '',
    sv_host_lico: false, sv_medical_insurance: false,
    sv_invitation_letter: false, sv_host_relation: '',
    doc_passport: true, doc_photo: true, doc_funds: true, doc_itinerary: true,
    doc_employment: true, doc_ties: true,
    doc_invitation: true, doc_host_status: true,
    doc_host_income: false, doc_medical_ins: false, doc_marriage: true,
    sig_name: 'Test TRV Client', sig_date: '2026-04-23',
    consent1: true, consent2: true, signature_image: ''
  };
}

function sampleDataVisitorVisaIc() {
  return {
    form_type: 'visitor_visa_inside_canada',
    submission_timestamp: new Date().toISOString(),
    full_name: 'TEST TRV-IC CLIENT — delete me',
    uci: '1234-5678',
    dob: '1999-04-11',
    sex: 'Female',
    country_birth: 'India',
    citizenship: 'India',
    phone: '647-555-1100',
    email: 'test.trvic@example.com',
    address_street: '215 Queen St E',
    address_city: 'Brampton',
    address_province: 'Ontario',
    postal_code: 'L6W 2B8',
    marital_status: 'Married',
    partner_first: 'Aman',
    partner_last: 'Singh',
    partner_dob: '1998-10-20',
    date_of_marriage: '2022-02-18',
    date_of_separation: '',
    previously_married: 'No',
    previous_spouse: '',
    current_status: 'Study Permit',
    status_from: '2024-09-01',
    status_to: '2027-08-31',
    status_doc_number: 'S123456789',
    status_holder_name: 'Sheridan College (DLI: O19395735882)',
    _count_address: 1,
    address_list: [
      {
        street: '215 Queen St E',
        city: 'Brampton',
        province: 'Ontario',
        from: '2024-09-01',
        to: '',
        present: 'Yes'
      }
    ],
    _count_employment: 1,
    employment_list: [
      {
        occupation: 'Customer Service Associate',
        company: 'FreshMart',
        company_address: 'Mississauga, ON',
        start: '2025-01-10',
        end: '',
        current: 'Yes',
        country: 'Canada',
        duties: 'Customer support, cashier, inventory checks.'
      }
    ],
    _count_education: 1,
    education_list: [
      {
        level: 'Diploma / Certificate',
        program: 'Business Administration',
        school: 'Sheridan College',
        school_address: 'Brampton, Ontario',
        start: '2024-09-01',
        end: '',
        present: 'Yes',
        country: 'Canada'
      }
    ],
    savings_amount: 'CAD 8,500',
    funds_source: 'Part-time employment and family support.',
    who_pays: 'Self',
    payer_details: '',
    aq_criminal: 'No',  aq_criminal_detail: '',
    aq_refusal: 'No',   aq_refusal_detail: '',
    aq_overstay: 'No',  aq_overstay_detail: '',
    aq_medical: 'No',   aq_medical_detail: '',
    aq_military: 'No',  aq_military_detail: '',
    aq_misrep: 'No',    aq_misrep_detail: '',
    doc_passport: true, doc_photo: true, doc_funds: true, doc_itinerary: true,
    doc_employment: true, doc_ties: true,
    doc_invitation: false, doc_host_status: false, doc_host_income: false, doc_medical_ins: false, doc_marriage: true,
    sig_name: 'Test TRV-IC Client', sig_date: '2026-05-05',
    consent1: true, consent2: true, signature_image: ''
  };
}

// ════════════════════════════════════════════════════════════════════
// SHEET WRITE  (config-driven)
// ════════════════════════════════════════════════════════════════════
function appendRow(d, cfg, folderUrl, pdfUrl) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(cfg.sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.sheetName);
    sheet.appendRow(cfg.headers);
    sheet.getRange(1, 1, 1, cfg.headers.length)
      .setBackground(cfg.headerColor).setFontColor('#ffffff').setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  const row = cfg.buildRow(d, folderUrl || '', pdfUrl || '');
  sheet.appendRow(row);

  // Highlight rows with any "Yes" in the form's admissibility columns
  if (cfg.aqColumns && cfg.aqColumns.length) {
    const lastRow = sheet.getLastRow();
    if (cfg.aqColumns.some(c => sheet.getRange(lastRow, c).getValue() === 'Yes')) {
      sheet.getRange(lastRow, 1, 1, cfg.headers.length).setBackground('#fff3cd');
    }
  }
}

// ── PGWP row builder ────────────────────────────────────────────────
function buildPgwpRow(d, folderUrl, pdfUrl) {
  const partner = ((d.partner_first || '') + ' ' + (d.partner_last || '')).trim();
  const edu     = serialize(d, 'edu_', ['name','program','credential','dli','start','end','grad','address']);
  const emp     = serialize(d, 'emp_', ['name','title','type','start','end','addr','reason']);
  return [
    d.submission_timestamp || new Date().toISOString(),
    d.full_name||'', d.uci||'', d.dob||'', d.country_birth||'', d.citizenship||'',
    d.current_status||'', d.status_from||'', d.status_to||'',
    d.passport_number||'', d.passport_country||'', d.passport_issued||'', d.passport_expiry||'',
    d.marital_status||'', partner, d.partner_dob||'', d.date_of_marriage||'', d.date_of_separation||'',
    d.phone||'', d.email||'', d.alt_phone||'',
    d.address_street||'', d.address_city||'', d.address_province||'', d.postal_code||'',
    d.first_entry_date||'', d.first_entry_place||'', d.recent_entry_date||'', d.recent_entry_place||'',
    edu, emp,
    d.aq_criminal||'',     d.aq_criminal_detail||'',
    d.aq_refusal||'',      d.aq_refusal_detail||'',
    d.aq_prev_canada||'',  d.aq_prev_canada_detail||'',
    d.aq_removal||'',      d.aq_removal_detail||'',
    d.aq_overstay||'',     d.aq_overstay_detail||'',
    d.aq_tb||'',           d.aq_tb_detail||'',
    d.aq_pending||'',      d.aq_pending_detail||'',
    d.aq_misrep||'',       d.aq_misrep_detail||'',
    d.aq_security||'',     d.aq_security_detail||'',
    d.aq_sp_valid||'',     d.aq_sp_valid_detail||'',
    d.work_permit_type||'',
    d.bg_program||'',
    d.bg_what_did||'',
    d.bg_why_apply||'',
    d.bg_which_wp||'',
    d.sig_name||'', d.sig_date||'',
    d.consent1 ? 'Yes' : 'No',
    d.consent2 ? 'Yes' : 'No',
    d.signature_image ? 'Yes' : 'No',
    folderUrl, pdfUrl
  ];
}

// ── Visitor Record row builder ──────────────────────────────────────
function buildVisitorRow(d, folderUrl, pdfUrl) {
  const partner = ((d.partner_first || '') + ' ' + (d.partner_last || '')).trim();
  return [
    d.submission_timestamp || new Date().toISOString(),
    'Visitor Record',
    d.full_name||'', d.uci||'', d.dob||'', d.country_birth||'', d.citizenship||'',
    d.phone||'', d.email||'', d.alt_phone||'',
    d.address_street||'', d.address_city||'', d.address_province||'', d.postal_code||'',
    d.marital_status||'', partner, d.partner_dob||'', d.date_of_marriage||'', d.date_of_separation||'',
    d.current_status||'', d.status_to||'', d.status_expired||'',
    d.inviter_present||'',  d.inviter_name||'', d.inviter_address||'',
    d.inviter_email||'',    d.inviter_phone||'',
    d.inviter_employed||'', d.inviter_job_title||'', d.inviter_bank_balance||'',
    d.original_entry_date||'', d.original_entry_place||'',
    d.recent_entry_date||'',   d.recent_entry_place||'',
    d.highest_education||'', d.program_name||'',
    d.school_name||'', d.school_address||'',
    d.edu_start_date||'', d.edu_end_date||'', d.additional_education||'',
    d.has_work_experience||'', d.occupation||'', d.company||'',
    d.work_start||'', d.work_end||'', d.work_address||'', d.additional_jobs||'',
    d.aq_criminal||'',   d.aq_criminal_detail||'',
    d.aq_refusal||'',    d.aq_refusal_detail||'',
    d.aq_removal||'',    d.aq_removal_detail||'',
    d.aq_overstay||'',   d.aq_overstay_detail||'',
    d.aq_medical||'',    d.aq_medical_detail||'',
    d.aq_cdn_status||'', d.aq_cdn_status_detail||'',
    d.doc_financial ? 'Yes' : 'No',
    d.doc_passport  ? 'Yes' : 'No',
    d.doc_status    ? 'Yes' : 'No',
    d.doc_photo     ? 'Yes' : 'No',
    d.doc_edu_emp   ? 'Yes' : 'No',
    d.doc_inviter   ? 'Yes' : 'No',
    d.doc_medical   ? 'Yes' : 'No',
    d.doc_marriage  ? 'Yes' : 'No',
    d.sig_name||'', d.sig_date||'',
    d.consent1 ? 'Yes' : 'No',
    d.consent2 ? 'Yes' : 'No',
    d.signature_image ? 'Yes' : 'No',
    folderUrl, pdfUrl
  ];
}

// ── Study Permit Extension row builder ─────────────────────────────
function buildStudyExtensionRow(d, folderUrl, pdfUrl) {
  const partner = ((d.partner_first || '') + ' ' + (d.partner_last || '')).trim();
  return [
    d.submission_timestamp || new Date().toISOString(),
    'Study Permit Extension',
    d.full_name||'', d.uci||'', d.dob||'', d.country_birth||'', d.citizenship||'',
    d.phone||'', d.email||'', d.alt_phone||'',
    d.address_street||'', d.address_city||'', d.address_province||'', d.postal_code||'',
    d.marital_status||'', partner, d.partner_dob||'', d.date_of_marriage||'', d.date_of_separation||'',
    d.previously_married||'', d.previous_spouse||'',
    d.current_status||'', d.status_to||'', d.status_expired||'',
    d.original_entry_date||'', d.original_entry_place||'', d.entry_purpose||'',
    d.recent_entry_date||'', d.recent_entry_place||'',
    d.current_institution||'', d.dli_number||'', d.current_program||'', d.credential||'',
    d.loa_received||'', d.course_start||'', d.course_end||'',
    d.tuition_fees||'', d.funds_available||'', d.who_pays||'', d.payer_details||'',
    d.highest_past_education||'', d.past_program||'', d.past_school||'', d.past_school_address||'',
    d.past_edu_start||'', d.past_edu_end||'', d.additional_education||'',
    d.aq_prior_applied||'', d.aq_prior_applied_detail||'',
    d.aq_refusal||'',       d.aq_refusal_detail||'',
    d.aq_criminal||'',      d.aq_criminal_detail||'',
    d.aq_political||'',     d.aq_political_detail||'',
    d.aq_military||'',      d.aq_military_detail||'',
    d.aq_medical||'',       d.aq_medical_detail||'',
    d.explain_yes||'',
    d.doc_passport       ? 'Yes' : 'No',
    d.doc_current_status ? 'Yes' : 'No',
    d.doc_marriage       ? 'Yes' : 'No',
    d.doc_loa            ? 'Yes' : 'No',
    d.doc_parents        ? 'Yes' : 'No',
    d.doc_transcripts    ? 'Yes' : 'No',
    d.doc_funds          ? 'Yes' : 'No',
    d.doc_photo          ? 'Yes' : 'No',
    d.doc_medical        ? 'Yes' : 'No',
    d.sig_name||'', d.sig_date||'',
    d.consent1 ? 'Yes' : 'No',
    d.consent2 ? 'Yes' : 'No',
    d.signature_image ? 'Yes' : 'No',
    folderUrl, pdfUrl
  ];
}

// ── Visitor Visa (TRV / Super Visa) row builder ─────────────────────
function buildVisitorVisaRow(d, folderUrl, pdfUrl) {
  const partner   = ((d.partner_first || '') + ' ' + (d.partner_last || '')).trim();
  const appType   = (d.super_visa === 'Yes') ? 'Super Visa' : 'Visitor Visa (TRV)';
  const education = flattenEducation(d);
  const employment = flattenEmployment(d);
  const children  = serializeList(d.children_list,     ['name','dob','relationship','country','accompanying']);
  const siblings  = serializeList(d.siblings_list,     ['name','dob','relationship','country','accompanying']);
  const cdnFamily = serializeList(d.canada_family_list,['name','relationship','status','city']);
  const travel    = serializeList(d.travel_list,       ['country','purpose','start','end']);
  const refusals  = serializeList(d.refusal_list,      ['country','visa_type','date','reason']);

  return [
    d.submission_timestamp || new Date().toISOString(),
    'Visitor Visa', appType,
    d.full_name||'', d.uci||'', d.dob||'', d.sex||'', d.native_language||'',
    d.country_birth||'', d.city_birth||'', d.citizenship||'', d.country_residence||'',
    d.passport_number||'', d.passport_country||'', d.passport_issued||'', d.passport_expiry||'',
    d.prior_passport||'', d.prior_passport_detail||'',
    d.phone||'', d.email||'', d.alt_phone||'',
    d.address_street||'', d.address_city||'', d.address_province||'', d.address_country||'', d.postal_code||'',
    d.marital_status||'', partner, d.partner_dob||'', d.date_of_marriage||'', d.date_of_separation||'',
    d.previously_married||'', d.previous_spouse||'',
    d.visit_purpose||'', d.visit_purpose_detail||'',
    d.intended_arrival||'', d.intended_departure||'',
    d.cities_to_visit||'', d.prior_canada||'', d.prior_canada_detail||'',
    d.has_host||'', d.host_name||'', d.host_relationship||'', d.host_status||'',
    d.host_address||'', d.host_email||'', d.host_phone||'',
    d.host_occupation||'', d.host_household_size||'', d.host_income||'',
    d.who_pays||'', d.payer_details||'',
    d.applicant_occupation||'', d.applicant_employer||'',
    education, employment,
    d.annual_income||'', d.savings_amount||'', d.funds_source||'',
    d.father_name||'', d.father_dob||'', d.father_country||'', d.father_occupation||'',
    d.mother_name||'', d.mother_dob||'', d.mother_country||'', d.mother_occupation||'',
    children, siblings, d.family_in_canada||'', cdnFamily,
    d.has_travel||'', travel, d.has_refusal||'', refusals,
    d.aq_criminal||'', d.aq_criminal_detail||'',
    d.aq_refusal||'',  d.aq_refusal_detail||'',
    d.aq_overstay||'', d.aq_overstay_detail||'',
    d.aq_medical||'',  d.aq_medical_detail||'',
    d.aq_military||'', d.aq_military_detail||'',
    d.aq_misrep||'',   d.aq_misrep_detail||'',
    d.sv_host_lico         ? 'Yes' : 'No',
    d.sv_medical_insurance ? 'Yes' : 'No',
    d.sv_invitation_letter ? 'Yes' : 'No',
    d.sv_host_relation || '',
    d.doc_passport     ? 'Yes' : 'No',
    d.doc_photo        ? 'Yes' : 'No',
    d.doc_funds        ? 'Yes' : 'No',
    d.doc_itinerary    ? 'Yes' : 'No',
    d.doc_employment   ? 'Yes' : 'No',
    d.doc_ties         ? 'Yes' : 'No',
    d.doc_invitation   ? 'Yes' : 'No',
    d.doc_host_status  ? 'Yes' : 'No',
    d.doc_host_income  ? 'Yes' : 'No',
    d.doc_medical_ins  ? 'Yes' : 'No',
    d.doc_marriage     ? 'Yes' : 'No',
    d.sig_name||'', d.sig_date||'',
    d.consent1 ? 'Yes' : 'No',
    d.consent2 ? 'Yes' : 'No',
    d.signature_image ? 'Yes' : 'No',
    folderUrl, pdfUrl
  ];
}

// ── Visitor Visa (Inside Canada) row builder ─────────────────────────
function buildVisitorVisaIcRow(d, folderUrl, pdfUrl) {
  const partner = ((d.partner_first || '') + ' ' + (d.partner_last || '')).trim();
  const addressHistory = flattenAddressHistory(d);
  const employment = flattenEmployment(d);
  const education = flattenEducation(d);
  return [
    d.submission_timestamp || new Date().toISOString(),
    'Visitor Visa (Inside Canada)',
    d.full_name||'', d.uci||'', d.dob||'', d.sex||'', d.country_birth||'', d.citizenship||'',
    d.phone||'', d.email||'', d.address_street||'', d.address_city||'', d.address_province||'', d.postal_code||'',
    d.marital_status||'', partner, d.partner_dob||'', d.date_of_marriage||'', d.date_of_separation||'',
    d.previously_married||'', d.previous_spouse||'',
    d.current_status||'', d.status_from||'', d.status_to||'', d.status_doc_number||'', d.status_holder_name||'',
    addressHistory, employment, education,
    d.savings_amount||'', d.funds_source||'', d.who_pays||'', d.payer_details||'',
    d.aq_criminal||'', d.aq_criminal_detail||'',
    d.aq_refusal||'',  d.aq_refusal_detail||'',
    d.aq_overstay||'', d.aq_overstay_detail||'',
    d.aq_medical||'',  d.aq_medical_detail||'',
    d.aq_military||'', d.aq_military_detail||'',
    d.aq_misrep||'',   d.aq_misrep_detail||'',
    d.doc_passport     ? 'Yes' : 'No',
    d.doc_photo        ? 'Yes' : 'No',
    d.doc_funds        ? 'Yes' : 'No',
    d.doc_itinerary    ? 'Yes' : 'No',
    d.doc_employment   ? 'Yes' : 'No',
    d.doc_ties         ? 'Yes' : 'No',
    d.doc_invitation   ? 'Yes' : 'No',
    d.doc_host_status  ? 'Yes' : 'No',
    d.doc_host_income  ? 'Yes' : 'No',
    d.doc_medical_ins  ? 'Yes' : 'No',
    d.doc_marriage     ? 'Yes' : 'No',
    d.sig_name||'', d.sig_date||'',
    d.consent1 ? 'Yes' : 'No',
    d.consent2 ? 'Yes' : 'No',
    d.signature_image ? 'Yes' : 'No',
    folderUrl, pdfUrl
  ];
}

// ════════════════════════════════════════════════════════════════════
// FOLDER + PDF + GMAIL DRAFT  (config-driven)
// ════════════════════════════════════════════════════════════════════
function generateClientArtifacts(d, cfg) {
  const clientName = sanitize(d.full_name || 'Unknown Client');
  const stamp      = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const parent     = DriveApp.getFolderById(DRAFT_PARENT_FOLDER_ID);
  const folderName = cfg.folderPrefix + ' - ' + clientName;
  const folder     = parent.createFolder(folderName);

  const html    = cfg.buildPdfHtml(d);
  const pdfBlob = Utilities.newBlob(html, 'text/html', 'tmp.html')
                    .getAs('application/pdf')
                    .setName(cfg.pdfFilePrefix + ' - ' + clientName + ' - ' + stamp + '.pdf');
  const pdfFile = folder.createFile(pdfBlob);

  if (d.signature_image && d.signature_image.indexOf('data:image') === 0) {
    try {
      folder.createFile(dataUrlToBlob(d.signature_image, 'signature - ' + clientName + '.png'));
    } catch (e) { Logger.log('Signature save failed: ' + e); }
  }

  if (d.email) {
    GmailApp.createDraft(d.email, cfg.emailSubject(clientName), cfg.buildEmailPlain(d, folder.getUrl()), {
      htmlBody:    cfg.buildEmailHtml(d, folder.getUrl()),
      attachments: [pdfBlob],
      name:        FIRM_NAME,
      replyTo:     FIRM_EMAIL
    });
  }

  return { folderUrl: folder.getUrl(), pdfUrl: pdfFile.getUrl() };
}

// ════════════════════════════════════════════════════════════════════
// PDF HTML — PGWP
// ════════════════════════════════════════════════════════════════════
function buildPgwpPdfHtml(d) {
  const partner = ((d.partner_first || '') + ' ' + (d.partner_last || '')).trim();
  const edu     = listEntries(d, 'edu_', [
                    ['Institution','name'],['Program','program'],['Credential','credential'],
                    ['DLI #','dli'],['Start','start'],['End','end'],['Graduated','grad'],['Address','address']]);
  const emp     = listEntries(d, 'emp_', [
                    ['Employer','name'],['Title','title'],['Type','type'],
                    ['Start','start'],['End','end'],['Address','addr'],['Reason left','reason']]);

  const aqRows = [
    ['Criminal charges',          d.aq_criminal,    d.aq_criminal_detail],
    ['Prior visa refusal',        d.aq_refusal,     d.aq_refusal_detail],
    ['Prior Canadian application',d.aq_prev_canada, d.aq_prev_canada_detail],
    ['Removal order',             d.aq_removal,     d.aq_removal_detail],
    ['Overstay',                  d.aq_overstay,    d.aq_overstay_detail],
    ['TB / serious medical',      d.aq_tb,          d.aq_tb_detail],
    ['Pending IRCC application',  d.aq_pending,     d.aq_pending_detail],
    ['Misrepresentation',         d.aq_misrep,      d.aq_misrep_detail],
    ['Security / war crime',      d.aq_security,    d.aq_security_detail],
    ['Valid study permit',        d.aq_sp_valid,    d.aq_sp_valid_detail]
  ];

  return [
    pdfHeadOpen('PGWP Intake Form', d),
    sect('Personal', [
      ['Full legal name', d.full_name],
      ['UCI', d.uci],
      ['Date of birth', d.dob],
      ['Country of birth', d.country_birth],
      ['Country of citizenship', d.citizenship]
    ]),
    sect('Current immigration status', [
      ['Status', d.current_status], ['Valid from', d.status_from], ['Valid to', d.status_to]
    ]),
    sect('Passport', [
      ['Passport #', d.passport_number], ['Country of issue', d.passport_country],
      ['Issue date', d.passport_issued], ['Expiry date', d.passport_expiry]
    ]),
    sect('Marital', [
      ['Status', d.marital_status], ['Partner', partner], ['Partner DOB', d.partner_dob],
      ['Marriage date', d.date_of_marriage], ['Separation date', d.date_of_separation]
    ]),
    sect('Contact & address', [
      ['Phone', d.phone], ['Email', d.email], ['Alt phone', d.alt_phone],
      ['Street', d.address_street], ['City', d.address_city],
      ['Province', d.address_province], ['Postal', d.postal_code]
    ]),
    sect('Entry to Canada', [
      ['First entry date', d.first_entry_date], ['First port', d.first_entry_place],
      ['Most recent entry', d.recent_entry_date], ['Most recent port', d.recent_entry_place]
    ]),
    edu ? '<h2>Education</h2>' + edu : '',
    emp ? '<h2>Employment</h2>' + emp : '',
    aqTable(aqRows),
    sect('Application background', [
      ['Work permit type',              d.work_permit_type],
      ['Program',                       d.bg_program],
      ['Background (what they did)',    d.bg_what_did],
      ['Reason for applying',           d.bg_why_apply],
      ['Work permit (own words)',       d.bg_which_wp]
    ]),
    declarationBlock(d),
    pdfFootClose()
  ].join('');
}

// ════════════════════════════════════════════════════════════════════
// PDF HTML — Visitor Record
// ════════════════════════════════════════════════════════════════════
function buildVisitorPdfHtml(d) {
  const partner = ((d.partner_first || '') + ' ' + (d.partner_last || '')).trim();

  const aqRows = [
    ['Criminal charges',                 d.aq_criminal,   d.aq_criminal_detail],
    ['Prior visa refusal',               d.aq_refusal,    d.aq_refusal_detail],
    ['Removal order',                    d.aq_removal,    d.aq_removal_detail],
    ['Overstay / non-compliance',        d.aq_overstay,   d.aq_overstay_detail],
    ['TB / serious medical',             d.aq_medical,    d.aq_medical_detail],
    ['Q33 — Valid Canadian status',      d.aq_cdn_status, d.aq_cdn_status_detail]
  ];

  const inviterSection = (d.inviter_present === 'Yes') ? sect('Inviter / financial support', [
    ['Inviter present', d.inviter_present],
    ['Name', d.inviter_name],
    ['Address', d.inviter_address],
    ['Email', d.inviter_email],
    ['Phone', d.inviter_phone],
    ['Employed', d.inviter_employed],
    ['Job title', d.inviter_job_title],
    ['Bank balance (CAD)', d.inviter_bank_balance ? ('$' + d.inviter_bank_balance) : '']
  ]) : sect('Inviter / financial support', [['Inviter present', 'No']]);

  const docs = [
    ['Proof of financial support',         d.doc_financial],
    ['Passport',                           d.doc_passport],
    ['Current/previous status document',   d.doc_status],
    ['Digital photo',                      d.doc_photo],
    ['Education / employment documents',   d.doc_edu_emp],
    ['Inviter\u2019s supporting documents',  d.doc_inviter],
    ['Medical examination (optional)',     d.doc_medical],
    ['Marriage certificate',               d.doc_marriage]
  ];
  const docTable = '<h2>Document checklist</h2><table>' +
    docs.map(r => '<tr><th>' + esc(r[0]) + '</th><td><strong>' + (r[1] ? '✓ Confirmed' : '— Not confirmed') + '</strong></td></tr>').join('') +
    '</table>';

  return [
    pdfHeadOpen('Visitor Record Extension Intake', d),
    sect('Personal', [
      ['Full legal name', d.full_name], ['UCI', d.uci], ['Date of birth', d.dob],
      ['Country of birth', d.country_birth], ['Country of citizenship', d.citizenship]
    ]),
    sect('Contact & address', [
      ['Phone', d.phone], ['Email', d.email], ['Alt phone', d.alt_phone],
      ['Street', d.address_street], ['City', d.address_city],
      ['Province', d.address_province], ['Postal', d.postal_code]
    ]),
    sect('Marital', [
      ['Status', d.marital_status], ['Partner', partner], ['Partner DOB', d.partner_dob],
      ['Marriage date', d.date_of_marriage], ['Separation date', d.date_of_separation]
    ]),
    sect('Status & restoration', [
      ['Current status', d.current_status],
      ['Status valid to', d.status_to],
      ['Status already expired (restoration?)', d.status_expired]
    ]),
    inviterSection,
    sect('Entry to Canada', [
      ['Original entry date', d.original_entry_date], ['Original entry place', d.original_entry_place],
      ['Most recent entry date', d.recent_entry_date], ['Most recent entry place', d.recent_entry_place]
    ]),
    sect('Education', [
      ['Highest level', d.highest_education], ['Program', d.program_name],
      ['School', d.school_name], ['School address', d.school_address],
      ['Start', d.edu_start_date], ['End', d.edu_end_date],
      ['Additional education', d.additional_education]
    ]),
    sect('Work experience', [
      ['Has work experience', d.has_work_experience],
      ['Occupation', d.occupation], ['Company', d.company],
      ['Start', d.work_start], ['End', d.work_end],
      ['Address', d.work_address], ['Additional jobs', d.additional_jobs]
    ]),
    aqTable(aqRows),
    docTable,
    declarationBlock(d),
    pdfFootClose()
  ].join('');
}

// ════════════════════════════════════════════════════════════════════
// PDF HTML — Study Permit Extension
// ════════════════════════════════════════════════════════════════════
function buildStudyExtensionPdfHtml(d) {
  const partner = ((d.partner_first || '') + ' ' + (d.partner_last || '')).trim();

  const aqRows = [
    ['Prior Canadian application',             d.aq_prior_applied, d.aq_prior_applied_detail],
    ['Prior visa refusal (any country)',       d.aq_refusal,       d.aq_refusal_detail],
    ['Criminal charges / arrests / convictions', d.aq_criminal,    d.aq_criminal_detail],
    ['Political / trade-union association',    d.aq_political,     d.aq_political_detail],
    ['Served in Army / Military',              d.aq_military,      d.aq_military_detail],
    ['TB / serious medical condition',         d.aq_medical,       d.aq_medical_detail]
  ];

  const feesLine = d.tuition_fees ? ('$' + d.tuition_fees + ' CAD') : '';
  const fundsLine = d.funds_available ? ('$' + d.funds_available + ' CAD') : '';

  const docs = [
    ['Passport (biodata + visa pages)',                    d.doc_passport],
    ['Current Canadian immigration documents',             d.doc_current_status],
    ['Marriage licence / certificate / spouse passport',   d.doc_marriage],
    ['Letter of Acceptance / Enrolment letter',            d.doc_loa],
    ['Parents\u2019 status / work permit / job letter / LMIA (minor SP only)', d.doc_parents],
    ['Recent transcripts (proof of academic standing)',    d.doc_transcripts],
    ['Proof of funds (fee receipts, bank statements, portfolio)', d.doc_funds],
    ['Passport-size digital photo',                        d.doc_photo],
    ['Medical examination (if done within 12 months)',     d.doc_medical]
  ];
  const docTable = '<h2>Document checklist</h2><table>' +
    docs.map(r => '<tr><th>' + esc(r[0]) + '</th><td><strong>' + (r[1] ? '✓ Confirmed' : '— Not confirmed') + '</strong></td></tr>').join('') +
    '</table>';

  return [
    pdfHeadOpen('Study Permit Extension Intake', d),
    sect('Personal', [
      ['Full legal name', d.full_name], ['UCI', d.uci], ['Date of birth', d.dob],
      ['Country of birth', d.country_birth], ['Country of citizenship', d.citizenship]
    ]),
    sect('Contact & address', [
      ['Phone', d.phone], ['Email', d.email], ['Alt phone', d.alt_phone],
      ['Street', d.address_street], ['City', d.address_city],
      ['Province', d.address_province], ['Postal', d.postal_code]
    ]),
    sect('Marital', [
      ['Status', d.marital_status], ['Partner', partner], ['Partner DOB', d.partner_dob],
      ['Marriage / relationship date', d.date_of_marriage],
      ['Separation / divorce date', d.date_of_separation],
      ['Previously married', d.previously_married],
      ['Previous spouse', d.previous_spouse]
    ]),
    sect('Status & restoration', [
      ['Current status', d.current_status],
      ['Status valid to', d.status_to],
      ['Status already expired (restoration?)', d.status_expired]
    ]),
    sect('Entry to Canada', [
      ['Original entry date', d.original_entry_date],
      ['Original entry place', d.original_entry_place],
      ['Purpose of original entry', d.entry_purpose],
      ['Most recent entry date', d.recent_entry_date],
      ['Most recent entry place', d.recent_entry_place]
    ]),
    sect('Current program (being extended)', [
      ['Institution', d.current_institution],
      ['DLI number', d.dli_number],
      ['Program', d.current_program],
      ['Credential', d.credential],
      ['LOA / Enrolment letter received', d.loa_received],
      ['Course start', d.course_start],
      ['Course end', d.course_end],
      ['Tuition fees', feesLine],
      ['Funds available', fundsLine],
      ['Who pays the fees', d.who_pays],
      ['Payer details', d.payer_details]
    ]),
    sect('Past education', [
      ['Highest past level', d.highest_past_education],
      ['Program', d.past_program],
      ['School', d.past_school],
      ['School address', d.past_school_address],
      ['Start', d.past_edu_start],
      ['End', d.past_edu_end],
      ['Additional education', d.additional_education]
    ]),
    aqTable(aqRows),
    d.explain_yes ? '<h2>Further explanation</h2><table><tr><td>' + esc(d.explain_yes) + '</td></tr></table>' : '',
    docTable,
    declarationBlock(d),
    pdfFootClose()
  ].join('');
}

// ════════════════════════════════════════════════════════════════════
// PDF HTML — Visitor Visa (TRV / Super Visa)
// ════════════════════════════════════════════════════════════════════
function buildVisitorVisaPdfHtml(d) {
  const partner  = ((d.partner_first || '') + ' ' + (d.partner_last || '')).trim();
  const isSuper  = (d.super_visa === 'Yes');
  const title    = isSuper ? 'Super Visa Intake' : 'Visitor Visa (TRV) Intake';

  const aqRows = [
    ['Criminal charges / convictions (any country)', d.aq_criminal, d.aq_criminal_detail],
    ['Prior visa refusal or denied entry',            d.aq_refusal,  d.aq_refusal_detail],
    ['Overstay or non-compliance in any country',     d.aq_overstay, d.aq_overstay_detail],
    ['TB or serious medical condition (last 2 yrs)',  d.aq_medical,  d.aq_medical_detail],
    ['Military / police / intelligence service',      d.aq_military, d.aq_military_detail],
    ['Misrepresentation alleged (immigration)',       d.aq_misrep,   d.aq_misrep_detail]
  ];

  const hostSection = (d.has_host === 'Yes') ? sect('Canadian host / inviter', [
    ['Host present',         d.has_host],
    ['Name',                 d.host_name],
    ['Relationship',         d.host_relationship],
    ['Status in Canada',     d.host_status],
    ['Address',              d.host_address],
    ['Email',                d.host_email],
    ['Phone',                d.host_phone],
    ['Occupation',           d.host_occupation],
    ['Household size',       d.host_household_size],
    ['Annual income (CAD)',  d.host_income ? ('$' + d.host_income) : '']
  ]) : sect('Canadian host / inviter', [['Host present', 'No — no inviter in Canada']]);

  const parentsSection = sect('Parents', [
    ['Father — name',        d.father_name],
    ['Father — DOB',         d.father_dob],
    ['Father — country',     d.father_country],
    ['Father — occupation',  d.father_occupation],
    ['Mother — name',        d.mother_name],
    ['Mother — DOB',         d.mother_dob],
    ['Mother — country',     d.mother_country],
    ['Mother — occupation',  d.mother_occupation]
  ]);

  const children = listFromArray(d.children_list, [
    ['Name','name'],['DOB','dob'],['Relationship','relationship'],
    ['Country of residence','country'],['Accompanying','accompanying']
  ]);
  const siblings = listFromArray(d.siblings_list, [
    ['Name','name'],['DOB','dob'],['Relationship','relationship'],
    ['Country of residence','country'],['Accompanying','accompanying']
  ]);
  const canadaFam = listFromArray(d.canada_family_list, [
    ['Name','name'],['Relationship','relationship'],
    ['Status in Canada','status'],['City','city']
  ]);
  const travel = listFromArray(d.travel_list, [
    ['Country','country'],['Purpose','purpose'],['Start','start'],['End','end']
  ]);
  const refusals = listFromArray(d.refusal_list, [
    ['Country','country'],['Visa type','visa_type'],['Date','date'],['Reason','reason']
  ]);
  const education = listFromArray(getVisitorVisaEducationEntries(d), [
    ['Level','level'],['Program','program'],['School','school'],
    ['School address','school_address'],['Start','start'],
    ['End','end'],['Present','present'],['Country','country']
  ]);
  const employment = listFromArray(getVisitorVisaEmploymentEntries(d), [
    ['Occupation','occupation'],['Company','company'],['Company address','company_address'],
    ['Start','start'],['End','end'],['Current','current'],
    ['Country','country'],['Duties','duties']
  ]);

  const superSection = isSuper ? sect('Super Visa acknowledgements', [
    ['Host meets LICO',                  d.sv_host_lico         ? '✓ Acknowledged' : '— Not acknowledged'],
    ['$100k+ medical insurance obtained',d.sv_medical_insurance ? '✓ Acknowledged' : '— Not acknowledged'],
    ['Notarized invitation letter',      d.sv_invitation_letter ? '✓ Acknowledged' : '— Not acknowledged'],
    ['Relation to host',                 d.sv_host_relation]
  ]) : '';

  const docs = [
    ['Passport (biodata + visa/stamp pages)',               d.doc_passport],
    ['Digital passport-style photo',                        d.doc_photo],
    ['Proof of funds (bank statements, FDs, salary slips)', d.doc_funds],
    ['Trip itinerary / travel plan',                        d.doc_itinerary],
    ['Employment letter / business registration',           d.doc_employment],
    ['Ties-to-home proof',                                  d.doc_ties],
    ['Invitation letter from Canadian host',                d.doc_invitation],
    ['Host\u2019s status document',                         d.doc_host_status],
    ['Host\u2019s income proof (NOA / T4 / pay stubs)',     d.doc_host_income],
    ['Medical insurance ($100k+, Canadian provider)',       d.doc_medical_ins],
    ['Marriage certificate',                                d.doc_marriage]
  ];
  const docTable = '<h2>Document checklist</h2><table>' +
    docs.map(r => '<tr><th>' + esc(r[0]) + '</th><td><strong>' + (r[1] ? '✓ Confirmed' : '— Not confirmed') + '</strong></td></tr>').join('') +
    '</table>';

  return [
    pdfHeadOpen(title, d),
    sect('Application type', [
      ['Category', isSuper ? 'Super Visa (parent/grandparent of Canadian citizen or PR)' : 'Standard Visitor Visa (TRV)']
    ]),
    sect('Personal', [
      ['Full legal name',         d.full_name],
      ['UCI',                     d.uci],
      ['Date of birth',           d.dob],
      ['Sex',                     d.sex],
      ['Native language',         d.native_language],
      ['Country of birth',        d.country_birth],
      ['City of birth',           d.city_birth],
      ['Country of citizenship',  d.citizenship],
      ['Country of residence',    d.country_residence]
    ]),
    sect('Passport', [
      ['Passport #',         d.passport_number],
      ['Country of issue',   d.passport_country],
      ['Issue date',         d.passport_issued],
      ['Expiry date',        d.passport_expiry],
      ['Prior passport',     d.prior_passport],
      ['Prior passport details', d.prior_passport_detail]
    ]),
    sect('Contact & home-country address', [
      ['Phone',        d.phone],
      ['Email',        d.email],
      ['Alt phone',    d.alt_phone],
      ['Street',       d.address_street],
      ['City',         d.address_city],
      ['Province / State', d.address_province],
      ['Country',      d.address_country],
      ['Postal',       d.postal_code]
    ]),
    sect('Marital', [
      ['Status',                d.marital_status],
      ['Partner',               partner],
      ['Partner DOB',           d.partner_dob],
      ['Marriage date',         d.date_of_marriage],
      ['Separation date',       d.date_of_separation],
      ['Previously married',    d.previously_married],
      ['Previous spouse',       d.previous_spouse]
    ]),
    sect('Purpose of visit & trip details', [
      ['Purpose',                 d.visit_purpose],
      ['Purpose details',         d.visit_purpose_detail],
      ['Intended arrival',        d.intended_arrival],
      ['Intended departure',      d.intended_departure],
      ['Cities to visit',         d.cities_to_visit],
      ['Previously visited Canada', d.prior_canada],
      ['Prior Canada details',    d.prior_canada_detail]
    ]),
    hostSection,
    sect('Financial support', [
      ['Who pays for the trip',    d.who_pays],
      ['Payer details',            d.payer_details],
      ['Applicant occupation',     d.applicant_occupation],
      ['Applicant employer',       d.applicant_employer],
      ['Annual income',            d.annual_income],
      ['Savings amount',           d.savings_amount],
      ['Source of funds',          d.funds_source]
    ]),
    education ? '<h2>Education History</h2>' + education : '',
    employment ? '<h2>Employment History</h2>' + employment : '',
    parentsSection,
    children  ? '<h2>Children</h2>' + children   : '',
    siblings  ? '<h2>Siblings</h2>' + siblings   : '',
    sect('Family in Canada', [['Any family member currently in Canada', d.family_in_canada]]),
    canadaFam ? canadaFam : '',
    sect('Travel history', [['Travel outside home country in last 10 years', d.has_travel]]),
    travel    ? travel    : '',
    sect('Prior refusals', [['Any prior visa refusal or entry denial (any country)', d.has_refusal]]),
    refusals  ? refusals  : '',
    aqTable(aqRows),
    superSection,
    docTable,
    declarationBlock(d),
    pdfFootClose()
  ].join('');
}

// ════════════════════════════════════════════════════════════════════
// PDF HTML — Visitor Visa (Inside Canada)
// ════════════════════════════════════════════════════════════════════
function buildVisitorVisaIcPdfHtml(d) {
  const partner = ((d.partner_first || '') + ' ' + (d.partner_last || '')).trim();

  const aqRows = [
    ['Criminal charges / convictions (any country)', d.aq_criminal, d.aq_criminal_detail],
    ['Prior visa refusal or denied entry',            d.aq_refusal,  d.aq_refusal_detail],
    ['Overstay or non-compliance in any country',     d.aq_overstay, d.aq_overstay_detail],
    ['TB or serious medical condition (last 2 yrs)',  d.aq_medical,  d.aq_medical_detail],
    ['Military / police / intelligence service',      d.aq_military, d.aq_military_detail],
    ['Misrepresentation alleged (immigration)',       d.aq_misrep,   d.aq_misrep_detail]
  ];

  const addresses = listFromArray(getVisitorVisaAddressEntries(d), [
    ['Street','street'],['City','city'],['Province','province'],
    ['From','from'],['To','to'],['Present','present']
  ]);
  const education = listFromArray(getVisitorVisaEducationEntries(d), [
    ['Level','level'],['Program','program'],['School','school'],
    ['School address','school_address'],['Start','start'],
    ['End','end'],['Present','present'],['Country','country']
  ]);
  const employment = listFromArray(getVisitorVisaEmploymentEntries(d), [
    ['Occupation','occupation'],['Company','company'],['Company address','company_address'],
    ['Start','start'],['End','end'],['Current','current'],
    ['Country','country'],['Duties','duties']
  ]);

  const docs = [
    ['Passport (biodata + visa/stamp pages)',               d.doc_passport],
    ['Digital passport-style photo',                        d.doc_photo],
    ['Proof of funds (bank statements, FDs, salary slips)', d.doc_funds],
    ['Trip itinerary / travel plan',                        d.doc_itinerary],
    ['Employment letter / work proof',                      d.doc_employment],
    ['Ties-to-home proof',                                  d.doc_ties],
    ['Invitation letter (if applicable)',                   d.doc_invitation],
    ['Host status document (if applicable)',                d.doc_host_status],
    ['Host income proof (if applicable)',                   d.doc_host_income],
    ['Medical insurance (if applicable)',                   d.doc_medical_ins],
    ['Marriage certificate (if applicable)',                d.doc_marriage]
  ];
  const docTable = '<h2>Document checklist</h2><table>' +
    docs.map(r => '<tr><th>' + esc(r[0]) + '</th><td><strong>' + (r[1] ? '✓ Confirmed' : '— Not confirmed') + '</strong></td></tr>').join('') +
    '</table>';

  return [
    pdfHeadOpen('Visitor Visa (Inside Canada) Intake', d),
    sect('Personal', [
      ['Full legal name', d.full_name],
      ['UCI', d.uci],
      ['Date of birth', d.dob],
      ['Sex', d.sex],
      ['Country of birth', d.country_birth],
      ['Country of citizenship', d.citizenship]
    ]),
    sect('Contact & current Canadian address', [
      ['Phone', d.phone],
      ['Email', d.email],
      ['Street', d.address_street],
      ['City', d.address_city],
      ['Province', d.address_province],
      ['Postal', d.postal_code]
    ]),
    sect('Marital details', [
      ['Status', d.marital_status],
      ['Partner', partner],
      ['Partner DOB', d.partner_dob],
      ['Marriage date', d.date_of_marriage],
      ['Separation date', d.date_of_separation],
      ['Previously married', d.previously_married],
      ['Previous spouse', d.previous_spouse]
    ]),
    sect('Current status in Canada', [
      ['Current status', d.current_status],
      ['Status from', d.status_from],
      ['Status valid to', d.status_to],
      ['Document number', d.status_doc_number],
      ['DLI / Employer', d.status_holder_name]
    ]),
    addresses ? '<h2>Address History (Canada)</h2>' + addresses : '',
    employment ? '<h2>Employment History</h2>' + employment : '',
    education ? '<h2>Education History</h2>' + education : '',
    sect('Funds', [
      ['Savings amount', d.savings_amount],
      ['Funds source', d.funds_source],
      ['Who pays', d.who_pays],
      ['Payer details', d.payer_details]
    ]),
    aqTable(aqRows),
    docTable,
    declarationBlock(d),
    pdfFootClose()
  ].join('');
}

// ── Shared PDF builders ─────────────────────────────────────────────
function pdfHeadOpen(title, d) {
  return [
    '<!DOCTYPE html><html><head><meta charset="utf-8"><style>',
    'body{font-family:Helvetica,Arial,sans-serif;color:#15202b;font-size:11px;margin:32px;}',
    'h1{font-size:18px;margin:0 0 4px;color:#1a3a5c;}',
    'h2{font-size:13px;margin:18px 0 6px;color:#1a3a5c;border-bottom:2px solid #1a3a5c;padding-bottom:3px;}',
    '.meta{color:#5b6b7d;font-size:10px;margin-bottom:14px;}',
    'table{width:100%;border-collapse:collapse;margin-bottom:6px;}',
    'th,td{padding:5px 8px;border-bottom:1px solid #e8edf3;text-align:left;vertical-align:top;}',
    'th{font-weight:600;width:38%;color:#5b6b7d;font-size:10px;text-transform:uppercase;letter-spacing:0.03em;}',
    'td{font-size:11px;}',
    '.foot{margin-top:24px;padding-top:10px;border-top:1px solid #e8edf3;color:#5b6b7d;font-size:9px;text-align:center;}',
    '.entry{background:#f5f6f8;padding:8px 10px;margin-bottom:6px;border-radius:4px;}',
    '.entry strong{color:#1a3a5c;}',
    '</style></head><body>',
    '<h1>' + esc(title) + '</h1>',
    '<div class="meta">' + esc(FIRM_NAME) + ' · ' + esc(FIRM_RCIC) + ' · Submitted: ' + esc(d.submission_timestamp || '') + '</div>'
  ].join('');
}

function pdfFootClose() {
  return '<div class="foot">' + esc(FIRM_NAME) + ' · ' + esc(FIRM_RCIC) + ' · ' + esc(FIRM_EMAIL) + ' · Brampton, Ontario</div></body></html>';
}

function sect(title, rows) {
  const trs = rows
    .filter(r => r[1] !== undefined && r[1] !== null && r[1] !== '')
    .map(r => '<tr><th>' + esc(r[0]) + '</th><td>' + esc(r[1]) + '</td></tr>')
    .join('');
  if (!trs) return '';
  return '<h2>' + esc(title) + '</h2><table>' + trs + '</table>';
}

function aqTable(aqRows) {
  return '<h2>Admissibility</h2><table>' +
    aqRows.map(r => {
      const ans = r[1] || '—';
      const flag = ans === 'Yes' ? ' style="background:#fff3cd"' : '';
      const det = r[2] ? '<div style="color:#5b6b7d;font-size:11px;margin-top:3px;">' + esc(r[2]) + '</div>' : '';
      return '<tr' + flag + '><th>' + esc(r[0]) + '</th><td><strong>' + esc(ans) + '</strong>' + det + '</td></tr>';
    }).join('') + '</table>';
}

function declarationBlock(d) {
  const sigImg = (d.signature_image && d.signature_image.indexOf('data:image') === 0)
    ? '<div style="margin-top:8px;"><img src="' + d.signature_image + '" style="max-height:80px;border:1px solid #ccc;"/></div>'
    : '<div style="color:#888;font-style:italic;margin-top:8px;">[No drawn signature]</div>';
  return '<h2>Declaration</h2><table>' +
    '<tr><th>Typed signature</th><td>' + esc(d.sig_name || '') + '</td></tr>' +
    '<tr><th>Signature date</th><td>' + esc(d.sig_date || '') + '</td></tr>' +
    '<tr><th>Consent — accuracy</th><td>' + (d.consent1 ? 'Yes' : 'No') + '</td></tr>' +
    '<tr><th>Consent — authorization</th><td>' + (d.consent2 ? 'Yes' : 'No') + '</td></tr>' +
    '<tr><th>Drawn signature</th><td>' + sigImg + '</td></tr>' +
    '</table>';
}

function listEntries(d, prefix, fields) {
  let html = '';
  for (let i = 1; i <= 20; i++) {
    const present = fields.some(f => d[prefix + f[1] + '_' + i]);
    if (!present) break;
    html += '<div class="entry">';
    fields.forEach(f => {
      const v = d[prefix + f[1] + '_' + i];
      if (v) html += '<div><strong>' + esc(f[0]) + ':</strong> ' + esc(v) + '</div>';
    });
    html += '</div>';
  }
  return html;
}

// ════════════════════════════════════════════════════════════════════
// EMAIL — PGWP
// ════════════════════════════════════════════════════════════════════
function buildPgwpEmailPlain(d, folderUrl) {
  const wpLabel = d.work_permit_type || 'Work Permit';
  return [
    'Hello ' + (d.full_name || '') + ',',
    '',
    'Thank you for completing the Work Permit intake form with ' + FIRM_NAME + '.',
    '',
    'Attached is a PDF copy of the information you submitted. Please review it carefully and reply to confirm everything is accurate. If anything needs to be corrected, let us know which section and what should change.',
    '',
    'Once you confirm, we will proceed with preparing and submitting your ' + wpLabel + ' application to IRCC.',
    '',
    'Internal reference folder: ' + folderUrl,
    '',
    'Best regards,', FIRM_NAME, FIRM_RCIC, FIRM_EMAIL
  ].join('\n');
}

function buildPgwpEmailHtml(d, folderUrl) {
  const wpLabel = d.work_permit_type || 'Work Permit';
  return [
    '<div style="font-family:Helvetica,Arial,sans-serif;color:#15202b;max-width:620px;line-height:1.55;">',
      '<p>Hello <strong>' + esc(d.full_name || '') + '</strong>,</p>',
      '<p>Thank you for completing the <strong>' + esc(wpLabel) + '</strong> intake form with ' + esc(FIRM_NAME) + '.</p>',
      '<p>Attached is a <strong>PDF copy</strong> of the information you submitted. Please review it carefully and reply to confirm that everything is accurate.</p>',
      '<p>If any detail needs to be corrected, simply tell us which section and what should change — no need to fill the form again.</p>',
      '<p>Once you confirm, we will proceed with preparing and submitting your <strong>' + esc(wpLabel) + '</strong> application to IRCC.</p>',
      '<p style="margin-top:28px;">Best regards,<br><strong>' + esc(FIRM_NAME) + '</strong><br>' + esc(FIRM_RCIC) + '<br>',
        '<a href="mailto:' + esc(FIRM_EMAIL) + '">' + esc(FIRM_EMAIL) + '</a></p>',
      '<hr style="border:none;border-top:1px solid #e8edf3;margin-top:24px;">',
      '<p style="font-size:11px;color:#888;">Internal reference: <a href="' + esc(folderUrl) + '">client folder</a></p>',
    '</div>'
  ].join('');
}

// ════════════════════════════════════════════════════════════════════
// EMAIL — Visitor Record
// ════════════════════════════════════════════════════════════════════
function buildVisitorEmailPlain(d, folderUrl) {
  const restoration = (d.status_expired === 'Yes')
    ? 'Because your status has already expired, a restoration fee of $246.25 will also apply (in addition to the $100 IRCC fee).'
    : 'Standard IRCC fee: $100.';
  return [
    'Hello ' + (d.full_name || '') + ',',
    '',
    'Thank you for completing the Visitor Record extension intake with ' + FIRM_NAME + '.',
    '',
    'Attached is a PDF copy of the information you submitted. Please review it carefully and reply to confirm everything is accurate.',
    '',
    'Estimated cost:',
    '  • ' + restoration,
    '  • Professional fee (' + FIRM_NAME + '): $300.00',
    '  • Payment method: e-transfer to ' + FIRM_EMAIL,
    '',
    'Once you confirm, we will proceed with preparing and submitting your Visitor Record extension application to IRCC.',
    '',
    'Internal reference folder: ' + folderUrl,
    '',
    'Best regards,', FIRM_NAME, FIRM_RCIC, FIRM_EMAIL
  ].join('\n');
}

function buildVisitorEmailHtml(d, folderUrl) {
  const restoration = (d.status_expired === 'Yes')
    ? '<li>Because your status has <strong>already expired</strong>, a restoration fee of <strong>$246.25</strong> will also apply (on top of the $100 IRCC fee).</li>'
    : '<li>Standard IRCC fee: <strong>$100</strong>.</li>';
  return [
    '<div style="font-family:Helvetica,Arial,sans-serif;color:#15202b;max-width:620px;line-height:1.55;">',
      '<p>Hello <strong>' + esc(d.full_name || '') + '</strong>,</p>',
      '<p>Thank you for completing the <strong>Visitor Record extension</strong> intake with ' + esc(FIRM_NAME) + '.</p>',
      '<p>Attached is a <strong>PDF copy</strong> of the information you submitted. Please review carefully and reply to confirm everything is accurate. If a detail needs correcting, just tell us what to change — no need to fill the form again.</p>',
      '<p style="margin-top:18px;"><strong>Estimated cost:</strong></p>',
      '<ul style="margin-top:6px;">',
        restoration,
        '<li>Professional fee (' + esc(FIRM_NAME) + '): <strong>$300.00</strong></li>',
        '<li>Payment method: e-transfer to <a href="mailto:' + esc(FIRM_EMAIL) + '">' + esc(FIRM_EMAIL) + '</a></li>',
      '</ul>',
      '<p>Once you confirm, we will prepare and submit your application to IRCC.</p>',
      '<p style="margin-top:28px;">Best regards,<br><strong>' + esc(FIRM_NAME) + '</strong><br>' + esc(FIRM_RCIC) + '<br>',
        '<a href="mailto:' + esc(FIRM_EMAIL) + '">' + esc(FIRM_EMAIL) + '</a></p>',
      '<hr style="border:none;border-top:1px solid #e8edf3;margin-top:24px;">',
      '<p style="font-size:11px;color:#888;">Internal reference: <a href="' + esc(folderUrl) + '">client folder</a></p>',
    '</div>'
  ].join('');
}

// ════════════════════════════════════════════════════════════════════
// EMAIL — Study Permit Extension
// ════════════════════════════════════════════════════════════════════
function buildStudyExtensionEmailPlain(d, folderUrl) {
  const restoration = (d.status_expired === 'Yes')
    ? 'Because your status has already expired, a restoration fee of $246.25 will also apply (in addition to the $150 IRCC fee).'
    : 'Standard IRCC fee: $150 (+ $85 biometric fee if applicable).';
  return [
    'Hello ' + (d.full_name || '') + ',',
    '',
    'Thank you for completing the Study Permit extension intake with ' + FIRM_NAME + '.',
    '',
    'Attached is a PDF copy of the information you submitted. Please review it carefully and reply to confirm everything is accurate.',
    '',
    'Estimated cost:',
    '  • ' + restoration,
    '  • Professional fee (' + FIRM_NAME + '): $200.00',
    '  • Payment method: e-transfer to ' + FIRM_EMAIL,
    '',
    'Once you confirm, we will proceed with preparing and submitting your Study Permit extension application to IRCC.',
    '',
    'Internal reference folder: ' + folderUrl,
    '',
    'Best regards,', FIRM_NAME, FIRM_RCIC, FIRM_EMAIL
  ].join('\n');
}

function buildStudyExtensionEmailHtml(d, folderUrl) {
  const restoration = (d.status_expired === 'Yes')
    ? '<li>Because your status has <strong>already expired</strong>, a restoration fee of <strong>$246.25</strong> will also apply (on top of the $150 IRCC fee).</li>'
    : '<li>IRCC fee: <strong>$150</strong> (+ <strong>$85 biometric</strong> if applicable).</li>';
  return [
    '<div style="font-family:Helvetica,Arial,sans-serif;color:#15202b;max-width:620px;line-height:1.55;">',
      '<p>Hello <strong>' + esc(d.full_name || '') + '</strong>,</p>',
      '<p>Thank you for completing the <strong>Study Permit extension</strong> intake with ' + esc(FIRM_NAME) + '.</p>',
      '<p>Attached is a <strong>PDF copy</strong> of the information you submitted. Please review carefully and reply to confirm everything is accurate. If a detail needs correcting, just tell us what to change — no need to fill the form again.</p>',
      '<p style="margin-top:18px;"><strong>Estimated cost:</strong></p>',
      '<ul style="margin-top:6px;">',
        restoration,
        '<li>Professional fee (' + esc(FIRM_NAME) + '): <strong>$200.00</strong></li>',
        '<li>Payment method: e-transfer to <a href="mailto:' + esc(FIRM_EMAIL) + '">' + esc(FIRM_EMAIL) + '</a></li>',
      '</ul>',
      '<p>Once you confirm, we will prepare and submit your application to IRCC.</p>',
      '<p style="margin-top:28px;">Best regards,<br><strong>' + esc(FIRM_NAME) + '</strong><br>' + esc(FIRM_RCIC) + '<br>',
        '<a href="mailto:' + esc(FIRM_EMAIL) + '">' + esc(FIRM_EMAIL) + '</a></p>',
      '<hr style="border:none;border-top:1px solid #e8edf3;margin-top:24px;">',
      '<p style="font-size:11px;color:#888;">Internal reference: <a href="' + esc(folderUrl) + '">client folder</a></p>',
    '</div>'
  ].join('');
}

// ════════════════════════════════════════════════════════════════════
// EMAIL — Visitor Visa (TRV / Super Visa)
// ════════════════════════════════════════════════════════════════════
function buildVisitorVisaEmailPlain(d, folderUrl) {
  const isSuper = (d.super_visa === 'Yes');
  const label   = isSuper ? 'Super Visa' : 'Visitor Visa (TRV)';
  return [
    'Hello ' + (d.full_name || '') + ',',
    '',
    'Thank you for completing the ' + label + ' intake with ' + FIRM_NAME + '.',
    '',
    'Attached is a PDF copy of the information you submitted. Please review it carefully and reply to confirm everything is accurate. If any detail needs correcting, let us know which section and what should change — no need to re-fill the form.',
    '',
    'Once you confirm, we will proceed with preparing and submitting your ' + label + ' application to IRCC.',
    '',
    'Internal reference folder: ' + folderUrl,
    '',
    'Best regards,', FIRM_NAME, FIRM_RCIC, FIRM_EMAIL
  ].join('\n');
}

function buildVisitorVisaEmailHtml(d, folderUrl) {
  const isSuper = (d.super_visa === 'Yes');
  const label   = isSuper ? 'Super Visa' : 'Visitor Visa (TRV)';
  return [
    '<div style="font-family:Helvetica,Arial,sans-serif;color:#15202b;max-width:620px;line-height:1.55;">',
      '<p>Hello <strong>' + esc(d.full_name || '') + '</strong>,</p>',
      '<p>Thank you for completing the <strong>' + esc(label) + '</strong> intake with ' + esc(FIRM_NAME) + '.</p>',
      '<p>Attached is a <strong>PDF copy</strong> of the information you submitted. Please review carefully and reply to confirm everything is accurate. If a detail needs correcting, just tell us what to change — no need to fill the form again.</p>',
      '<p>Once you confirm, we will prepare and submit your <strong>' + esc(label) + '</strong> application to IRCC.</p>',
      '<p style="margin-top:28px;">Best regards,<br><strong>' + esc(FIRM_NAME) + '</strong><br>' + esc(FIRM_RCIC) + '<br>',
        '<a href="mailto:' + esc(FIRM_EMAIL) + '">' + esc(FIRM_EMAIL) + '</a></p>',
      '<hr style="border:none;border-top:1px solid #e8edf3;margin-top:24px;">',
      '<p style="font-size:11px;color:#888;">Internal reference: <a href="' + esc(folderUrl) + '">client folder</a></p>',
    '</div>'
  ].join('');
}

// ════════════════════════════════════════════════════════════════════
// EMAIL — Visitor Visa (Inside Canada)
// ════════════════════════════════════════════════════════════════════
function buildVisitorVisaIcEmailPlain(d, folderUrl) {
  return [
    'Hello ' + (d.full_name || '') + ',',
    '',
    'Thank you for completing the TRV (Inside Canada) intake with ' + FIRM_NAME + '.',
    '',
    'Attached is a PDF copy of the information you submitted. Please review it carefully and reply to confirm everything is accurate. If any detail needs correcting, let us know what to change — no need to re-fill the form.',
    '',
    'Once you confirm, we will proceed with preparing and submitting your TRV application to IRCC.',
    '',
    'Internal reference folder: ' + folderUrl,
    '',
    'Best regards,', FIRM_NAME, FIRM_RCIC, FIRM_EMAIL
  ].join('\n');
}

function buildVisitorVisaIcEmailHtml(d, folderUrl) {
  return [
    '<div style="font-family:Helvetica,Arial,sans-serif;color:#15202b;max-width:620px;line-height:1.55;">',
      '<p>Hello <strong>' + esc(d.full_name || '') + '</strong>,</p>',
      '<p>Thank you for completing the <strong>TRV (Inside Canada)</strong> intake with ' + esc(FIRM_NAME) + '.</p>',
      '<p>Attached is a <strong>PDF copy</strong> of the information you submitted. Please review carefully and reply to confirm everything is accurate. If a detail needs correcting, just tell us what to change — no need to fill the form again.</p>',
      '<p>Once you confirm, we will prepare and submit your <strong>TRV</strong> application to IRCC.</p>',
      '<p style="margin-top:28px;">Best regards,<br><strong>' + esc(FIRM_NAME) + '</strong><br>' + esc(FIRM_RCIC) + '<br>',
        '<a href="mailto:' + esc(FIRM_EMAIL) + '">' + esc(FIRM_EMAIL) + '</a></p>',
      '<hr style="border:none;border-top:1px solid #e8edf3;margin-top:24px;">',
      '<p style="font-size:11px;color:#888;">Internal reference: <a href="' + esc(folderUrl) + '">client folder</a></p>',
    '</div>'
  ].join('');
}

// ════════════════════════════════════════════════════════════════════
// HELPERS
// ════════════════════════════════════════════════════════════════════
function serialize(d, prefix, fields) {
  let out = '';
  for (let i = 1; i <= 20; i++) {
    const parts = fields.map(f => d[prefix + f + '_' + i] || '');
    if (!parts[0] && !parts[1]) break;
    out += '[' + i + '] ' + parts.join(' | ') + '\n';
  }
  return out.trim();
}

// Flatten an array of plain objects (from a repeater) into newline-delimited text for a sheet cell.
function serializeList(list, keys) {
  if (!Array.isArray(list) || !list.length) return '';
  return list.map((item, i) => {
    const parts = keys.map(k => (item && item[k] != null) ? String(item[k]) : '');
    return '[' + (i + 1) + '] ' + parts.join(' | ');
  }).join('\n');
}

// Render an array of plain objects (from a repeater) as entry-card HTML for the PDF.
function listFromArray(list, fields) {
  if (!Array.isArray(list) || !list.length) return '';
  let html = '';
  list.forEach(item => {
    let inner = '';
    fields.forEach(f => {
      const v = item && item[f[1]];
      if (v) inner += '<div><strong>' + esc(f[0]) + ':</strong> ' + esc(v) + '</div>';
    });
    if (inner) html += '<div class="entry">' + inner + '</div>';
  });
  return html;
}

function truthyFlag(v) {
  const s = String(v || '').trim().toLowerCase();
  return s === 'yes' || s === 'true' || s === 'on' || s === '1';
}

function getVisitorVisaEducationEntries(d) {
  const list = Array.isArray(d.education_list) ? d.education_list : [];
  const normalized = list.map(item => ({
    level:          item.level || item.edu_level || '',
    program:        item.program || item.edu_program || '',
    school:         item.school || item.edu_school || '',
    school_address: item.school_address || item.edu_school_address || '',
    start:          item.start || item.edu_start || '',
    end:            item.end || item.edu_end || '',
    present:        truthyFlag(item.present || item.edu_present) ? 'Yes' : '',
    country:        item.country || item.edu_country || ''
  })).filter(item =>
    item.level || item.program || item.school || item.school_address ||
    item.start || item.end || item.present || item.country
  );
  if (normalized.length) return normalized;

  const count = parseInt(d._count_education || 0, 10);
  const max = count > 0 ? count : 20;
  const out = [];
  for (let i = 1; i <= max; i++) {
    const row = {
      level:          d['edu_level_' + i] || '',
      program:        d['edu_program_' + i] || '',
      school:         d['edu_school_' + i] || '',
      school_address: d['edu_school_address_' + i] || '',
      start:          d['edu_start_' + i] || '',
      end:            d['edu_end_' + i] || '',
      present:        truthyFlag(d['edu_present_' + i]) ? 'Yes' : '',
      country:        d['edu_country_' + i] || ''
    };
    const hasData = row.level || row.program || row.school || row.school_address || row.start || row.end || row.present || row.country;
    if (!count && !hasData) break;
    if (hasData) out.push(row);
  }
  return out;
}

function getVisitorVisaEmploymentEntries(d) {
  const list = Array.isArray(d.employment_list) ? d.employment_list : [];
  const normalized = list.map(item => ({
    occupation:      item.occupation || item.emp_occupation || '',
    company:         item.company || item.emp_company || '',
    company_address: item.company_address || item.emp_company_address || '',
    start:           item.start || item.emp_start || '',
    end:             item.end || item.emp_end || '',
    current:         truthyFlag(item.current || item.emp_current) ? 'Yes' : '',
    country:         item.country || item.emp_country || '',
    duties:          item.duties || item.emp_duties || ''
  })).filter(item =>
    item.occupation || item.company || item.company_address ||
    item.start || item.end || item.current || item.country || item.duties
  );
  if (normalized.length) return normalized;

  const count = parseInt(d._count_employment || 0, 10);
  const max = count > 0 ? count : 20;
  const out = [];
  for (let i = 1; i <= max; i++) {
    const row = {
      occupation:      d['emp_occupation_' + i] || '',
      company:         d['emp_company_' + i] || '',
      company_address: d['emp_company_address_' + i] || '',
      start:           d['emp_start_' + i] || '',
      end:             d['emp_end_' + i] || '',
      current:         truthyFlag(d['emp_current_' + i]) ? 'Yes' : '',
      country:         d['emp_country_' + i] || '',
      duties:          d['emp_duties_' + i] || ''
    };
    const hasData = row.occupation || row.company || row.company_address || row.start || row.end || row.current || row.country || row.duties;
    if (!count && !hasData) break;
    if (hasData) out.push(row);
  }
  return out;
}

function getVisitorVisaAddressEntries(d) {
  const list = Array.isArray(d.address_list) ? d.address_list : [];
  const normalized = list.map(item => ({
    street:   item.street || item.addr_street || '',
    city:     item.city || item.addr_city || '',
    province: item.province || item.addr_province || '',
    from:     item.from || item.addr_from || '',
    to:       item.to || item.addr_to || '',
    present:  truthyFlag(item.present || item.addr_present) ? 'Yes' : ''
  })).filter(item =>
    item.street || item.city || item.province || item.from || item.to || item.present
  );
  if (normalized.length) return normalized;

  const count = parseInt(d._count_address || 0, 10);
  const max = count > 0 ? count : 20;
  const out = [];
  for (let i = 1; i <= max; i++) {
    const row = {
      street:   d['addr_street_' + i] || '',
      city:     d['addr_city_' + i] || '',
      province: d['addr_province_' + i] || '',
      from:     d['addr_from_' + i] || '',
      to:       d['addr_to_' + i] || '',
      present:  truthyFlag(d['addr_present_' + i]) ? 'Yes' : ''
    };
    const hasData = row.street || row.city || row.province || row.from || row.to || row.present;
    if (!count && !hasData) break;
    if (hasData) out.push(row);
  }
  return out;
}

function flattenEducation(d) {
  const entries = getVisitorVisaEducationEntries(d);
  return entries.map((row, i) => {
    const title = [row.level, row.program].filter(Boolean).join(' — ');
    const school = [row.school, row.school_address].filter(Boolean).join(', ');
    const end = row.present ? 'Present' : row.end;
    const period = [row.start, end].filter(Boolean).join(' to ');
    const country = row.country ? ('Country: ' + row.country) : '';
    return '[' + (i + 1) + '] ' + [title, school, period, country].filter(Boolean).join(' | ');
  }).join('\n');
}

function flattenEmployment(d) {
  const entries = getVisitorVisaEmploymentEntries(d);
  return entries.map((row, i) => {
    const role = [row.occupation, row.company].filter(Boolean).join(' — ');
    const end = row.current ? 'Current' : row.end;
    const period = [row.start, end].filter(Boolean).join(' to ');
    const location = [row.company_address, row.country ? ('Country: ' + row.country) : ''].filter(Boolean).join(' | ');
    const duties = row.duties ? ('Duties: ' + row.duties) : '';
    return '[' + (i + 1) + '] ' + [role, period, location, duties].filter(Boolean).join(' | ');
  }).join('\n');
}

function flattenAddressHistory(d) {
  const entries = getVisitorVisaAddressEntries(d);
  return entries.map((row, i) => {
    const location = [row.street, row.city, row.province].filter(Boolean).join(', ');
    const end = row.present ? 'Present' : row.to;
    const period = [row.from, end].filter(Boolean).join(' to ');
    return '[' + (i + 1) + '] ' + [location, period].filter(Boolean).join(' | ');
  }).join('\n');
}

function sanitize(s) {
  return String(s).replace(/[\\\/:*?"<>|]/g, '').replace(/\s+/g, ' ').trim().substring(0, 100);
}

function esc(s) {
  return String(s == null ? '' : s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/\n/g, '<br>');
}

function dataUrlToBlob(dataUrl, filename) {
  const match = dataUrl.match(/^data:(.+?);base64,(.+)$/);
  if (!match) throw new Error('Invalid data URL');
  return Utilities.newBlob(Utilities.base64Decode(match[2]), match[1], filename);
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
