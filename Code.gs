/*****************************************************************
 * BBCS-AA Membership Registration — Google Apps Script Backend
 *
 * Responsibilities:
 *   1. Receive JSON payload (POST) from the static web app.
 *   2. Save the 3 uploaded files (NID image, profile picture,
 *      payment proof) to dedicated Google Drive folders.
 *   3. Append one row per submission to a Google Sheet, including
 *      the public Drive URLs of the uploaded files.
 *   4. Generate a PDF confirmation receipt of the submitted data and
 *      email it to the applicant's own address.
 *   5. Return a JSON response the front end can read.
 *
 * See the "SETUP & DEPLOYMENT" instructions at the bottom of this
 * file for the exact steps to wire this up to your own Sheet/Drive.
 *****************************************************************/

// ---- CONFIGURATION -------------------------------------------------

// Leave blank ("") to auto-create/use a spreadsheet bound to this
// script, or paste a specific Spreadsheet ID to target an existing file.
const SPREADSHEET_ID = "";

const SHEET_NAME = "Registrations";

// Parent Drive folder name that will contain the three subfolders below.
// Leave blank to create the subfolders directly in "My Drive".
const DRIVE_ROOT_FOLDER_NAME = "BBCS-AA Membership Uploads";

const FOLDER_NAMES = {
  nidImage: "NID Images",
  profilePic: "Profile Pictures",
  paymentProof: "Payment Proofs",
};

const SHEET_HEADERS = [
  "Timestamp",
  "Name (Bangla)",
  "Name (English)",
  "Father's Name",
  "Mother's Name",
  "Date of Birth",
  "Nationality",
  "Blood Group",
  "Year of SSC",
  "Duration of Study",
  "Occupation & Designation",
  "Organisation & Duration",
  "Contact No (WhatsApp)",
  "Email",
  "Present Address",
  "Permanent Address",
  "NID Number",
  "Membership Type",
  "Membership Fee",
  "Payment Method",
  "Declaration Accepted",
  "NID Image URL",
  "Profile Picture URL",
  "Payment Proof URL",
];

// Display name the confirmation email is sent "from" (the actual address
// is always the Google account this script is deployed/authorized under).
const EMAIL_SENDER_NAME = "Bagerhat Bohumukhi Collegiate School Alumni Association";

// Set to false to skip PDF generation + emailing entirely (Sheet/Drive
// logging will still happen as normal).
const SEND_CONFIRMATION_EMAIL = true;

// ---------------------------------------------------------------------

/**
 * Handles POST requests from the front end.
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return jsonResponse({ status: "error", message: "Empty request body." });
    }

    const data = JSON.parse(e.postData.contents);
    validatePayload(data);

    const folders = getOrCreateUploadFolders_();

    // Decode each uploaded file to a Blob once, so the same Blob can be
    // both saved to Drive and (for the profile picture) embedded in the
    // PDF receipt, without re-decoding the base64 twice.
    const nidBlob = decodeFileBlob_(data.files && data.files.nidImage);
    const profileBlob = decodeFileBlob_(data.files && data.files.profilePic);
    const paymentBlob = decodeFileBlob_(data.files && data.files.paymentProof);

    const nidUrl = saveBlobToDrive_(nidBlob, folders.nidImage, data.nameEn, "NID");
    const profileUrl = saveBlobToDrive_(profileBlob, folders.profilePic, data.nameEn, "Profile");
    const paymentUrl = saveBlobToDrive_(paymentBlob, folders.paymentProof, data.nameEn, "Payment");

    const urls = { nidUrl, profileUrl, paymentUrl };
    appendRowToSheet_(data, urls);

    let emailStatus = "skipped";
    if (SEND_CONFIRMATION_EMAIL && data.email) {
      try {
        const pdfBlob = buildReceiptPdf_(data, urls, profileBlob);
        sendConfirmationEmail_(data, pdfBlob);
        emailStatus = "sent";
      } catch (mailErr) {
        // A failed email must never fail the whole registration — the
        // submission is already safely logged in Drive/Sheets by this point.
        emailStatus = "failed: " + (mailErr.message || mailErr);
      }
    }

    return jsonResponse({
      status: "success",
      message: "Registration saved.",
      urls: urls,
      email: emailStatus,
    });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message || String(err) });
  }
}

/**
 * Simple GET handler — useful to confirm the deployment is live
 * by visiting the Web App URL directly in a browser.
 */
function doGet() {
  return jsonResponse({ status: "ok", message: "BBCS-AA registration endpoint is live." });
}

/**
 * Builds a JSON text response. Apps Script web apps deployed with
 * "Anyone" access automatically allow cross-origin reads of this
 * response for simple (non-preflighted) requests — which is why the
 * front end posts with a text/plain Content-Type to avoid a CORS
 * preflight OPTIONS request (Apps Script cannot answer OPTIONS).
 */
function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Basic required-field validation, mirroring the client-side checks.
 */
function validatePayload(data) {
  const required = ["nameBn", "nameEn", "contactNo", "presentAddress", "nidNumber", "membershipType", "paymentMethod"];
  const missing = required.filter((key) => !data[key]);
  if (missing.length) {
    throw new Error("Missing required field(s): " + missing.join(", "));
  }
  if (!data.files || !data.files.nidImage || !data.files.profilePic || !data.files.paymentProof) {
    throw new Error("One or more required file uploads are missing.");
  }
}

/**
 * Returns (creating if necessary) the three upload subfolders,
 * nested inside DRIVE_ROOT_FOLDER_NAME.
 */
function getOrCreateUploadFolders_() {
  const root = DRIVE_ROOT_FOLDER_NAME
    ? getOrCreateFolder_(DriveApp.getRootFolder(), DRIVE_ROOT_FOLDER_NAME)
    : DriveApp.getRootFolder();

  return {
    nidImage: getOrCreateFolder_(root, FOLDER_NAMES.nidImage),
    profilePic: getOrCreateFolder_(root, FOLDER_NAMES.profilePic),
    paymentProof: getOrCreateFolder_(root, FOLDER_NAMES.paymentProof),
  };
}

function getOrCreateFolder_(parent, name) {
  const existing = parent.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();
  return parent.createFolder(name);
}

/**
 * Decodes a { base64, mimeType, fileName } object into an in-memory Blob.
 * Returns null if no file object was supplied.
 */
function decodeFileBlob_(fileObj) {
  if (!fileObj || !fileObj.base64) return null;
  const bytes = Utilities.base64Decode(fileObj.base64);
  const extension = guessExtension_(fileObj.mimeType, fileObj.fileName);
  const name = fileObj.fileName || ("file" + extension);
  return Utilities.newBlob(bytes, fileObj.mimeType || "application/octet-stream", name);
}

/**
 * Saves a decoded Blob into the given Drive folder, made viewable via
 * link, and returns its shareable URL. Returns "" if blob is null.
 */
function saveBlobToDrive_(blob, folder, personName, label) {
  if (!blob) return "";

  const safeName = (personName || "Unknown").replace(/[^\w\-]+/g, "_");
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT", "yyyyMMdd_HHmmss");
  const extension = guessExtension_(blob.getContentType(), blob.getName());
  const fileName = `${safeName}_${label}_${timestamp}${extension}`;

  const file = folder.createFile(blob.copyBlob().setName(fileName));
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return file.getUrl();
}

function guessExtension_(mimeType, originalName) {
  if (originalName && originalName.includes(".")) {
    return "." + originalName.split(".").pop();
  }
  const map = {
    "image/jpeg": ".jpg",
    "image/png": ".png",
    "image/webp": ".webp",
    "application/pdf": ".pdf",
  };
  return map[mimeType] || "";
}

/**
 * Appends a new row to the target sheet, creating the sheet and
 * header row on first use.
 */
function appendRowToSheet_(data, urls) {
  const sheet = getOrCreateSheet_();

  sheet.appendRow([
    new Date(),
    data.nameBn || "",
    data.nameEn || "",
    data.fatherName || "",
    data.motherName || "",
    data.dob || "",
    data.nationality || "",
    data.bloodGroup || "",
    data.sscYear || "",
    data.studyDuration || "",
    data.occupation || "",
    data.organisation || "",
    data.contactNo || "",
    data.email || "",
    data.presentAddress || "",
    data.permanentAddress || "",
    data.nidNumber || "",
    data.membershipType || "",
    data.membershipFee || "",
    data.paymentMethod || "",
    data.declaration ? "Yes" : "No",
    urls.nidUrl || "",
    urls.profileUrl || "",
    urls.paymentUrl || "",
  ]);
}

function getOrCreateSheet_() {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet() || SpreadsheetApp.create("BBCS-AA Membership Registrations");

  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(SHEET_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, SHEET_HEADERS.length).setFontWeight("bold");
  }
  return sheet;
}

/**
 * Builds a one-page PDF receipt of the submitted data (optionally with
 * the applicant's profile picture at the top) by temporarily creating a
 * Google Doc, exporting it to PDF, then discarding the Doc. This is the
 * standard Apps Script pattern for HTML/data → PDF, since Blob objects
 * cannot be converted to PDF directly.
 */
function buildReceiptPdf_(data, urls, profileBlob) {
  const doc = DocumentApp.create("TEMP_BBCS_Receipt_" + new Date().getTime());
  const body = doc.getBody();
  body.setMarginTop(40).setMarginBottom(40).setMarginLeft(50).setMarginRight(50);

  const title = body.appendParagraph("Bagerhat Bohumukhi Collegiate School Alumni Association");
  title.setHeading(DocumentApp.ParagraphHeading.HEADING2);
  title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);

  const subtitle = body.appendParagraph("Membership Registration — Confirmation Copy");
  subtitle.setHeading(DocumentApp.ParagraphHeading.HEADING4);
  subtitle.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
  subtitle.setSpacingAfter(14);

  if (profileBlob) {
    try {
      const imgParagraph = body.appendParagraph("");
      imgParagraph.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
      const image = imgParagraph.appendInlineImage(profileBlob.copyBlob());
      const maxWidth = 110;
      const ratio = maxWidth / image.getWidth();
      image.setWidth(maxWidth);
      image.setHeight(Math.round(image.getHeight() * ratio));
      body.appendParagraph(" ").setSpacingAfter(6);
    } catch (imgErr) {
      // If the image can't be embedded (unsupported format, etc.) the
      // receipt still generates fine without it.
    }
  }

  const rows = [
    ["Submission Date", Utilities.formatDate(new Date(), Session.getScriptTimeZone() || "GMT", "dd MMM yyyy, hh:mm a")],
    ["Name (Bangla)", data.nameBn],
    ["Name (English)", data.nameEn],
    ["Father's Name", data.fatherName],
    ["Mother's Name", data.motherName],
    ["Date of Birth", data.dob],
    ["Nationality", data.nationality],
    ["Blood Group", data.bloodGroup],
    ["Year of SSC", data.sscYear],
    ["Duration of Study", data.studyDuration],
    ["Occupation & Designation", data.occupation],
    ["Organisation & Duration", data.organisation],
    ["Contact No (WhatsApp)", data.contactNo],
    ["Email", data.email],
    ["Present Address", data.presentAddress],
    ["Permanent Address", data.permanentAddress],
    ["NID Number", data.nidNumber],
    ["Membership Type", data.membershipType],
    ["Membership Fee (BDT)", data.membershipFee],
    ["Payment Method", data.paymentMethod],
  ];

  const table = body.appendTable();
  rows.forEach((row) => {
    const tr = table.appendTableRow();
    const labelCell = tr.appendTableCell(row[0]);
    labelCell.setWidth(170);
    labelCell.getChild(0).asParagraph().setBold(true).setFontSize(9);
    const valueCell = tr.appendTableCell(String(row[1] || "—"));
    valueCell.getChild(0).asParagraph().setFontSize(9);
  });

  body.appendParagraph(" ").setSpacingBefore(10);
  const docsNote = body.appendParagraph(
    "Uploaded documents are stored securely. NID image and payment proof links:"
  );
  docsNote.setFontSize(9).setItalic(true);
  body.appendParagraph("NID Image: " + (urls.nidUrl || "—")).setFontSize(8);
  body.appendParagraph("Payment Proof: " + (urls.paymentUrl || "—")).setFontSize(8);

  const footer = body.appendParagraph(
    "This is an automatically generated confirmation. Please keep this copy for your records."
  );
  footer.setItalic(true).setFontSize(8).setSpacingBefore(14);

  doc.saveAndClose();

  const pdfBlob = DriveApp.getFileById(doc.getId())
    .getAs(MimeType.PDF)
    .setName((data.nameEn || "Applicant").replace(/[^\w\-]+/g, "_") + "_BBCS-AA_Receipt.pdf");

  // The temporary Google Doc is only a means to generate the PDF —
  // discard it so it doesn't clutter Drive.
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  return pdfBlob;
}

/**
 * Emails the PDF receipt to the applicant's own address.
 */
function sendConfirmationEmail_(data, pdfBlob) {
  const subject = "BBCS-AA Membership Application Received — " + data.nameEn;

  const htmlBody = `
    <div style="font-family: Arial, sans-serif; color:#241C13; max-width:560px;">
      <h2 style="color:#142E52; margin-bottom:4px;">Bagerhat Bohumukhi Collegiate School Alumni Association</h2>
      <p style="color:#5B5041; margin-top:0;">স্মৃতির আল্পনায়, প্রিয় আঙ্গিনায় — ঐতিহ্য · সম্প্রীতি · বন্ধন</p>
      <p>প্রিয় ${escapeHtml_(data.nameEn)},</p>
      <p>আপনার সদস্যপদের আবেদন সফলভাবে গৃহীত হয়েছে। আপনার পূরণকৃত তথ্যের একটি PDF কপি এই ইমেইলের সাথে সংযুক্ত করা হলো, ভবিষ্যতে রেফারেন্সের জন্য এটি সংরক্ষণ করুন।</p>
      <p><strong>সদস্যপদের ধরন:</strong> ${escapeHtml_(data.membershipType)}<br>
      <strong>পেমেন্ট পদ্ধতি:</strong> ${escapeHtml_(data.paymentMethod)}</p>
      <p>আপনার প্রদত্ত তথ্য যাচাইয়ের পর কর্তৃপক্ষ প্রয়োজনে যোগাযোগ করবে।</p>
      <p style="margin-top:24px; color:#5B5041; font-size:0.85em;">এটি একটি স্বয়ংক্রিয় বার্তা, দয়া করে এই ইমেইলের সরাসরি উত্তর দেবেন না।</p>
    </div>
  `;

  MailApp.sendEmail({
    to: data.email,
    subject: subject,
    htmlBody: htmlBody,
    name: EMAIL_SENDER_NAME,
    attachments: [pdfBlob],
  });
}

/**
 * Minimal HTML-escaping for values interpolated into the email body.
 */
function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

/*****************************************************************
 * SETUP & DEPLOYMENT INSTRUCTIONS
 * -----------------------------------------------------------------
 * 1. CREATE THE PROJECT
 *    - Go to https://script.google.com → "New project".
 *    - Delete the default code and paste this entire file's contents.
 *    - Rename the project (top-left) to e.g. "BBCS-AA Registration Backend".
 *
 * 2. (OPTIONAL) TARGET A SPECIFIC SHEET
 *    - By default this script will use the spreadsheet it is bound
 *      to, or create a new one called "BBCS-AA Membership
 *      Registrations" the first time it runs.
 *    - To use an existing Google Sheet instead: open that Sheet,
 *      copy the ID from its URL (the long string between /d/ and
 *      /edit), and paste it into SPREADSHEET_ID above.
 *
 * 3. SAVE AND AUTHORIZE
 *    - Click the disk icon (Ctrl/Cmd+S) to save.
 *    - Click "Run" ▶ once on any function (e.g. doGet) to trigger
 *      the permission prompt. Approve access to Drive and Sheets
 *      for your Google account.
 *
 * 4. DEPLOY AS WEB APP
 *    - Click "Deploy" → "New deployment".
 *    - Click the gear icon next to "Select type" → choose "Web app".
 *    - Description: "BBCS-AA registration endpoint".
 *    - "Execute as": Me (your account).
 *    - "Who has access": Anyone.  (Required so the public form can
 *      reach it — no Google login is exposed to end users.)
 *    - Click "Deploy", then authorize again if prompted.
 *    - Copy the generated "Web app URL" — it looks like:
 *        https://script.google.com/macros/s/AKfycb.../exec
 *
 * 5. CONNECT THE FRONT END
 *    - Open app.js and paste that URL into the SCRIPT_URL constant
 *      at the top of the file.
 *
 * 6. RE-DEPLOYING AFTER CHANGES
 *    - Any time you edit Code.gs, go to "Deploy" → "Manage
 *      deployments" → pencil/edit icon → change "Version" to
 *      "New version" → "Deploy". Editing the code alone does NOT
 *      update the live Web App URL's behaviour until you do this.
 *
 * 7. TESTING
 *    - Visit the Web App URL directly in a browser — you should see
 *      {"status":"ok","message":"BBCS-AA registration endpoint is live."}
 *    - Submit a test entry from the deployed website and confirm a
 *      new row appears in the "Registrations" sheet, and that the
 *      three files appear inside "BBCS-AA Membership Uploads" in
 *      Google Drive (My Drive of the account that deployed the script).
 *    - Check the test email address's inbox for the confirmation
 *      message with the PDF receipt attached (also check Spam the
 *      first time, since the sending address is new to that inbox).
 *
 * NOTES ON THE PDF RECEIPT & CONFIRMATION EMAIL
 *    - The email is sent by MailApp.sendEmail from the Google account
 *      that owns this deployment (the one you authorized in step 3),
 *      shown to recipients as "Bagerhat Bohumukhi Collegiate School
 *      Alumni Association" (see EMAIL_SENDER_NAME above).
 *    - MailApp has a daily sending quota: 100 emails/day for a normal
 *      @gmail.com account, or a higher quota for Google Workspace
 *      accounts. Each registration uses exactly one email.
 *    - The PDF is built by briefly creating a Google Doc, exporting it
 *      to PDF, then trashing the Doc — this needs no extra setup, but
 *      the first run will prompt you to authorize Google Docs access
 *      alongside Drive/Sheets in step 3.
 *    - If a PDF or email step fails for any reason (e.g. quota
 *      exceeded), the registration itself is NOT lost: the Drive
 *      files and Sheet row are already saved before the email is
 *      attempted, and the JSON response includes an "email" field
 *      ("sent", "skipped", or "failed: ...") so you can see the
 *      outcome in the browser console.
 *    - To turn the email off entirely (e.g. during testing), set
 *      SEND_CONFIRMATION_EMAIL to false above.
 *
 * NOTES ON CORS
 *    - Apps Script Web Apps cannot set custom CORS headers or answer
 *      OPTIONS preflight requests. To avoid triggering a preflight,
 *      app.js sends the JSON payload with
 *      "Content-Type: text/plain;charset=utf-8" — the body is still
 *      valid JSON and is parsed normally via JSON.parse in doPost.
 *      Do not add extra custom headers to the fetch() call in app.js,
 *      or the browser will attempt a preflight that Apps Script
 *      cannot fulfil.
 *****************************************************************/
