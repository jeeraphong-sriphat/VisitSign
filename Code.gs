/**
 * ระบบบันทึกบุคคลภายนอกเข้าปฏิบัติงาน/ติดต่อในห้องปฏิบัติการ
 * + ระบบสรุปรายเดือนพร้อมลายเซ็นดิจิทัลของหัวหน้า
 *
 * วิธีติดตั้ง:
 * 1. เปิด Google Sheet ที่มี ID ด้านล่าง -> Extensions > Apps Script
 * 2. วางไฟล์นี้เป็น Code.gs และวาง Index.html เป็นไฟล์ HTML อีกไฟล์
 * 3. รันฟังก์ชัน setup() หนึ่งครั้ง (จะสร้างชีตสรุปรายเดือน + โฟลเดอร์เก็บลายเซ็น)
 * 4. Deploy > New deployment > Web app
 *    - Execute as: Me
 *    - Who has access: ตามนโยบายของแล็บ (แนะนำ: เฉพาะบุคคลในองค์กร)
 */

// ==================== CONFIG ====================
const SHEET_ID = '1eG6k4H8KXk5z8ZLz6QKcJxGe4ePnWNeckqj1svaKnsk';
const LOG_SHEET_NAME = 'Sheet1';           // ชีตบันทึกบุคคลภายนอก (มีอยู่แล้ว)
const SUMMARY_SHEET_NAME = 'สรุปรายเดือน';   // ชีตสรุป+ลายเซ็นหัวหน้า (จะสร้างให้อัตโนมัติ)
const SIGNATURE_FOLDER_PROP = 'SIGNATURE_FOLDER_ID';

// ข้อมูลหัวฟอร์ม อ้างอิงจาก FM-LAB-031 Rev.00
const FORM_CONFIG = {
  formCode: 'FM-LAB-031',
  revision: '00',
  approvedDate: '01/04/2567',
  department: 'แผนกปฏิบัติการกลาง',
  phone: '0-5393-4629',
  documentTitle: 'แบบบันทึกรายชื่อบุคคลภายนอกเข้า-ออกห้องปฏิบัติการศูนย์ศรีพัฒน์',
  orgName: 'ศูนย์ศรีพัฒน์ คณะแพทยศาสตร์ มหาวิทยาลัยเชียงใหม่',
  orgAddress: '110/392 อาคารศรีพัฒน์ ถนนอินทวโรรส ตำบลศรีภูมิ อำเภอเมือง จังหวัดเชียงใหม่ 50200',
  orgPhone: '0-5393-6900-1',
  signerTitleDefault: 'หัวหน้าแผนกปฏิบัติการกลาง'
};

const ACTIVITY_OPTIONS = [
  'ซ่อมบำรุง/ติดตั้งเครื่องมือ',
  'ส่งของ/พัสดุ/เอกสาร',
  'ตรวจสอบ/สอบเทียบเครื่องมือ (Calibration)',
  'ตรวจประเมิน/เยี่ยมสำรวจ (Audit/Survey)',
  'ประชุม/นัดหมาย',
  'ทำความสะอาด/บำรุงรักษาสถานที่',
  'ฝึกอบรม/ดูงาน',
  'อื่นๆ (โปรดระบุ)'
];

// ==================== SETUP ====================
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // สร้างชีตสรุปรายเดือน ถ้ายังไม่มี
  let summarySheet = ss.getSheetByName(SUMMARY_SHEET_NAME);
  if (!summarySheet) {
    summarySheet = ss.insertSheet(SUMMARY_SHEET_NAME);
    summarySheet.appendRow([
      'เดือน/ปี', 'จำนวนรายการทั้งหมด', 'ผู้ลงนามรับทราบ',
      'ตำแหน่ง', 'วันที่ลงนาม', 'ลายเซ็น'
    ]);
    summarySheet.setFrozenRows(1);
    summarySheet.getRange('A:A').setNumberFormat('@'); // กันคอลัมน์เดือน/ปีถูกแปลงเป็นวันที่อัตโนมัติ
  }

  // สร้างโฟลเดอร์เก็บรูปลายเซ็น ถ้ายังไม่มี
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(SIGNATURE_FOLDER_PROP)) {
    const folder = DriveApp.createFolder('Visitor Log - Signatures');
    props.setProperty(SIGNATURE_FOLDER_PROP, folder.getId());
  }

  Logger.log('Setup complete.');
}

function getOrCreateSignatureFolder_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty(SIGNATURE_FOLDER_PROP);
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* fallthrough */ }
  }
  const folder = DriveApp.createFolder('Visitor Log - Signatures');
  props.setProperty(SIGNATURE_FOLDER_PROP, folder.getId());
  return folder;
}

// ==================== WEB APP ENTRY ====================
function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');
  return template.evaluate()
    .setTitle('บันทึกบุคคลภายนอก - ห้องปฏิบัติการ')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ==================== ENTRY FORM ====================
function getActivityOptions() {
  return ACTIVITY_OPTIONS;
}

function getFormConfig() {
  return FORM_CONFIG;
}

/**
 * data = { date: 'yyyy-MM-dd', name: string, activity: string,
 *          activityOther: string, signatureBase64: string }
 */
function submitVisitorEntry(data) {
  if (!data.date || !data.name || !data.activity || !data.signatureBase64) {
    throw new Error('กรุณากรอกข้อมูลให้ครบถ้วน และลงลายเซ็น');
  }

  const activityText = (data.activity === 'อื่นๆ (โปรดระบุ)' && data.activityOther)
    ? 'อื่นๆ: ' + data.activityOther
    : data.activity;

  const signatureUrl = saveSignatureImage_(data.signatureBase64, data.name, data.date);

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOG_SHEET_NAME);
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1, 1, 4).setValues([[data.date, data.name, activityText, '']]);

  // ใส่รูปลายเซ็นลงในเซลล์ (IMAGE formula เพื่อให้แสดงในชีตได้เลย)
  sheet.getRange(row, 4).setFormula('=IMAGE("' + signatureUrl + '")');

  return { success: true, row: row };
}

function saveSignatureImage_(base64Data, personName, dateStr) {
  const folder = getOrCreateSignatureFolder_();
  const cleanBase64 = base64Data.split(',').pop(); // ตัด "data:image/png;base64," ออก
  const bytes = Utilities.base64Decode(cleanBase64);
  const blob = Utilities.newBlob(bytes, 'image/png',
    'sig_' + dateStr + '_' + personName.replace(/\s+/g, '_') + '_' + new Date().getTime() + '.png');
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return 'https://drive.google.com/uc?id=' + file.getId();
}

// ==================== MONTHLY SUMMARY ====================
/**
 * yearMonth = 'yyyy-MM'
 * คืนค่ารายการทั้งหมดในเดือนนั้น + สถานะว่าลงนามแล้วหรือยัง
 */
function getMonthlySummary(yearMonth) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(LOG_SHEET_NAME);
  const values = sheet.getDataRange().getValues();
  const formulas = sheet.getDataRange().getFormulas();
  const entries = [];

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (!row[0]) continue;
    const rowDate = (row[0] instanceof Date) ? row[0] : new Date(row[0]);
    const rowYm = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'yyyy-MM');
    if (rowYm === yearMonth) {
      entries.push({
        date: Utilities.formatDate(rowDate, Session.getScriptTimeZone(), 'dd/MM/yyyy'),
        name: row[1],
        activity: row[2],
        sigUrl: extractImageUrl_(formulas[i][3])
      });
    }
  }

  const signOffStatus = getSignOffStatus_(yearMonth);
  return { entries: entries, count: entries.length, signOff: signOffStatus };
}

// ดึง URL รูปจากสูตร =IMAGE("...") ในเซลล์
function extractImageUrl_(formula) {
  if (!formula) return '';
  const match = formula.match(/=IMAGE\("([^"]+)"/i);
  return match ? match[1] : '';
}

function getSignOffStatus_(yearMonth) {
  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SUMMARY_SHEET_NAME);
  if (!sheet) return { signed: false };
  const values = sheet.getDataRange().getValues();
  const formulas = sheet.getDataRange().getFormulas();
  for (let i = 1; i < values.length; i++) {
    if (normalizeYearMonth_(values[i][0]) === yearMonth) {
      return {
        signed: true,
        signerName: values[i][2],
        signerTitle: values[i][3],
        signedDate: values[i][4] instanceof Date
          ? Utilities.formatDate(values[i][4], Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')
          : values[i][4],
        sigUrl: extractImageUrl_(formulas[i][5])
      };
    }
  }
  return { signed: false };
}

// Google Sheets มักแปลงข้อความรูปแบบ "yyyy-MM" (เช่น "2026-09") เป็นวันที่ (Date) ให้อัตโนมัติ
// ฟังก์ชันนี้แปลงกลับให้เป็น string "yyyy-MM" เสมอ ไม่ว่าเซลล์จะเก็บเป็น Date หรือ Text ก็ตาม
// เพื่อให้เทียบค่ากับ yearMonth ที่ส่งมาจาก client ได้ถูกต้อง
function normalizeYearMonth_(val) {
  if (val instanceof Date) {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM');
  }
  return String(val);
}

/**
 * data = { yearMonth, signerName, signerTitle, signatureBase64 }
 */
function submitMonthlySignOff(data) {
  if (!data.yearMonth || !data.signerName || !data.signatureBase64) {
    throw new Error('ข้อมูลไม่ครบถ้วน');
  }

  const existing = getSignOffStatus_(data.yearMonth);
  if (existing && existing.signed) {
    throw new Error('เดือนนี้มีการลงนามรับทราบแล้ว โดย ' + existing.signerName);
  }

  const summary = getMonthlySummary(data.yearMonth);
  const signatureUrl = saveSignatureImage_(data.signatureBase64, data.signerName, data.yearMonth);

  const sheet = SpreadsheetApp.openById(SHEET_ID).getSheetByName(SUMMARY_SHEET_NAME);
  const row = sheet.getLastRow() + 1;
  sheet.getRange(row, 1).setNumberFormat('@'); // บังคับให้เป็นข้อความ กัน Sheets แปลงเป็นวันที่อัตโนมัติ
  sheet.getRange(row, 1, 1, 5).setValues([[
    data.yearMonth, summary.count, data.signerName,
    data.signerTitle || FORM_CONFIG.signerTitleDefault, new Date()
  ]]);
  sheet.getRange(row, 6).setFormula('=IMAGE("' + signatureUrl + '")');

  return { success: true };
}

// ==================== EXPORT PDF ====================
/**
 * สร้างไฟล์ PDF รายงานประจำเดือน โดยใช้ Google Sheet "ReportForm" เป็นต้นแบบ
 * (ต้องลงนามรับทราบแล้วเท่านั้น)
 * yearMonth = 'yyyy-MM'
 * คืนค่า { base64, filename } ให้ฝั่ง client แปลงเป็นไฟล์ดาวน์โหลด
 */
const REPORT_TEMPLATE_SHEET_NAME = 'ReportForm';

function exportMonthlyReportPDF(yearMonth) {
  const summary = getMonthlySummary(yearMonth);

  if (!summary.signOff || !summary.signOff.signed) {
    throw new Error('เดือนนี้ยังไม่ได้ลงนามรับทราบ กรุณาให้หัวหน้าลงนามก่อน จึงจะ Export รายงานได้');
  }

  const srcSs = SpreadsheetApp.openById(SHEET_ID);
  const templateSheet = srcSs.getSheetByName(REPORT_TEMPLATE_SHEET_NAME);
  if (!templateSheet) {
    throw new Error('ไม่พบชีตต้นแบบชื่อ "' + REPORT_TEMPLATE_SHEET_NAME + '" กรุณาตรวจสอบชื่อชีตอีกครั้ง');
  }

  // สร้างสเปรดชีตชั่วคราว แล้วคัดลอกชีตต้นแบบเข้าไป (ไม่แก้ไขต้นแบบจริง)
  const tempSs = SpreadsheetApp.create(
    'TEMP_' + FORM_CONFIG.formCode + '_' + yearMonth + '_' + new Date().getTime()
  );
  const tempSheet = templateSheet.copyTo(tempSs);

  // ลบชีตเริ่มต้น (Sheet1) ของไฟล์ใหม่ทิ้ง เหลือแค่ชีตที่คัดลอกมา
  tempSs.getSheets().forEach(function (s) {
    if (s.getSheetId() !== tempSheet.getSheetId()) {
      tempSs.deleteSheet(s);
    }
  });

  try {
    fillReportTemplate_(tempSheet, yearMonth, summary);
    SpreadsheetApp.flush();

    const pdfBytes = exportSheetAsPdf_(tempSs.getId(), tempSheet.getSheetId());
    const filename = FORM_CONFIG.formCode + '_สรุปรายเดือน_' + yearMonth + '.pdf';

    return {
      base64: Utilities.base64Encode(pdfBytes),
      filename: filename
    };
  } finally {
    // ลบไฟล์ชั่วคราวทิ้งเสมอ ไม่ว่าจะสำเร็จหรือ error เพื่อไม่ให้ Drive รก
    DriveApp.getFileById(tempSs.getId()).setTrashed(true);
  }
}

// เติมข้อมูลจริงลงในชีตสำเนา แทนที่ placeholder ทั้งหมด
function fillReportTemplate_(sheet, yearMonth, summary) {
  replaceAllText_(sheet, '{{MONTH}}', yearMonth);

  const dateCell = sheet.createTextFinder('{{DATE}}').matchEntireCell(false).findNext();
  if (!dateCell) {
    throw new Error('ไม่พบ {{DATE}} ในชีตต้นแบบ กรุณาตรวจสอบว่าใส่ placeholder ไว้ถูกต้อง');
  }
  const templateRow = dateCell.getRow();
  const dateCol = dateCell.getColumn();
  const lastCol = sheet.getLastColumn();

  const nameCell = sheet.createTextFinder('{{NAME}}').matchEntireCell(false).findNext();
  const sigCell = sheet.createTextFinder('{{SIGNATURE}}').matchEntireCell(false).findNext();
  const actCell = sheet.createTextFinder('{{ACTIVITY}}').matchEntireCell(false).findNext();
  if (!nameCell || !sigCell || !actCell) {
    throw new Error('ไม่พบ placeholder {{NAME}}, {{SIGNATURE}} หรือ {{ACTIVITY}} ในแถวข้อมูล กรุณาตรวจสอบชีตต้นแบบ');
  }
  const nameCol = nameCell.getColumn();
  const sigCol = sigCell.getColumn();
  const actCol = actCell.getColumn();

  const entries = summary.entries;
  const n = entries.length;

  if (n === 0) {
    sheet.getRange(templateRow, dateCol).setValue('');
    sheet.getRange(templateRow, nameCol).setValue('');
    sheet.getRange(templateRow, actCol).setValue('');
    sheet.getRange(templateRow, sigCol).setValue('');
  } else {
    if (n > 1) {
      sheet.insertRowsAfter(templateRow, n - 1);
      const srcRange = sheet.getRange(templateRow, 1, 1, lastCol);
      for (let i = 1; i < n; i++) {
        srcRange.copyTo(sheet.getRange(templateRow + i, 1, 1, lastCol));
      }
    }
    for (let i = 0; i < n; i++) {
      const r = templateRow + i;
      const e = entries[i];
      sheet.getRange(r, dateCol).setValue(e.date);
      sheet.getRange(r, nameCol).setValue(e.name);
      sheet.getRange(r, actCol).setValue(e.activity);
      if (e.sigUrl) {
        sheet.getRange(r, sigCol).setFormula('=IMAGE("' + e.sigUrl + '")');
      } else {
        sheet.getRange(r, sigCol).setValue('');
      }
    }
  }

  // จุดลงนามหัวหน้า (ค้นหาใหม่หลังแทรกแถว เผื่อตำแหน่งเลื่อนลง)
  replaceAllText_(sheet, '{{SIGNER_TITLE}}', summary.signOff.signerTitle || FORM_CONFIG.signerTitleDefault);
  replaceAllText_(sheet, '{{SIGN_DATE}}', summary.signOff.signedDate || '');

  const signOffCell = sheet.createTextFinder('{{SIGNOFF_SIGNATURE}}').matchEntireCell(false).findNext();
  if (signOffCell) {
    if (summary.signOff.sigUrl) {
      signOffCell.setFormula('=IMAGE("' + summary.signOff.sigUrl + '")');
    } else {
      signOffCell.setValue('');
    }
  }
}

function replaceAllText_(sheet, searchText, replaceText) {
  sheet.createTextFinder(searchText).matchEntireCell(false).replaceAllWith(replaceText);
}

// เรียก export URL ของ Google Sheets เพื่อดึงเฉพาะชีตที่ต้องการเป็น PDF
function exportSheetAsPdf_(spreadsheetId, sheetId) {
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export' +
    '?format=pdf&gid=' + sheetId +
    '&size=A4&portrait=true&fitw=true&scale=4' +
    '&top_margin=0.4&bottom_margin=0.4&left_margin=0.4&right_margin=0.4' +
    '&gridlines=false&printtitle=false&sheetnames=false&pagenum=UNDEFINED&horizontal_alignment=CENTER';

  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error('Export PDF ไม่สำเร็จ (HTTP ' + response.getResponseCode() + ')');
  }
  return response.getBlob().getBytes();
}
