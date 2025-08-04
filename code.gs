// Global constants for sheet names, folder name, and headers
const MEDIA_SHEET_NAME = 'สื่อการสอน';
const TEACHER_SHEET_NAME = 'ครู';
const SETTINGS_SHEET_NAME = 'ตั้งค่าระบบ';
const UPLOAD_FOLDER_NAME = 'ไฟล์ภาพสื่อการสอน';

const MEDIA_HEADERS = ['ลำดับที่', 'รหัสอ้างอิง', 'วันที่ผลิตสื่อ', 'รายการสื่อ/อุปกรณ์', 'จำนวน', 'ประเภทสื่อ', 'ลิงก์ภาพสื่อ', 'ID ภาพสื่อ', 'ครูผู้ผลิตสื่อ', 'หมายเหตุ', 'เดือน', 'ปี'];
const TEACHER_HEADERS = ['รหัสครู', 'ชื่อ-สกุล', 'ตำแหน่ง'];
const SETTINGS_HEADERS = ['Key', 'Value'];

/**
 * @summary Main function to serve the web app.
 * @param {object} e - The event parameter.
 * @returns {HtmlOutput} The HTML service output.
 */
function doGet(e) {
  checkAndSetup(); // Check and create necessary sheets/folders on load
  const html = HtmlService.createHtmlOutputFromFile('index');
  html.setTitle('ระบบทะเบียนสื่อ โรงเรียนบ้านนานวล');
  html.setFaviconUrl('https://img5.pic.in.th/file/secure-sv1/273218374_306049724897300_8948544915894559738_n.png');
  return html;
}

/**
 * @summary Checks for and sets up required sheets and the upload folder if they don't exist.
 */
function checkAndSetup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Check for MEDIA_SHEET_NAME
  if (!ss.getSheetByName(MEDIA_SHEET_NAME)) {
    const sheet = ss.insertSheet(MEDIA_SHEET_NAME);
    sheet.appendRow(MEDIA_HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange("B:B").setNumberFormat('@'); // Set reference ID column to plain text
    sheet.getRange("A:A").setNumberFormat('000'); // Set sequence number format
  }

  // Check for TEACHER_SHEET_NAME
  if (!ss.getSheetByName(TEACHER_SHEET_NAME)) {
    const sheet = ss.insertSheet(TEACHER_SHEET_NAME);
    sheet.appendRow(TEACHER_HEADERS);
    sheet.setFrozenRows(1);
    // Add some sample teachers
    sheet.appendRow([Utilities.getUuid(), 'นายคุณากร ธนที', 'ครู']);
    sheet.appendRow([Utilities.getUuid(), 'นางสาวสมหญิง ใจดี', 'ครูผู้ช่วย']);
  }
  
  // Check for SETTINGS_SHEET_NAME
  if (!ss.getSheetByName(SETTINGS_SHEET_NAME)) {
    const sheet = ss.insertSheet(SETTINGS_SHEET_NAME);
    sheet.appendRow(SETTINGS_HEADERS);
    sheet.setFrozenRows(1);
    sheet.appendRow(['schoolName', 'โรงเรียนบ้านนานวล']);
    sheet.appendRow(['districtName', 'สำนักงานเขตพื้นที่การศึกษาประถมศึกษาสุรินทร์ เขต 2 ']);
  }

  // Check for Google Drive Folder
  const folders = DriveApp.getFoldersByName(UPLOAD_FOLDER_NAME);
  if (!folders.hasNext()) {
    DriveApp.createFolder(UPLOAD_FOLDER_NAME);
  }
}

/**
 * @summary Gets the ID of the upload folder.
 * @returns {string} The folder ID.
 */
function getUploadFolderId() {
    const folders = DriveApp.getFoldersByName(UPLOAD_FOLDER_NAME);
    if (folders.hasNext()) {
        return folders.next().getId();
    } else {
        // If folder was deleted, recreate it
        return DriveApp.createFolder(UPLOAD_FOLDER_NAME).getId();
    }
}

/**
 * @summary Gets all necessary initial data for the web app.
 * @returns {object} An object containing media, teachers, settings, and the next media ID.
 */
function getInitialData() {
  try {
    const media = getMediaData();
    const teachers = getTeachers();
    const settings = getSettings();
    const nextId = getNextMediaId();
    
    return {
      success: true,
      media: media,
      teachers: teachers,
      settings: settings,
      nextId: nextId
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * @summary Retrieves all media data from the spreadsheet.
 * @returns {Array<object>} An array of media objects.
 */
function getMediaData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEDIA_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  return data.map(row => {
    let obj = {};
    headers.forEach((header, index) => {
      // Format date for display
      if (header === 'วันที่ผลิตสื่อ' && row[index] instanceof Date) {
        obj[header] = Utilities.formatDate(row[index], Session.getScriptTimeZone(), 'yyyy-MM-dd');
      } else {
        obj[header] = row[index];
      }
    });
    return obj;
  });
}

/**
 * @summary Retrieves all teacher data from the spreadsheet.
 * @returns {Array<object>} An array of teacher objects.
 */
function getTeachers() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEACHER_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  
  return data.filter(row => row[0]).map(row => { // Filter out empty rows
    let obj = {};
    headers.forEach((header, index) => {
      obj[header] = row[index];
    });
    return obj;
  });
}

/**
 * @summary Retrieves system settings from the spreadsheet.
 * @returns {object} An object containing system settings.
 */
function getSettings() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  data.shift(); // remove headers
  let settings = {};
  data.forEach(row => {
    if(row[0]) {
      settings[row[0]] = row[1];
    }
  });
  return settings;
}

/**
 * @summary Gets the next available ID for a new media item.
 * @returns {number} The next sequential media ID.
 */
function getNextMediaId() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MEDIA_SHEET_NAME);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    return 1;
  }
  // Get value from the 'ลำดับที่' column (index 1) of the last row
  const lastId = sheet.getRange(lastRow, 1).getValue();
  return !isNaN(lastId) ? parseInt(lastId) + 1 : lastRow;
}


/**
 * @summary Adds a new media record to the sheet, including uploading an image.
 * @param {object} formData - The form data from the client.
 * @param {object} fileData - The file data (base64 string, mimeType, name).
 * @returns {object} A success or error object.
 */
function addNewMedia(formData, fileData) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const mediaSheet = ss.getSheetByName(MEDIA_SHEET_NAME);
    
    let fileUrl = '';
    let fileId = '';
    
    // Handle file upload
    if (fileData && fileData.base64) {
      const folderId = getUploadFolderId();
      const folder = DriveApp.getFolderById(folderId);
      const decoded = Utilities.base64Decode(fileData.base64, Utilities.Charset.UTF_8);
      const blob = Utilities.newBlob(decoded, fileData.mimeType, fileData.name);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileUrl = `https://lh3.googleusercontent.com/d/${file.getId()}`; // Direct link for img src
      fileId = file.getId();
    }
    
    const productionDate = new Date(formData.productionDate);
    const month = Utilities.formatDate(productionDate, Session.getScriptTimeZone(), 'MMMM');
    const year = Utilities.formatDate(productionDate, Session.getScriptTimeZone(), 'yyyy');

    const newRow = [
      formData.sequence,
      Utilities.getUuid(),
      productionDate,
      formData.mediaName,
      formData.quantity,
      formData.mediaType,
      fileUrl,
      fileId,
      formData.teacher,
      formData.notes,
      month,
      year
    ];
    
    mediaSheet.appendRow(newRow);
    
    return { success: true, message: 'บันทึกข้อมูลสื่อเรียบร้อยแล้ว' };

  } catch (e) {
    Logger.log(e);
    return { success: false, error: 'เกิดข้อผิดพลาดในการบันทึก: ' + e.message };
  }
}

/**
 * @summary Updates the system settings.
 * @param {object} newSettings - An object with new settings values.
 * @returns {object} A success or error object.
 */
function updateSettings(newSettings) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    data.forEach((row, index) => {
      if (index > 0) { // Skip header row
        const key = row[0];
        if (newSettings.hasOwnProperty(key)) {
          sheet.getRange(index + 1, 2).setValue(newSettings[key]);
        }
      }
    });
    
    return { success: true, message: 'บันทึกการตั้งค่าเรียบร้อยแล้ว' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * @summary Adds a new teacher to the system.
 * @param {object} teacherData - An object with teacher's name and position.
 * @returns {object} A success or error object with the new list of teachers.
 */
function addTeacher(teacherData) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEACHER_SHEET_NAME);
    const newId = Utilities.getUuid();
    sheet.appendRow([newId, teacherData.name, teacherData.position]);
    return { success: true, message: 'เพิ่มข้อมูลครูเรียบร้อยแล้ว', teachers: getTeachers() };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * @summary Deletes a teacher from the system.
 * @param {string} teacherId - The unique ID of the teacher to delete.
 * @returns {object} A success or error object with the new list of teachers.
 */
function deleteTeacher(teacherId) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEACHER_SHEET_NAME);
    const data = sheet.getDataRange().getValues();
    
    for (let i = data.length - 1; i >= 1; i--) {
      if (data[i][0] === teacherId) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'ลบข้อมูลครูเรียบร้อยแล้ว', teachers: getTeachers() };
      }
    }
    
    return { success: false, error: 'ไม่พบครูที่ต้องการลบ' };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

/**
 * @summary Imports teacher data from an array and adds them to the sheet.
 * @param {Array<object>} teachersData - An array of teacher objects.
 * @returns {object} A success or error object with the new list of teachers.
 */
function importTeachers(teachersData) {
  if (!Array.isArray(teachersData) || teachersData.length === 0) {
    return { success: false, error: 'ไม่มีข้อมูลที่จะนำเข้า หรือรูปแบบข้อมูลไม่ถูกต้อง' };
  }

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEACHER_SHEET_NAME);
    let importedCount = 0;
    
    teachersData.forEach(teacher => {
      const name = teacher['ชื่อ-สกุล'];
      const position = teacher['ตำแหน่ง'];

      if (name && position) {
        const newId = Utilities.getUuid();
        sheet.appendRow([newId, name, position]);
        importedCount++;
      }
    });

    if (importedCount === 0) {
      return { success: false, error: 'ไม่พบข้อมูลที่ถูกต้องในไฟล์ Excel โปรดตรวจสอบว่ามีคอลัมน์ "ชื่อ-สกุล" และ "ตำแหน่ง"' };
    }
    
    return { success: true, message: `นำเข้าข้อมูลครู ${importedCount} คนสำเร็จ`, teachers: getTeachers() };
  } catch (e) {
    return { success: false, error: 'เกิดข้อผิดพลาดระหว่างการนำเข้า: ' + e.message };
  }
}
