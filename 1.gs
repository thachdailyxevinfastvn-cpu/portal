// === CẤU HÌNH ===
const TARGET_FOLDER_ID = '1YP4xBxsMx2YEZrzFojCQJsFPvZ6X8_aB'; 
const TEMPLATE_SHEET_NAME = 'Template';
const LOG_SHEET_NAME = 'Log';

// === HÀM CHÍNH (JSONP ENTRYPOINT) ===
function doGet(e) {
  const action = e.parameter.action;
  const callback = e.parameter.callback;
  let response;

  try {
    switch(action) {
      case 'getGroups':
        response = getDocumentGroups();
        break;
      case 'getPlaceholders':
        const groupName = e.parameter.group;
        response = getPlaceholdersForGroup(groupName);
        break;
      case 'processFormCreation':
        const payload = JSON.parse(e.parameter.payload);
        response = processFormCreation(payload);
        break;
      default:
        throw new Error('Hành động không hợp lệ');
    }
  } catch (err) {
    response = { status: 'error', message: err.message };
  }

  return ContentService.createTextOutput(`${callback}(${JSON.stringify(response)})`)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

// === CÁC HÀM LOGIC ===

function getDocumentGroups() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEMPLATE_SHEET_NAME);
  if (!sheet) return { status: 'error', message: `Không tìm thấy trang tính '${TEMPLATE_SHEET_NAME}'` };
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { status: 'success', data: [] };

  const range = sheet.getRange(2, 2, lastRow - 1, 1);
  const values = range.getValues().flat().filter(String);
  
  // SỬA LỖI: Chỉ trả về các giá trị duy nhất, không lặp lại
  const uniqueValues = [...new Set(values)];
  return { status: 'success', data: uniqueValues };
}

function getPlaceholdersForGroup(groupName) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(groupName);
    if (!sheet) {
      return { status: 'error', message: `Không tìm thấy trang tính cho nhóm: '${groupName}'` };
    }
    if (sheet.getLastRow() < 2) {
      return { status: 'success', data: [] };
    }

    const lastColumn = sheet.getLastColumn();
    if (lastColumn < 1) return { status: 'success', data: [] }; 

    const names = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
    const labels = sheet.getRange(2, 1, 1, lastColumn).getValues()[0];

    const placeholders = [];
    for (let i = 0; i < names.length; i++) {
      if (names[i] && labels[i]) {
        placeholders.push({
          name: names[i].toString().trim(),
          label: labels[i].toString().trim()
        });
      }
    }
    return { status: 'success', data: placeholders };
  } catch (e) {
    return { status: 'error', message: `Không tìm thấy trang tính cho nhóm '${groupName}': ${e.message}` };
  }
}

/**
 * Hàm chính để xử lý việc tạo văn bản.
 * SỬA LỖI CUỐI CÙNG: Thay thế placeholder theo thứ tự ưu tiên để khắc phục lỗi _TEXT.
 */
function processFormCreation(payload) {
  const { groupName, placeholders, userName } = payload;
  try {
    const today = new Date();
    
    placeholders['<<TODAY>>'] = Utilities.formatDate(today, "GMT+7", "dd/MM/yyyy");

    for (const key in placeholders) {
      if (key.toUpperCase().includes('NGAY') && /^\d{4}-\d{2}-\d{2}$/.test(placeholders[key])) {
        const [year, month, day] = placeholders[key].split('-');
        placeholders[key] = `${day}/${month}/${year}`;
      }
    }

    const ngayKyHdValue = placeholders['<<NGAY_KY_HD>>'];
    if (ngayKyHdValue) {
      placeholders['<<NGAY_KY_HD_TEXT>>'] = convertDateToTextGS(ngayKyHdValue);
    }
    
    const templateSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(TEMPLATE_SHEET_NAME);
    if (!templateSheet) throw new Error(`Không tìm thấy trang tính '${TEMPLATE_SHEET_NAME}'`);
    const targetFolder = DriveApp.getFolderById(TARGET_FOLDER_ID);
    const data = templateSheet.getDataRange().getValues();
    const groupRow = data.find(row => row[1] === groupName);
    if (!groupRow) throw new Error(`Không tìm thấy nhóm '${groupName}' trong trang Template.`);
    const templateLinks = groupRow.slice(2).filter(String);
    if (templateLinks.length === 0) throw new Error(`Nhóm '${groupName}' không có link template nào.`);
    
    const generatedFiles = []; 
    templateLinks.forEach(link => {
      const templateId = extractIdFromUrl(link);
      const templateFile = DriveApp.getFileById(templateId);
      const templateName = templateFile.getName();
      const customerName = placeholders['<<TEN_KH>>'] || 'KhongCoTen';
      const fileCreationDate = Utilities.formatDate(today, "GMT+7", "dd-MM-yyyy"); 
      const newFileName = `${templateName} - ${customerName} - ${fileCreationDate}`;
      const newFile = templateFile.makeCopy(newFileName, targetFolder);
      const doc = DocumentApp.openById(newFile.getId());
      const body = doc.getBody();
      
      // ===== BẮT ĐẦU PHẦN SỬA LỖI CUỐI CÙNG =====
      // Lấy ra tất cả các key (placeholder) từ dữ liệu
      const allKeys = Object.keys(placeholders);
      
      // Sắp xếp các key theo độ dài từ DÀI NHẤT đến NGẮN NHẤT
      allKeys.sort((a, b) => b.length - a.length);
      
      // Lặp qua mảng key đã được sắp xếp và thực hiện thay thế
      // Việc này đảm bảo <<NGAY_KY_HD_TEXT>> luôn được thay thế trước <<NGAY_KY_HD>>
      for (const key of allKeys) {
        const value = placeholders[key] || '';
        body.replaceText(key, value);
      }
      // ===== KẾT THÚC PHẦN SỬA LỖI CUỐI CÙNG =====

      doc.saveAndClose();
      newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);
      generatedFiles.push({ name: templateName, url: newFile.getUrl() });
    });
    
    writeDataToGroupSheet(groupName, placeholders);
    logResult(userName, 'Thành công', generatedFiles);
    return { status: 'success', files: generatedFiles }; 
  } catch (e) {
    logResult(userName, `Thất bại: ${e.message}`, []);
    return { status: 'error', message: e.message };
  }
}




function writeDataToGroupSheet(groupName, data) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(groupName);
    if (!sheet) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const newRow = headers.map(header => data[header] || '');
    sheet.appendRow(newRow);
  } catch(e) { console.error(`Lỗi khi ghi dữ liệu vào sheet '${groupName}': ${e.message}`); }
}

function logResult(userName, status, filesArray) {
  try {
    const logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(LOG_SHEET_NAME);
    if (logSheet) {
      const timestamp = new Date();
      const links = filesArray.map(file => file.url);
      const rowData = [timestamp, userName, status, ...links];
      logSheet.appendRow(rowData);
    }
  } catch (e) { console.error(`Không thể ghi log: ${e.message}`); }
}

function extractIdFromUrl(url) {
  const match = url.match(/[-\w]{25,}/);
  if (match) return match[0];
  throw new Error(`URL không hợp lệ: ${url}`);
}

function convertDateToTextGS(dateString_ddMMyyyy) {
  if (!dateString_ddMMyyyy) return '';
  const parts = dateString_ddMMyyyy.split('/');
  if (parts.length !== 3) return dateString_ddMMyyyy; // Trả lại giá trị cũ nếu không đúng định dạng
  const [day, month, year] = parts;
  return `ngày ${day} tháng ${month} năm ${year}`;
}