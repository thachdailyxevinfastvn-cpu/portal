/**
 * @OnlyCurrentDoc
 */

/**
 * Tạo menu tùy chỉnh khi người dùng mở bảng tính.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tùy Chỉnh Automation')
    .addItem('1. Tạo sheet con từ Template (v4)', 'createSheetsFromTemplate')
    .addToUi();
}

/**
 * Tạo các sheet con từ sheet "Template".
 * Phiên bản này đã sửa lỗi regex để nhận dạng được placeholder chứa dấu "/".
 */
function createSheetsFromTemplate() {
  const ui = SpreadsheetApp.getUi();
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const templateSheet = ss.getSheetByName('Template');

    if (!templateSheet) {
      throw new Error('Sheet "Template" không tồn tại. Vui lòng tạo sheet này.');
    }

    const data = templateSheet.getDataRange().getValues();

    // Duyệt qua từng dòng trong sheet "Template", bắt đầu từ dòng 2
    for (let i = 1; i < data.length; i++) {
      const rowData = data[i];
      const newSheetName = (rowData[1] || '').toString().trim(); // Cột B

      if (!newSheetName) {
        continue; // Bỏ qua dòng trống
      }
      
      const allPlaceholders = new Set();

      // Quét các cột từ C (index 2) trở đi trong dòng hiện tại
      for (let c = 2; c < rowData.length; c++) {
        const cellValue = (rowData[c] || '').toString().trim();
        
        if (cellValue.includes('docs.google.com/document')) {
          const docId = extractFileId(cellValue);
          if (docId) {
            try {
              const doc = DocumentApp.openById(docId);
              const placeholdersInDoc = extractPlaceholders(doc);
              
              if (placeholdersInDoc.length > 0) {
                 placeholdersInDoc.forEach(ph => allPlaceholders.add(ph));
              }
            } catch (e) {
              Logger.log(`LỖI: Không thể mở hoặc xử lý file Doc ID: ${docId}. ${e.message}`);
            }
          }
        }
      }

      // Sau khi đã quét hết các link trong dòng, tiến hành tạo sheet
      const existingSheet = ss.getSheetByName(newSheetName);
      if (existingSheet) {
        ss.deleteSheet(existingSheet);
      }
      const newSheet = ss.insertSheet(newSheetName);

      const finalHeaders = ['stt', ...Array.from(allPlaceholders)];
      const subtitles = finalHeaders.map(h => (h === 'stt' ? 'Số thứ tự' : `Nhập ${h}`));

      if (finalHeaders.length > 1) {
        newSheet.getRange(1, 1, 1, finalHeaders.length).setValues([finalHeaders]);
        newSheet.getRange(2, 1, 1, subtitles.length).setValues([subtitles]);
      } else {
        const defaultHeaders = ['stt', 'TEN_KH', 'NGAYKY_HDMB', 'SO_TIEN'];
        newSheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
      }

      newSheet.setFrozenRows(2);
      newSheet.getRange("A:A").setNumberFormat('@');
    }

    ui.alert('✅ Hoàn tất! Đã cập nhật các sheet với logic mới nhất.');
  } catch (err) {
    Logger.log(`LỖI NGHIÊM TRỌNG: ${err.message}`);
    ui.alert(`❌ Đã xảy ra lỗi nghiêm trọng: ${err.message}`);
  }
}

/**
 * Trích xuất ID file từ URL Google Docs.
 * @param {string} url
 * @return {string|null}
 */
function extractFileId(url) {
  const match = url.match(/\/d\/([a-zA-Z0-9\-_]+)/);
  return match ? match[1] : null;
}

/**
 * Trích xuất tất cả các placeholder <<...>> từ một file Google Doc.
 * @param {GoogleAppsScript.Document.Document} doc
 * @return {string[]} Mảng các placeholder.
 */
function extractPlaceholders(doc) {
  const headerText = doc.getHeader() ? doc.getHeader().getText() : '';
  const footerText = doc.getFooter() ? doc.getFooter().getText() : '';
  const bodyText = doc.getBody().getText();
  const fullText = `${headerText} ${bodyText} ${footerText}`;

  const placeholders = new Set();
  
  // *** ĐÂY LÀ DÒNG ĐÃ ĐƯỢC SỬA ***
  // Đã thêm ký tự "/" vào trong bộ ký tự cho phép [A-Za-z0-9_/]
  const regex = /<<\s*([A-Za-z0-9_/]+)\s*>>/g;
  
  let match;
  while ((match = regex.exec(fullText)) !== null) {
    placeholders.add(match[1]);
  }
  return Array.from(placeholders);
}