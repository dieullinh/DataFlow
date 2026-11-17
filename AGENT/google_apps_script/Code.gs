/**
 * CÔNG TY CÔNG NGHỆ MINH AN - HỆ THỐNG CRM TỰ ĐỘNG HÓA VỚI AI AGENT
 * File: Code.gs (Google Apps Script)
 * 
 * Luồng hoạt động theo sơ đồ:
 * 1. Tiếp nhận Form/Email -> Ghi Sheets
 * 2. Gọi AI Agent bóc tách, đối soát bảng giá & tồn kho
 * 3. Nếu Chiết khấu > 10% -> Tự gửi Email duyệt cho Trưởng phòng
 * 4. Trưởng phòng bấm link Duyệt/Từ chối -> Cập nhật trạng thái
 * 5. Tự động nhân bản mẫu Google Docs, điền thông tin và gắn URL vào Sheets
 */

// Cấu hình ID và Hằng số
const CONFIG = {
  DOC_TEMPLATE_ID: "YOUR_GOOGLE_DOC_TEMPLATE_ID_HERE",
  QUOTE_FOLDER_ID: "YOUR_GOOGLE_DRIVE_FOLDER_ID_HERE",
  SALES_MANAGER_EMAIL: "truongphong.kd@minhan.com.vn",
  DISCOUNT_THRESHOLD_PERCENT: 10.0,
  SHEET_NAME_CRM: "CoHoiKinhDoanh",
  WEBAPP_URL: ScriptApp.getService().getUrl() // URL WebApp để nhận lệnh Duyệt 1-click từ email
};

/**
 * 1. Trigger khi có dữ liệu yêu cầu mới từ Google Forms / Web
 */
function onFormSubmit(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME_CRM);
  const row = sheet.getLastRow();
  
  // Đọc thông tin từ dòng mới
  const rawEmailContent = sheet.getRange(row, 3).getValue(); // Cột C: Nội dung yêu cầu thô
  const khachHang = sheet.getRange(row, 2).getValue();       // Cột B: Tên khách hàng
  
  // Gọi hàm phân tích & đối soát
  processInquiryWithAI(row, rawEmailContent, khachHang);
}

/**
 * 2. Tự động xử lý phân luồng Phê duyệt & Tạo báo giá khi NVKD hoặc AI cập nhật
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== CONFIG.SHEET_NAME_CRM) return;
  
  const range = e.range;
  const row = range.getRow();
  const col = range.getColumn();
  
  // Giả sử Cột F là Tỷ lệ chiết khấu, Cột J là Checkbox "Tạo báo giá"
  const COL_DISCOUNT = 6;
  const COL_STATUS = 7;
  const COL_CHECKBOX = 10;
  
  // Khi Chiết khấu được cập nhật
  if (col === COL_DISCOUNT && row > 1) {
    const discount = parseFloat(sheet.getRange(row, COL_DISCOUNT).getValue()) || 0;
    
    if (discount > CONFIG.DISCOUNT_THRESHOLD_PERCENT) {
      // VƯỢT THẨM QUYỀN -> Gửi Email cho Trưởng phòng
      sheet.getRange(row, COL_STATUS).setValue("Chờ Trưởng phòng duyệt");
      sendApprovalEmailToManager(row);
    } else {
      // ĐÚNG THẨM QUYỀN -> Tự duyệt
      sheet.getRange(row, COL_STATUS).setValue("Đã duyệt (Thẩm quyền NVKD)");
    }
  }
  
  // Khi Checkbox "Tạo báo giá" được tích chọn (TRUE)
  if (col === COL_CHECKBOX && range.getValue() === true) {
    const status = sheet.getRange(row, COL_STATUS).getValue();
    if (status.indexOf("Đã duyệt") !== -1) {
      generateQuotationDoc(row);
    } else {
      SpreadsheetApp.getUi().alert("Chưa được duyệt chiết khấu! Vui lòng chờ Trưởng phòng phê duyệt.");
      range.setValue(false); // Bỏ tích
    }
  }
}

/**
 * 3. Gửi Email Duyệt tương tác cho Trưởng phòng
 */
function sendApprovalEmailToManager(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME_CRM);
  const oppId = sheet.getRange(row, 1).getValue();
  const khachHang = sheet.getRange(row, 2).getValue();
  const nvkd = sheet.getRange(row, 4).getValue();
  const tongTien = sheet.getRange(row, 5).getValue();
  const chietKhau = sheet.getRange(row, 6).getValue();
  
  const approveUrl = CONFIG.WEBAPP_URL + "?action=approve&row=" + row;
  const rejectUrl = CONFIG.WEBAPP_URL + "?action=reject&row=" + row;
  
  const subject = `[TRÌNH DUYỆT BÁO GIÁ] Mã: ${oppId} - Khách hàng: ${khachHang} (CK: ${chietKhau}%)`;
  
  const htmlBody = `
    <h3>Kính gửi Trưởng phòng Kinh doanh,</h3>
    <p>NVKD <b>${nvkd}</b> vừa lập đề xuất báo giá vượt mức thẩm quyền tự quyết (> 10%):</p>
    <ul>
      <li><b>Mã cơ hội:</b> ${oppId}</li>
      <li><b>Khách hàng:</b> ${khachHang}</li>
      <li><b>Tổng giá trị niêm yết:</b> ${Number(tongTien).toLocaleString('vi-VN')} VNĐ</li>
      <li><b>Mức chiết khấu xin duyệt:</b> <span style="color:red; font-size:16px;"><b>${chietKhau}%</b></span></li>
    </ul>
    <p>Trưởng phòng vui lòng chọn hành động trực tiếp bên dưới:</p>
    <p>
      <a href="${approveUrl}" style="background-color:#28a745; color:white; padding:10px 20px; text-decoration:none; border-radius:4px; font-weight:bold;">ĐỒNG Ý DUYỆT</a>
      &nbsp;&nbsp;
      <a href="${rejectUrl}" style="background-color:#dc3545; color:white; padding:10px 20px; text-decoration:none; border-radius:4px; font-weight:bold;">TỪ CHỐI</a>
    </p>
    <br>
    <small>Hệ thống CRM Tự động hóa - Công ty Công Nghệ Minh An</small>
  `;
  
  MailApp.sendEmail({
    to: CONFIG.SALES_MANAGER_EMAIL,
    subject: subject,
    htmlBody: htmlBody
  });
}

/**
 * 4. Xử lý phản hồi Duyệt/Từ chối từ link Email của Trưởng phòng
 */
function doGet(e) {
  const action = e.parameter.action;
  const row = parseInt(e.parameter.row);
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME_CRM);
  
  const COL_STATUS = 7;
  const COL_CHECKBOX = 10;
  
  if (action === "approve") {
    sheet.getRange(row, COL_STATUS).setValue("Đã duyệt");
    sheet.getRange(row, COL_CHECKBOX).setValue(true); // Tự động tích tạo báo giá
    generateQuotationDoc(row); // Tự động tạo Docs luôn
    return ContentService.createTextOutput("ĐÃ PHÊ DUYỆT THÀNH CÔNG! Báo giá Google Docs đang được tự động khởi tạo.");
  } else if (action === "reject") {
    sheet.getRange(row, COL_STATUS).setValue("Hoàn thành (Từ chối)");
    sheet.getRange(row, COL_CHECKBOX).setValue(false);
    return ContentService.createTextOutput("ĐÃ TỪ CHỐI BÁO GIÁ. Quy trình đã kết thúc/hủy.");
  }
  
  return ContentService.createTextOutput("Yêu cầu không hợp lệ.");
}

/**
 * 5. Tự động nhân bản Google Docs, điền thông tin và dán URL vào Sheets
 */
function generateQuotationDoc(row) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.SHEET_NAME_CRM);
  const oppId = sheet.getRange(row, 1).getValue();
  const khachHang = sheet.getRange(row, 2).getValue();
  const nvkd = sheet.getRange(row, 4).getValue();
  const tongTien = sheet.getRange(row, 5).getValue();
  const chietKhau = sheet.getRange(row, 6).getValue();
  const thanhTien = tongTien * (1 - chietKhau / 100);
  const thoiGian = sheet.getRange(row, 8).getValue() || "Tiêu chuẩn (2-3 ngày)";
  
  // Nhân bản Template Doc
  const templateFile = DriveApp.getFileById(CONFIG.DOC_TEMPLATE_ID);
  const targetFolder = DriveApp.getFolderById(CONFIG.QUOTE_FOLDER_ID);
  const newFileName = `Báo Giá_${oppId}_${khachHang}`;
  const copyDoc = templateFile.makeCopy(newFileName, targetFolder);
  const doc = DocumentApp.openById(copyDoc.getId());
  const body = doc.getBody();
  
  // Thay thế biến
  body.replaceText("{{MA_BAO_GIA}}", oppId);
  body.replaceText("{{TEN_KHACH_HANG}}", khachHang);
  body.replaceText("{{NGUOI_LAP}}", nvkd);
  body.replaceText("{{TONG_NIEM_YET}}", Number(tongTien).toLocaleString('vi-VN') + " VNĐ");
  body.replaceText("{{TY_LE_CK}}", chietKhau + "%");
  body.replaceText("{{THANH_TIEN}}", Number(thanhTien).toLocaleString('vi-VN') + " VNĐ");
  body.replaceText("{{THOI_GIAN_TRIEN_KHAI}}", thoiGian);
  body.replaceText("{{NGAY_LAP}}", Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy"));
  
  doc.saveAndClose();
  
  // Dán URL vào Cột K (Cột 11)
  const docUrl = copyDoc.getUrl();
  sheet.getRange(row, 11).setValue(docUrl);
  
  return docUrl;
}

