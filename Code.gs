var SHEET_NAME = "Nama Sheet yang akan jadi sumber data"; 
var WEB_APP_URL = "isi dengan link yang didapatkan dari Deploy"; 

function startApprovalProcess() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var data = sheet.getDataRange().getValues();
  
  // Looping mulai dari baris 2 (i=1)
  for (var i = 1; i < data.length; i++) {
    var rowNumber = i + 1;
    var emailPengaju = data[i][2]; // Kolom C
    var statusSentACC1 = data[i][13]; // Kolom M
    
    // HANYA proses jika ada email pengaju dan belum dikirim (Sent)
    if (emailPengaju && statusSentACC1 !== "Sent") {
      sendApprovalEmail(rowNumber, "ACC1");
      sheet.getRange(rowNumber, 14).setValue("Sent"); 
    }
  }
}

function getApproverEmail(level) {
  var ss = SpreadsheetApp.openById("isi dengan id dari database/google sheet"); // Sesuaikan ID jika perlu
  var sheet = ss.getSheetByName("nama sheet konfigurasi alamat email setiap ACC");
  
  var columnMap = {
    "ACC1": 4,
    "ACC2": 5,
    "ACC3": 6,
    "Executor": 7
  };
  return sheet.getRange(3, columnMap[level]).getValue();
}

function sendApprovalEmail(rowNumber, level) {
  // 1. CEK KEAMANAN: Jangan lanjut jika rowNumber tidak valid
  if (!rowNumber || rowNumber < 2) return; 

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rowData = sheet.getRange(rowNumber, 1, 1, 25).getValues()[0];
  var emailPengaju = rowData[2];
  if (!emailPengaju) return; 
  var approvalTarget = getApproverEmail(level);
  var attachmentUrl = rowData[9];
  var fileId = attachmentUrl.match(/[-\w]{25,}/); 
  var attachments = [];
  
  try {
    var fileId = attachmentUrl.match(/[-\w]{25,}/); 
    if (fileId) {
      var file = DriveApp.getFileById(fileId[0]);
      var mimeType = file.getMimeType();

      // Jika file adalah Google Docs/Sheets, paksa jadi PDF agar profesional
      if (mimeType === MimeType.GOOGLE_DOCS || mimeType === MimeType.GOOGLE_SHEETS) {
        attachments.push(file.getAs(MimeType.PDF).setName(file.getName() + ".pdf"));
      } else {
        // Jika sudah PDF/JPG/PNG, lampirkan apa adanya
        attachments.push(file.getBlob());
      }
    }
  } catch (e) {
    Logger.log("Gagal melampirkan file: " + e.toString());
  }
                     
  var tanggalInput = new Date(rowData[0]); 
  var tanggalFormatted = Utilities.formatDate(tanggalInput, Session.getScriptTimeZone(), "dd/MM/yyyy");
  var template = HtmlService.createTemplateFromFile('Template_CreditLimit');
  template.data = { 
    tanggal: tanggalFormatted,
    diajukan: rowData[1],
    metode: rowData[3],
    nomor : rowData[4],
    customer: rowData[5], 
    limit: rowData[6],
    amount: rowData[7],
    so: rowData[8], 
    lampiran: attachmentUrl,
    po: rowData[10],
    note: rowData[11], 
    amount: rowData[7],
    acc1: rowData[14],
    acc1_timestamp: rowData[15],
    acc2: rowData[17],
    acc2_timestamp: rowData[18],
    acc3: rowData[20],
    acc3_timestamp: rowData[21],
    exec: rowData[23],
    exec_timestamp: rowData[24],
    approveUrl: WEB_APP_URL + "?action=approve&level=" + level + "&row=" + rowNumber, 
    declineUrl: WEB_APP_URL + "?action=decline&level=" + level + "&row=" + rowNumber + "&email=" + emailPengaju 
  };
  
  MailApp.sendEmail({ 
    to: approvalTarget, 
    subject: "Approval Request Nomor " + rowData[4], 
    htmlBody: template.evaluate().getContent(),
    attachments: attachments 
  });
}

// 3. SATU-SATUNYA FUNGSI DOGET
function doGet(e) {
  if (!e || !e.parameter) return ContentService.createTextOutput("Error: Invalid Link.");

  var action = e.parameter.action;
  var row = parseInt(e.parameter.row);
  var level = e.parameter.level;
  
  if (isNaN(row)) return ContentService.createTextOutput("Error: Baris tidak valid.");

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  
  // 1. CEK STATUS SEBELUM PROSES
  var checkCol = (level === "acc1") ? 15 : (level === "acc2") ? 18 : (level === "acc3") ? 21 : (level === "Executor") ? 24 : 15;
  if (sheet.getRange(row, checkCol).getValue() === "APPROVED" || sheet.getRange(row, checkCol).getValue() === "DECLINED") {
    return ContentService.createTextOutput("Permintaan ini sudah pernah diproses sebelumnya.");
  }
  // 2. LOGIKA APPROVE
  if (action === "approve") {
    var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");

    if (level === "acc1") {
      sheet.getRange(row, 15).setValue("APPROVED");
      sheet.getRange(row, 16).setValue(formattedDate);
      sendApprovalEmail(row, "acc2");
      sheet.getRange(row, 17).setValue("Sent");
    } 
    else if (level === "acc2") {
      sheet.getRange(row, 18).setValue("APPROVED");
      sheet.getRange(row, 19).setValue(formattedDate);
      sendApprovalEmail(row, "acc3");
      sheet.getRange(row, 20).setValue("Sent");
    } 
    else if (level === "acc3") {
      sheet.getRange(row, 21).setValue("APPROVED");
      sheet.getRange(row, 22).setValue(formattedDate);
      sendApprovalEmail(row, "Executor");
      sheet.getRange(row, 23).setValue("Sent");
    }
    else if (level === "Executor") {
      sheet.getRange(row, 24).setValue("APPROVED");
      sheet.getRange(row, 25).setValue(formattedDate);
      
      // GENERATE PDF
      var pdfFile = generatePDF(row);
      var emailPengaju = sheet.getRange(row, 3).getValue(); // Kolom C
      
      MailApp.sendEmail({
        to: emailPengaju,
        subject: "APPROVAL Status",
        body: "Memo telah disetujui, berikut terlampir file PDF-nya.",
        attachments: [pdfFile.getAs(MimeType.PDF)]
      });
      return ContentService.createTextOutput("Approval Berhasil & PDF Terkirim.");
    }
    return ContentService.createTextOutput("Approval Berhasil untuk level: " + level);
  } 
  
  // 3. LOGIKA DECLINE
  else if (action === "decline") {
    var statusCol = (level === "acc1") ? 15 : (level === "acc2") ? 18 : (level === "acc3") ? 21 : (level === "Executor") ? 24 : 15;
    sheet.getRange(row, statusCol).setValue("DECLINED");
    sheet.getRange(row, statusCol + 1).setValue(new Date());
    var so = sheet.getRange(row, 8).getValue();
    
    MailApp.sendEmail({
      to: e.parameter.email, 
      subject: "Pengajuan ACC Credit Limit" + so + "Ditolak", 
      body: "Maaf, permintaan Anda ditolak oleh " + level + "."
    });
    return ContentService.createTextOutput("Permintaan berhasil ditolak oleh " + level);
  }
}

function onEdit(e) {
  var range = e.range;
  var sheet = range.getSheet();
  
  // Cek apakah yang diedit adalah kolom 24 (W) dan nilainya "APPROVED"
  if (sheet.getName() === "nama sheet database/sumber data" && range.getColumn() === 24 && range.getValue() === "APPROVED") {
    var row = range.getRow();
    
    // Panggil fungsi PDF Anda
    try {
      var pdfFile = generatePDF(row);
      var emailPengaju = sheet.getRange(row, 3).getValue();
      
      MailApp.sendEmail({
        to: emailPengaju,
        subject: "Approval Disetujui",
        body: "Memo telah disetujui, berikut terlampir file PDF-nya.",
        attachments: [pdfFile.getAs(MimeType.PDF)]
      });
    } catch (err) {
      Logger.log("Gagal buat PDF: " + err.toString());
    }
  }
}

function generatePDF(rowNumber) {
  if (!rowNumber || rowNumber < 2) return; 

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  var rowData = sheet.getRange(rowNumber, 1, 1, 25).getValues()[0];  
  // Jika sheet ini masih null, berarti ID Spreadsheet Anda salah
  if (!sheet) {
    throw new Error("Gagal mengakses sheet. Cek ID Spreadsheet Anda!");
  }
  Logger.log("Mencoba akses baris ke: " + rowNumber);
  
  // Baris 167 (tempat error Anda)
  var rowData = sheet.getRange(rowNumber, 1, 1, 25).getValues()[0];
  var nomorurut = rowData[4];
  var months = ["", "I", "II", "III", "IV", "V", "VI", "VII", "VIII", "IX", "X", "XI", "XII"];
  var month = months[new Date().getMonth() + 1];
  var year = new Date().getFullYear();
  
  // Gabungkan nomor urut dengan prefix Anda
  var nomorMemo = nomorurut + "/prefix/" + month + "/" + year;
  var templateId = "id dari format google docs";
  var folder = DriveApp.getFolderById("folder lokasi file");
  
  var copy = DriveApp.getFileById(templateId).makeCopy("Memo_" + nomorMemo, folder);
  var doc = DocumentApp.openById(copy.getId());
  var body = doc.getBody();
  var tanggalInput = new Date(rowData[0]); 
  var tanggalFormatted = Utilities.formatDate(tanggalInput, Session.getScriptTimeZone(), "dd/MM/yyyy");
  
  // Ganti blok replaceText Anda dengan ini:
  body.replaceText("{{nomor_memo}}", nomorMemo);
  body.replaceText("{{tanggal_ajuan}}", tanggalFormatted); // Kolom A: Timestamp sebagai pengganti nomor memo (atau sesuaikan jika punya kolom lain)
  body.replaceText("{{nama_pengaju}}", rowData[1]); // Kolom B: Request By (Diubah dari index 2 ke 1)
  body.replaceText("{{nama_pengaju}}", rowData[3]); // Kolom D: Request By (Diubah dari index 2 ke 1)
  body.replaceText("{{customer_name}}", rowData[5]); 
  body.replaceText("{{current_limit}}", rowData[6]);
  body.replaceText("{{total_transaksi}}", rowData[7]); // Kolom G: Transaction Amount (Sudah benar)
  body.replaceText("{{nomor_so}}", rowData[8]);    // Kolom H: SO Number (Sudah benar)
  body.replaceText("{{no_po}}", rowData[10]);       // Kolom J: PO Number (Sudah benar)
  body.replaceText("{{keterangan}}", rowData[11]);
  body.replaceText("{{acc1}}", rowData[14]);
  body.replaceText("{{acc1_timestamp}}", rowData[15]);
  body.replaceText("{{acc2}}", rowData[17]);
  body.replaceText("{{acc2_timestamp}}", rowData[18]);
  body.replaceText("{{acc3}}", rowData[20]);
  body.replaceText("{{acc3_timestamp}}", rowData[21]);
  body.replaceText("{{exec}}", rowData[23]);
  body.replaceText("{{exec_timestamp}}", rowData[24]);
  body.replaceText("{{tanggal_memo}}", Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy"));
  
  doc.saveAndClose();
  // 2. Beri jeda kecil agar proses penyimpanan selesai (opsional tapi disarankan)
  Utilities.sleep(1000); 
  
  var pdfBlob = copy.getAs("application/pdf");
  var pdfFile = folder.createFile(pdfBlob);
  copy.setTrashed(true);
  
  return pdfFile;
}
