/**
 * SISTEM PENDAFTARAN DRIVER GARJEK - BACKEND GOOGLE APPS SCRIPT
 * File: Code.gs
 * Description: Backend pemroses data pendaftaran, penyimpanan berkas ke Google Drive, 
 *              dan manajemen database di Google Sheets secara otomatis.
 */

// Nama Sheet Database
const SHEET_DRIVER_NAME = "Data_Driver";
const SHEET_LOGS_NAME = "System_Logs";
const DRIVE_FOLDER_NAME = "GARJEK_BERKAS_DRIVER_UPLOADS";

/**
 * FUNGSI SETUP DATABASE OTOMATIS
 * Jalankan fungsi ini 1 kali dari editor Apps Script untuk membuat Sheet & Folder Drive secara otomatis.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Setup Sheet Data_Driver
  let driverSheet = ss.getSheetByName(SHEET_DRIVER_NAME);
  if (!driverSheet) {
    driverSheet = ss.insertSheet(SHEET_DRIVER_NAME);
  }
  
  // Header Kolom Database Driver
  const headers = [
    "ID Registrasi",
    "Tanggal Daftar",
    "NIK",
    "Nama Lengkap",
    "No Telepon/WA",
    "Email",
    "Wilayah Domisili",
    "Alamat Lengkap",
    "Layanan",
    "Kendaraan & Plat",
    "Link Foto KTP",
    "Link Foto SIM",
    "Link Foto STNK",
    "Link Foto Kendaraan",
    "Link Pas Foto",
    "Link Struk DANA 150k",
    "Status Verifikasi",
    "Alasan Penolakan"
  ];
  
  driverSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  
  // Formatting Header
  const headerRange = driverSheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground("#03045e")
             .setFontColor("#ffffff")
             .setFontWeight("bold")
             .setHorizontalAlignment("center")
             .setVerticalAlignment("middle");
  driverSheet.setRowHeight(1, 40);
  driverSheet.setFrozenRows(1);

  // Auto resize kolom
  for (let i = 1; i <= headers.length; i++) {
    driverSheet.setColumnWidth(i, 160);
  }

  // 2. Setup Folder Google Drive untuk Berkas
  getOrCreateUploadFolder();

  // 3. Setup Sheet System_Logs
  let logSheet = ss.getSheetByName(SHEET_LOGS_NAME);
  if (!logSheet) {
    logSheet = ss.insertSheet(SHEET_LOGS_NAME);
    logSheet.getRange(1, 1, 1, 3).setValues([["Timestamp", "Aktivitas", "Detail"]]);
    logSheet.getRange(1, 1, 1, 3).setBackground("#10b981").setFontColor("#ffffff").setFontWeight("bold");
  }

  Logger.log("✅ Setup Database Google Sheets & Folder Google Drive GARJEK Berhasil!");
  return "Setup Berhasil! Sheet & Folder Google Drive Siap Digunakan.";
}

/**
 * Membuka atau membuat folder Google Drive penyimpanan berkas pendaftaran
 */
function getOrCreateUploadFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  if (folders.hasNext()) {
    return folders.next();
  } else {
    const folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return folder;
  }
}

/**
 * Menyimpan berkas gambar Base64 ke Google Drive dan mengembalikan URL Akses Publik
 */
function saveBase64ToDrive(base64Data, fileName, mimeType) {
  try {
    if (!base64Data || !base64Data.includes("base64,")) {
      return base64Data || "Tidak ada berkas";
    }

    const folder = getOrCreateUploadFolder();
    const cleanBase64 = base64Data.split("base64,")[1];
    const decodedBlob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), mimeType || "image/jpeg", fileName);
    
    const file = folder.createFile(decodedBlob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return file.getUrl();
  } catch (err) {
    Logger.log("Error Upload Drive: " + err.message);
    return "Gagal Upload: " + err.message;
  }
}

/**
 * HANDLER GET REQUEST (Membaca Data / Cek NIK / Dashboard Admin)
 */
function doGet(e) {
  const response = { success: false, data: null, message: "" };
  
  try {
    const action = e.parameter.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DRIVER_NAME);

    if (!sheet) {
      response.message = "Sheet 'Data_Driver' belum dibuat. Jalankan setupDatabase() terlebih dahulu.";
      return createJsonResponse(response);
    }

    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const dataRows = rows.slice(1);

    // ACTION 1: Cek Status Pendaftaran berdasarkan NIK
    if (action === "checkStatus") {
      const nikQuery = (e.parameter.nik || "").trim();
      
      const foundRow = dataRows.find(row => String(row[2]).trim() === nikQuery);
      
      if (foundRow) {
        response.success = true;
        response.data = {
          regId: foundRow[0],
          timestamp: foundRow[1],
          nik: foundRow[2],
          nama: foundRow[3],
          noHp: foundRow[4],
          email: foundRow[5],
          wilayah: foundRow[6],
          alamat: foundRow[7],
          layanan: foundRow[8],
          kendaraan: foundRow[9],
          fileKtp: foundRow[10],
          fileSim: foundRow[11],
          fileStnk: foundRow[12],
          fileKendaraan: foundRow[13],
          filePasFoto: foundRow[14],
          fileBayar: foundRow[15],
          status: foundRow[16],
          alasanPenolakan: foundRow[17] || ""
        };
        response.message = "Data pendaftar ditemukan.";
      } else {
        response.success = false;
        response.message = "NIK tidak ditemukan dalam sistem GARJEK.";
      }
    } 
    // ACTION 2: Get Semua Data Driver untuk Dashboard Admin
    else if (action === "getAllDrivers") {
      const driverList = dataRows.map(row => ({
        regId: row[0],
        timestamp: row[1],
        nik: row[2],
        nama: row[3],
        noHp: row[4],
        email: row[5],
        wilayah: row[6],
        alamat: row[7],
        layanan: row[8],
        kendaraan: row[9],
        fileKtp: row[10],
        fileSim: row[11],
        fileStnk: row[12],
        fileKendaraan: row[13],
        filePasFoto: row[14],
        fileBayar: row[15],
        status: row[16],
        alasanPenolakan: row[17] || ""
      }));

      response.success = true;
      response.data = driverList;
      response.message = "Berhasil memuat seluruh data driver.";
    } 
    // ACTION 3: Run Auto Setup melalui URL GET
    else if (action === "setup") {
      const setupMsg = setupDatabase();
      response.success = true;
      response.message = setupMsg;
    } 
    else {
      response.message = "Action tidak dikenali.";
    }

  } catch (err) {
    response.success = false;
    response.message = "Server Error: " + err.message;
  }

  return createJsonResponse(response);
}

/**
 * HANDLER POST REQUEST (Pendaftaran Baru & Verifikasi Admin)
 */
function doPost(e) {
  const response = { success: false, message: "", data: null };

  try {
    let postData;
    if (e.postData && e.postData.contents) {
      postData = JSON.parse(e.postData.contents);
    } else {
      postData = e.parameter;
    }

    const action = postData.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_DRIVER_NAME);

    if (!sheet) {
      response.message = "Database belum siap. Jalankan setupDatabase() terlebih dahulu.";
      return createJsonResponse(response);
    }

    // ACTION 1: Pendaftaran Driver Baru
    if (action === "registerDriver") {
      const nik = String(postData.nik).trim();
      
      // Cek Duplikasi NIK
      const existingRows = sheet.getDataRange().getValues();
      for (let i = 1; i < existingRows.length; i++) {
        if (String(existingRows[i][2]).trim() === nik) {
          response.message = "NIK " + nik + " sudah terdaftar di sistem!";
          return createJsonResponse(response);
        }
      }

      // Format Waktu
      const now = new Date();
      const timestamp = Utilities.formatDate(now, "GMT+7", "yyyy-MM-dd HH:mm:ss");
      const regId = "GARJEK-" + Math.floor(100000 + Math.random() * 900000);

      // Upload Gambar ke Google Drive
      const fileKtpUrl = saveBase64ToDrive(postData.fileKtp, `KTP_${nik}_${postData.nama}.jpg`, "image/jpeg");
      const fileSimUrl = saveBase64ToDrive(postData.fileSim, `SIM_${nik}_${postData.nama}.jpg`, "image/jpeg");
      const fileStnkUrl = saveBase64ToDrive(postData.fileStnk, `STNK_${nik}_${postData.nama}.jpg`, "image/jpeg");
      const fileKendaraanUrl = saveBase64ToDrive(postData.fileKendaraan, `KENDARAAN_${nik}_${postData.nama}.jpg`, "image/jpeg");
      const filePasFotoUrl = saveBase64ToDrive(postData.filePasFoto, `PASFOTO_${nik}_${postData.nama}.jpg`, "image/jpeg");
      const fileBayarUrl = saveBase64ToDrive(postData.fileBayar, `STRUK_DANA_${nik}_${postData.nama}.jpg`, "image/jpeg");

      // Baris Data Baru
      const newRow = [
        regId,
        timestamp,
        nik,
        postData.nama.toUpperCase(),
        postData.noHp,
        postData.email,
        postData.wilayah.toUpperCase(),
        postData.alamat.toUpperCase(),
        postData.layanan,
        postData.kendaraan.toUpperCase(),
        fileKtpUrl,
        fileSimUrl,
        fileStnkUrl,
        fileKendaraanUrl,
        filePasFotoUrl,
        fileBayarUrl,
        "Pending", // Status awal
        ""         // Alasan penolakan kosong
      ];

      sheet.appendRow(newRow);

      response.success = true;
      response.message = "Pendaftaran Driver GARJEK Berhasil Disimpan!";
      response.data = { regId: regId, nik: nik, status: "Pending" };
    }

    // ACTION 2: Update Status Verifikasi (Admin)
    else if (action === "updateDriverStatus") {
      const regId = postData.regId;
      const newStatus = postData.status; // Verified / Rejected / Pending
      const reason = postData.alasanPenolakan || "";

      const data = sheet.getDataRange().getValues();
      let rowFound = -1;

      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === regId) {
          rowFound = i + 1; // 1-indexed row number
          break;
        }
      }

      if (rowFound !== -1) {
        sheet.getRange(rowFound, 17).setValue(newStatus); // Kolom Q: Status
        sheet.getRange(rowFound, 18).setValue(reason);    // Kolom R: Alasan Penolakan

        response.success = true;
        response.message = `Status driver ${regId} berhasil diubah menjadi ${newStatus}.`;
      } else {
        response.message = "ID Registrasi tidak ditemukan.";
      }
    }

    else {
      response.message = "Action POST tidak valid.";
    }

  } catch (err) {
    response.success = false;
    response.message = "Error Server POST: " + err.message;
  }

  return createJsonResponse(response);
}

/**
 * Formatter Response JSON dengan header CORS agar bisa diakses dari Web Front-end
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}