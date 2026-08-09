function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // 1. Verifikasi Login Admin
    if (action === 'admin_login') {
      const sheet = getOrCreateSheet(ss, 'Admin_Credentials', ['Username', 'Password']);
      const records = sheet.getDataRange().getValues();
      let isValid = false;

      // Cek apakah data username & password cocok
      for (let i = 1; i < records.length; i++) {
        if (String(records[i][0]).trim() === String(data.username).trim() && 
            String(records[i][1]).trim() === String(data.password).trim()) {
          isValid = true;
          break;
        }
      }

      if (isValid) {
        return createJsonResponse({ status: 'success', message: 'Login berhasil!' });
      } else {
        return createJsonResponse({ status: 'error', message: 'Username atau Password salah!' });
      }
    }

    // 2. Pendaftaran Wajah Baru
    if (action === 'register') {
      const sheet = getOrCreateSheet(ss, 'Data_Pegawai', ['Timestamp', 'User_ID', 'Nama', 'Descriptors']);
      const timestamp = new Date();
      
      sheet.appendRow([
        timestamp, 
        data.id, 
        data.name, 
        JSON.stringify(data.descriptors)
      ]);
      
      return createJsonResponse({ status: 'success', message: 'Data pegawai berhasil tersimpan!' });
    }

    // 3. Pencatatan Absensi (Termasuk Kolom Radius/Jarak)
    if (action === 'absen') {
      const sheet = getOrCreateSheet(ss, 'Log_Absensi', ['Timestamp', 'User_ID', 'Nama', 'Koordinat_GPS', 'Jarak_Radius', 'Status']);
      const timestamp = new Date();
      
      sheet.appendRow([
        timestamp, 
        data.id, 
        data.name, 
        data.gps || "Tidak Terdeteksi", 
        data.radius || "0 meter",  // <--- Menyimpan data radius ke Google Sheets
        data.status || "Hadir"
      ]);
      
      return createJsonResponse({ status: 'success', message: 'Absensi berhasil dicatat!' });
    }

    return createJsonResponse({ status: 'error', message: 'Action tidak dikenal!' });

  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.message });
  }
}

function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getOrCreateSheet(ss, 'Data_Pegawai', ['Timestamp', 'User_ID', 'Nama', 'Descriptors']);
    const data = sheet.getDataRange().getValues();
    let users = [];

    // Baca data mulai dari baris ke-2 (mengabaikan header)
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && data[i][3]) { // Pastikan ID dan Descriptors tidak kosong
        users.push({
          id: data[i][1],
          name: data[i][2],
          descriptors: JSON.parse(data[i][3])
        });
      }
    }

    return createJsonResponse(users);
  } catch (error) {
    return createJsonResponse([]);
  }
}

// ==========================================
// FUNGSI BANTUAN (HELPER FUNCTIONS)
// ==========================================

// Fungsi untuk memastikan Sheet ada, jika belum ada akan dibuatkan beserta Headernya
function getOrCreateSheet(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e2e8f0");
  }
  return sheet;
}

// Fungsi pembantu response JSON
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}