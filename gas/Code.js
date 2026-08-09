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

    if (action === 'validate') {
      const validation = validateSheets(ss);
      return createJsonResponse({ status: 'success', validation });
    }

    // 2. Pendaftaran Wajah Baru
    if (action === 'register') {
      if (!data.id || !data.name || !Array.isArray(data.descriptors) || data.descriptors.length === 0) {
        return createJsonResponse({ status: 'error', message: 'Data registrasi tidak lengkap atau tidak valid.' });
      }

      const dataSheet = getOrCreateSheet(ss, 'Data_Pegawai', ['Timestamp', 'User_ID', 'Nama', 'Descriptors']);
      const lastRow = dataSheet.getLastRow();
      const existingUsers = lastRow > 1
        ? dataSheet.getRange(2, 2, lastRow - 1, 1).getValues()
        : [];

      if (existingUsers.some(row => String(row[0]).trim() === String(data.id).trim())) {
        return createJsonResponse({ status: 'error', message: 'User ID sudah terdaftar. Gunakan ID lain atau periksa database.' });
      }

      const timestamp = new Date();
      dataSheet.appendRow([
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

function validateSheets(ss) {
  const requiredSheets = [
    { name: 'Admin_Credentials', headers: ['Username', 'Password'] },
    { name: 'Data_Pegawai', headers: ['Timestamp', 'User_ID', 'Nama', 'Descriptors'] },
    { name: 'Log_Absensi', headers: ['Timestamp', 'User_ID', 'Nama', 'Koordinat_GPS', 'Jarak_Radius', 'Status'] }
  ];

  const validation = {};

  requiredSheets.forEach(sheetInfo => {
    try {
      const sheet = ss.getSheetByName(sheetInfo.name);
      if (!sheet) {
        validation[sheetInfo.name] = { ok: false, reason: 'Sheet tidak ditemukan' };
        return;
      }

      const headerValues = sheet.getRange(1, 1, 1, sheetInfo.headers.length).getValues()[0];
      const missingHeader = sheetInfo.headers.find((expected, index) => String(headerValues[index] || '').trim() !== expected);
      if (missingHeader) {
        validation[sheetInfo.name] = { ok: false, reason: 'Header sheet tidak sesuai atau hilang' };
        return;
      }

      if (sheetInfo.name === 'Data_Pegawai') {
        const dataRowCount = Math.max(sheet.getLastRow() - 1, 0);
        const values = dataRowCount > 0
          ? sheet.getRange(2, 1, dataRowCount, sheetInfo.headers.length).getValues()
          : [];
        const userIds = {};
        for (let i = 0; i < values.length; i++) {
          const row = values[i];
          const rowId = String(row[1] || '').trim();
          const rowDescriptors = row[3];

          if (!rowId) {
            validation[sheetInfo.name] = { ok: false, reason: `Baris ${i + 2} memiliki User_ID kosong` };
            return;
          }

          if (userIds[rowId]) {
            validation[sheetInfo.name] = { ok: false, reason: `Duplikasi User_ID di baris ${userIds[rowId]} & ${i + 2}` };
            return;
          }
          userIds[rowId] = i + 2;

          try {
            const parsed = JSON.parse(String(rowDescriptors || '[]'));
            if (!Array.isArray(parsed) || parsed.length === 0) {
              validation[sheetInfo.name] = { ok: false, reason: `Descriptors tidak valid pada baris ${i + 2}` };
              return;
            }
          } catch (ex) {
            validation[sheetInfo.name] = { ok: false, reason: `Descriptors JSON gagal di-parse pada baris ${i + 2}` };
            return;
          }
        }
      }

      validation[sheetInfo.name] = { ok: true, reason: 'OK' };
    } catch (err) {
      validation[sheetInfo.name] = { ok: false, reason: `Kesalahan validasi: ${err.message}` };
    }
  });

  return validation;
}

// Fungsi pembantu response JSON
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}