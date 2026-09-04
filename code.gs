/**
 * BACKEND GOOGLE APPS SCRIPT - SISTEM ARSIP DIGITAL SISWA
 * Terintegrasi Google Spreadsheet & Google Drive
 * SMA Negeri 1 Kutowinangun
 */

// Nama-nama Tab Lembar Kerja
const SHEET_STUDENTS = "STUDENTS_DATA";
const SHEET_VERIF = "VERIFIKASI_STATUS";
const SHEET_ADMINS = "ADMIN_USERS";

// Nama Folder Utama Penyimpanan Berkas di Google Drive
const DRIVE_FOLDER_NAME = "ARSIP_BERKAS_SISWA_DRIVE";

/**
 * Penanganan permintaan GET (Pencarian arsip oleh siswa & pengambilan seluruh data master saat reload).
 */
function doGet(e) {
  try {
    const action = e && e.parameter ? e.parameter.action : '';

    // Ambil seluruh data dokumen & status verifikasi saat reload / inisialisasi aplikasi
    if (action === 'getAllData') {
      return handleGetAllData();
    }

    // Pencarian berkas berdasarkan 10 digit NISN
    if (action === 'searchNISN') {
      const searchNISN = e.parameter.nisn;
      return handleSearchNISN(searchNISN);
    }

    return createJsonResponse({
      status: 'error',
      message: 'Parameter aksi GET tidak valid. Gunakan action=getAllData atau action=searchNISN.'
    });

  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  }
}

/**
 * Mengambil seluruh data dokumen dari tab STUDENTS_DATA dan status verifikasi dari VERIFIKASI_STATUS.
 */
function handleGetAllData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
  let sheetVerif = ss.getSheetByName(SHEET_VERIF);

  if (!sheetStudents || !sheetVerif) {
    setupDatabase();
    sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    sheetVerif = ss.getSheetByName(SHEET_VERIF);
  }

  const studentRows = sheetStudents.getDataRange().getValues();
  const documents = [];
  
  // Baca baris data dokumen (mulai baris kedua setelah header)
  for (let i = 1; i < studentRows.length; i++) {
    const row = studentRows[i];
    if (row[0] && row[1]) {
      const rawNisn = String(row[1]).trim().replace(/^'/, '');
      const rawFileUrl = String(row[6]).trim();
      const rawFileId = String(row[5]).trim();
      
      documents.push({
        id: String(row[0]),
        nisn: rawNisn,
        nama_siswa: String(row[2]),
        kelas: String(row[3]),
        jenis_dokumen: String(row[4]),
        file_id: rawFileId,
        file_url: rawFileUrl || (rawFileId ? ('https://lh3.googleusercontent.com/d/' + rawFileId) : ''),
        file_name: String(row[7] || ''),
        uploaded_at: row[8] ? new Date(row[8]).toISOString() : new Date().toISOString(),
        uploaded_by: String(row[9] || 'admin_tu')
      });
    }
  }

  const verifRows = sheetVerif.getDataRange().getValues();
  const verifications = {};
  for (let j = 1; j < verifRows.length; j++) {
    const vRow = verifRows[j];
    if (vRow[0]) {
      const vNisn = String(vRow[0]).trim().replace(/^'/, '');
      verifications[vNisn] = {
        status: String(vRow[1] || 'PENDING'),
        note: String(vRow[2] || ''),
        confirmed_at: vRow[3] ? new Date(vRow[3]).toISOString() : '',
        validator_role: String(vRow[4] || 'siswa')
      };
    }
  }

  return createJsonResponse({
    status: 'success',
    documents: documents,
    verifications: verifications
  });
}

/**
 * Mencari berkas siswa berdasarkan NISN dari Google Spreadsheet.
 */
function handleSearchNISN(searchNISN) {
  if (!searchNISN) {
    return createJsonResponse({ status: 'error', message: 'NISN tidak boleh kosong.' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
  let sheetVerif = ss.getSheetByName(SHEET_VERIF);

  if (!sheetStudents || !sheetVerif) {
    setupDatabase();
    sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    sheetVerif = ss.getSheetByName(SHEET_VERIF);
  }

  const cleanQuery = String(searchNISN).trim().replace(/^0+/, '') || '0';
  const rawQuery = String(searchNISN).trim();
  const rows = sheetStudents.getDataRange().getValues();
  const matchedDocs = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const cellNISN = String(row[1] || '').trim().replace(/^'/, '');
    const cleanCellNISN = cellNISN.replace(/^0+/, '') || '0';

    if (cleanCellNISN === cleanQuery || cellNISN === rawQuery) {
      const fileId = String(row[5] || '').trim();
      const fileUrl = String(row[6] || '').trim();
      matchedDocs.push({
        id: String(row[0]),
        nisn: cellNISN,
        nama_siswa: String(row[2]),
        kelas: String(row[3]),
        jenis_dokumen: String(row[4]),
        file_id: fileId,
        file_url: fileUrl || (fileId ? `https://lh3.googleusercontent.com/d/${fileId}` : ''),
        file_name: String(row[7] || ''),
        uploaded_at: row[8] ? new Date(row[8]).toISOString() : '',
        uploaded_by: String(row[9] || 'admin')
      });
    }
  }

  // Ambil status verifikasi siswa
  let verificationInfo = { status: 'PENDING', note: '', confirmed_at: '', validator_role: '' };
  const verifRows = sheetVerif.getDataRange().getValues();
  for (let j = 1; j < verifRows.length; j++) {
    const vNISN = String(verifRows[j][0] || '').trim().replace(/^'/, '');
    const cleanVNISN = vNISN.replace(/^0+/, '') || '0';
    if (cleanVNISN === cleanQuery || vNISN === rawQuery) {
      verificationInfo = {
        status: String(verifRows[j][1] || 'PENDING'),
        note: String(verifRows[j][2] || ''),
        confirmed_at: verifRows[j][3] ? new Date(verifRows[j][3]).toISOString() : '',
        validator_role: String(verifRows[j][4] || 'siswa')
      };
      break;
    }
  }

  return createJsonResponse({
    status: 'success',
    found: matchedDocs.length > 0,
    documents: matchedDocs,
    verification: verificationInfo
  });
}

/**
 * Penanganan permintaan POST (Unggah berkas, autentikasi login, import batch, hapus, dan verifikasi).
 */
function doPost(e) {
  try {
    let payload = {};
    if (e && e.postData && e.postData.contents) {
      payload = JSON.parse(e.postData.contents);
    } else {
      return createJsonResponse({ status: 'error', message: 'Tidak ada payload data yang diterima.' });
    }

    const action = payload.action;

    switch (action) {
      case 'login':
        return handleLoginAuth(payload.username, payload.password, payload.expectedRole);
      
      case 'uploadDocument':
        return handleUploadDocument(payload);

      case 'batchImportDocuments':
        return handleBatchImportDocuments(payload.documents);

      case 'clearAllDocuments':
        return handleClearAllDocuments();

      case 'editDocument':
        return handleEditDocument(payload);

      case 'deleteDocument':
        return handleDeleteDocument(payload.id);

      case 'studentConfirm':
        return handleStudentConfirmation(payload.nisn, payload.status, payload.note, payload.validator_role);

      case 'adminResetLock':
        return handleAdminResetLock(payload.nisn);

      default:
        return createJsonResponse({ status: 'error', message: `Aksi '${action}' tidak dikenali.` });
    }

  } catch (err) {
    return createJsonResponse({ status: 'error', message: err.toString() });
  }
}

/**
 * Mengimpor berkas kolektif dari Excel/CSV dengan mekanisme Upsert (menimpa berkas jika kombinasi NISN dan jenis dokumen sudah ada).
 */
function handleBatchImportDocuments(documents) {
  if (!Array.isArray(documents) || documents.length === 0) {
    return createJsonResponse({ status: 'error', message: 'Data berkas kolektif kosong.' });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    if (!sheetStudents) {
      setupDatabase();
      sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    }

    const existingData = sheetStudents.getDataRange().getValues();
    const rowsToAppend = [];
    let updatedCount = 0;
    let addedCount = 0;

    documents.forEach(doc => {
      const cleanNISN = String(doc.nisn || '').trim().replace(/^'/, '');
      const docNISNFormatted = `'${cleanNISN}`;
      const docJenis = String(doc.jenis_dokumen || '').trim().toLowerCase();
      const docCleanCompare = cleanNISN.replace(/^0+/, '') || '0';

      let matchedRowIdx = -1;
      for (let r = 1; r < existingData.length; r++) {
        const rowNISN = String(existingData[r][1] || '').trim().replace(/^'/, '');
        const rowCleanCompare = rowNISN.replace(/^0+/, '') || '0';
        const rowJenis = String(existingData[r][4] || '').trim().toLowerCase();

        if ((rowCleanCompare === docCleanCompare || rowNISN === cleanNISN) && rowJenis === docJenis) {
          matchedRowIdx = r + 1; // 1-based index di Google Sheets
          break;
        }
      }

      const fileId = String(doc.file_id || '').trim();
      const fileUrl = doc.file_url || (fileId ? `https://lh3.googleusercontent.com/d/${fileId}` : '');
      const timestamp = doc.uploaded_at || new Date().toISOString();

      if (matchedRowIdx !== -1) {
        // Update baris yang sudah ada (menimpa data lama)
        sheetStudents.getRange(matchedRowIdx, 3, 1, 8).setValues([[
          doc.nama_siswa || '',
          doc.kelas || '',
          doc.jenis_dokumen || '',
          fileId,
          fileUrl,
          doc.file_name || '',
          timestamp,
          doc.uploaded_by || 'admin_tu'
        ]]);
        updatedCount++;
      } else {
        // Tambah sebagai baris baru
        const docId = doc.id || `DOC-${new Date().getFullYear()}-${Math.floor(1000 + Math.random() * 9000)}`;
        rowsToAppend.push([
          docId,
          docNISNFormatted,
          doc.nama_siswa || '',
          doc.kelas || '',
          doc.jenis_dokumen || '',
          fileId,
          fileUrl,
          doc.file_name || '',
          timestamp,
          doc.uploaded_by || 'admin_tu'
        ]);
        addedCount++;
      }
    });

    if (rowsToAppend.length > 0) {
      sheetStudents.getRange(sheetStudents.getLastRow() + 1, 1, rowsToAppend.length, 10).setValues(rowsToAppend);
    }

    return createJsonResponse({
      status: 'success',
      message: `Impor kolektif berhasil: ${updatedCount} berkas diperbarui & ${addedCount} berkas baru tersimpan permanen.`,
      updatedCount: updatedCount,
      addedCount: addedCount
    });

  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Membersihkan seluruh baris data arsip di lembar STUDENTS_DATA (menyisakan baris header).
 */
function handleClearAllDocuments() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    if (!sheetStudents) {
      setupDatabase();
      sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    }

    const lastRow = sheetStudents.getLastRow();
    if (lastRow > 1) {
      sheetStudents.deleteRows(2, lastRow - 1);
    }

    return createJsonResponse({
      status: 'success',
      message: 'Seluruh data berkas arsip pada Spreadsheet berhasil dibersihkan.'
    });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mengunggah berkas tunggal (Base64) ke folder Google Drive dan mencatatnya ke Spreadsheet.
 */
function handleUploadDocument(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    if (!sheetStudents) {
      setupDatabase();
      sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    }

    const targetFolder = getOrCreateDriveFolder(DRIVE_FOLDER_NAME);
    let fileId = '';
    let fileUrl = '';
    const cleanNISN = String(payload.nisn).trim().replace(/^'/, '');
    const fileName = `${payload.jenis_dokumen.toLowerCase().replace(/\s+/g, '_')}_${cleanNISN}.jpg`;

    // Konversi gambar Base64 ke Berkas Google Drive jika ada data gambar Base64
    if (payload.base64Data && payload.base64Data.includes('base64,')) {
      const pureBase64 = payload.base64Data.split('base64,')[1];
      const decodedBytes = Utilities.base64Decode(pureBase64);
      const blob = Utilities.newBlob(decodedBytes, 'image/jpeg', fileName);
      const driveFile = targetFolder.createFile(blob);
      driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      fileId = driveFile.getId();
      fileUrl = `https://lh3.googleusercontent.com/d/${fileId}`;
    } else if (payload.file_url) {
      fileUrl = payload.file_url;
      fileId = payload.file_id || '';
    }

    const docId = `DOC-${new Date().getFullYear()}-${Math.floor(1000 + Math.random() * 9000)}`;
    const nowIso = new Date().toISOString();

    // Periksa apakah berkas sejenis untuk siswa ini sudah ada (Upsert)
    const existingData = sheetStudents.getDataRange().getValues();
    const docCleanCompare = cleanNISN.replace(/^0+/, '') || '0';
    const cleanJenis = String(payload.jenis_dokumen).trim().toLowerCase();
    let existingRowIdx = -1;

    for (let r = 1; r < existingData.length; r++) {
      const rNISN = String(existingData[r][1] || '').trim().replace(/^'/, '');
      const rClean = rNISN.replace(/^0+/, '') || '0';
      const rJenis = String(existingData[r][4] || '').trim().toLowerCase();

      if ((rClean === docCleanCompare || rNISN === cleanNISN) && rJenis === cleanJenis) {
        existingRowIdx = r + 1;
        break;
      }
    }

    if (existingRowIdx !== -1) {
      sheetStudents.getRange(existingRowIdx, 3, 1, 8).setValues([[
        payload.nama_siswa,
        payload.kelas,
        payload.jenis_dokumen,
        fileId,
        fileUrl,
        fileName,
        nowIso,
        payload.uploaded_by || 'admin_tu'
      ]]);
    } else {
      sheetStudents.appendRow([
        docId,
        `'${cleanNISN}`,
        payload.nama_siswa,
        payload.kelas,
        payload.jenis_dokumen,
        fileId,
        fileUrl,
        fileName,
        nowIso,
        payload.uploaded_by || 'admin_tu'
      ]);
    }

    return createJsonResponse({
      status: 'success',
      message: 'Berkas berhasil disimpan ke Drive & Spreadsheet.',
      file_id: fileId,
      file_url: fileUrl,
      doc_id: docId
    });

  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Memperbarui data berkas dokumen yang sudah tersimpan.
 */
function handleEditDocument(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    if (!sheetStudents) return createJsonResponse({ status: 'error', message: 'Sheet tidak ditemukan.' });

    const rows = sheetStudents.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(payload.id).trim()) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return createJsonResponse({ status: 'error', message: 'ID dokumen tidak ditemukan.' });
    }

    const cleanNISN = String(payload.nisn).trim().replace(/^'/, '');
    sheetStudents.getRange(targetRow, 2, 1, 6).setValues([[
      `'${cleanNISN}`,
      payload.nama_siswa,
      payload.kelas,
      payload.jenis_dokumen,
      payload.file_id || rows[targetRow - 1][5],
      payload.file_url || rows[targetRow - 1][6]
    ]]);

    return createJsonResponse({ status: 'success', message: 'Data berkas berhasil diperbarui.' });

  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menghapus baris berkas dari Spreadsheet dan memindahkan file di Drive ke kotak sampah.
 */
function handleDeleteDocument(docId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    if (!sheetStudents) return createJsonResponse({ status: 'error', message: 'Sheet tidak ditemukan.' });

    const rows = sheetStudents.getDataRange().getValues();
    let targetRow = -1;
    let fileIdToDelete = '';

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][0]).trim() === String(docId).trim()) {
        targetRow = i + 1;
        fileIdToDelete = String(rows[i][5] || '').trim();
        break;
      }
    }

    if (targetRow === -1) {
      return createJsonResponse({ status: 'error', message: 'Dokumen tidak ditemukan.' });
    }

    // Hapus file di Drive jika fileId valid
    if (fileIdToDelete && fileIdToDelete.length > 15) {
      try {
        const driveFile = DriveApp.getFileById(fileIdToDelete);
        driveFile.setTrashed(true);
      } catch (fErr) {
        console.warn("File Drive tidak ditemukan atau sudah terhapus:", fErr);
      }
    }

    sheetStudents.deleteRow(targetRow);
    return createJsonResponse({ status: 'success', message: 'Berkas berhasil dihapus permanen.' });

  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menyimpan konfirmasi mandiri siswa atau validasi guru (BENAR / SALAH) dan mengunci status.
 */
function handleStudentConfirmation(nisn, status, note, role) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetVerif = ss.getSheetByName(SHEET_VERIF);
    if (!sheetVerif) {
      setupDatabase();
      sheetVerif = ss.getSheetByName(SHEET_VERIF);
    }

    const cleanNISN = String(nisn).trim().replace(/^'/, '');
    const cleanCompare = cleanNISN.replace(/^0+/, '') || '0';
    const rows = sheetVerif.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < rows.length; i++) {
      const vNISN = String(rows[i][0] || '').trim().replace(/^'/, '');
      const vClean = vNISN.replace(/^0+/, '') || '0';
      if (vClean === cleanCompare || vNISN === cleanNISN) {
        targetRow = i + 1;
        break;
      }
    }

    const nowIso = new Date().toISOString();
    const validatorRole = role || 'siswa';

    if (targetRow !== -1) {
      sheetVerif.getRange(targetRow, 2, 1, 4).setValues([[
        status,
        note || '',
        nowIso,
        validatorRole
      ]]);
    } else {
      sheetVerif.appendRow([
        `'${cleanNISN}`,
        status,
        note || '',
        nowIso,
        validatorRole
      ]);
    }

    return createJsonResponse({
      status: 'success',
      message: status === 'BENAR' ? 'Data berhasil disahkan dan dikunci!' : 'Laporan koreksi berhasil dicatat.'
    });

  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mereset status kunci verifikasi siswa agar siswa/guru dapat mengonfirmasi ulang.
 */
function handleAdminResetLock(nisn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetVerif = ss.getSheetByName(SHEET_VERIF);
    if (!sheetVerif) return createJsonResponse({ status: 'error', message: 'Sheet tidak ditemukan.' });

    const cleanNISN = String(nisn).trim().replace(/^'/, '');
    const cleanCompare = cleanNISN.replace(/^0+/, '') || '0';
    const rows = sheetVerif.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < rows.length; i++) {
      const vNISN = String(rows[i][0] || '').trim().replace(/^'/, '');
      const vClean = vNISN.replace(/^0+/, '') || '0';
      if (vClean === cleanCompare || vNISN === cleanNISN) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow !== -1) {
      sheetVerif.deleteRow(targetRow);
    }

    return createJsonResponse({ status: 'success', message: `Status verifikasi NISN ${cleanNISN} berhasil direset.` });

  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Memvalidasi kredensial pengguna (admin atau guru) berdasarkan data di tab ADMIN_USERS.
 */
function handleLoginAuth(username, password, expectedRole) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetAdmins = ss.getSheetByName(SHEET_ADMINS);
  if (!sheetAdmins) {
    setupDatabase();
    sheetAdmins = ss.getSheetByName(SHEET_ADMINS);
  }

  const rows = sheetAdmins.getDataRange().getValues();
  const inputUser = String(username || '').trim().toLowerCase();
  const inputPass = String(password || '').trim();
  const inputRole = String(expectedRole || '').trim().toLowerCase();

  for (let i = 1; i < rows.length; i++) {
    const dbUser = String(rows[i][0] || '').trim().toLowerCase();
    const dbPass = String(rows[i][1] || '').trim();
    const dbRole = String(rows[i][2] || '').trim().toLowerCase();

    if (dbUser === inputUser && dbPass === inputPass) {
      if (!inputRole || dbRole === inputRole) {
        return createJsonResponse({
          status: 'success',
          username: dbUser,
          role: dbRole,
          name: String(rows[i][3] || dbUser)
        });
      }
    }
  }

  return createJsonResponse({
    status: 'error',
    message: 'Username atau password tidak sesuai.'
  });
}

/**
 * Inisialisasi Database Otomatis (Zero-Config).
 * Jalankan fungsi ini sekali di Apps Script untuk membuat semua tab sheet dan folder Drive secara mandiri.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Tab STUDENTS_DATA
  let sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheetStudents) {
    sheetStudents = ss.insertSheet(SHEET_STUDENTS);
    sheetStudents.appendRow([
      "ID_BERKAS",
      "NISN",
      "NAMA_SISWA",
      "KELAS",
      "JENIS_DOKUMEN",
      "FILE_ID_DRIVE",
      "FILE_URL",
      "NAMA_FILE",
      "UPLOADED_AT",
      "UPLOADED_BY"
    ]);
    sheetStudents.getRange("A1:J1").setFontWeight("bold").setBackground("#0284c7").setFontColor("#ffffff");
    sheetStudents.setFrozenRows(1);
  }

  // 2. Tab VERIFIKASI_STATUS
  let sheetVerif = ss.getSheetByName(SHEET_VERIF);
  if (!sheetVerif) {
    sheetVerif = ss.insertSheet(SHEET_VERIF);
    sheetVerif.appendRow([
      "NISN",
      "STATUS_VERIFIKASI",
      "CATATAN_KOREKSI",
      "CONFIRMED_AT",
      "VALIDATOR_ROLE"
    ]);
    sheetVerif.getRange("A1:E1").setFontWeight("bold").setBackground("#4f46e5").setFontColor("#ffffff");
    sheetVerif.setFrozenRows(1);
  }

  // 3. Tab ADMIN_USERS
  let sheetAdmins = ss.getSheetByName(SHEET_ADMINS);
  if (!sheetAdmins) {
    sheetAdmins = ss.insertSheet(SHEET_ADMINS);
    sheetAdmins.appendRow([
      "USERNAME",
      "PASSWORD",
      "ROLE",
      "NAMA_LENGKAP"
    ]);
    sheetAdmins.getRange("A1:D1").setFontWeight("bold").setBackground("#0f172a").setFontColor("#ffffff");
    sheetAdmins.setFrozenRows(1);

    // Tambahkan Akun Bawaan Standar
    sheetAdmins.appendRow(["admin", "123456", "admin", "Administrator Tata Usaha"]);
    sheetAdmins.appendRow(["guru", "123456", "guru", "Wali Kelas / Guru"]);
  }

  // 4. Buat Folder Google Drive jika belum ada
  getOrCreateDriveFolder(DRIVE_FOLDER_NAME);

  console.log("Database dan Folder Drive berhasil disiapkan secara otomatis.");
}

/**
 * Helper untuk mendapatkan atau membuat folder di Google Drive dengan izin akses publik tautan.
 */
function getOrCreateDriveFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    const existingFolder = folders.next();
    existingFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return existingFolder;
  }
  const newFolder = DriveApp.createFolder(folderName);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return newFolder;
}

/**
 * Helper untuk membuat respon JSON standar.
 */
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
