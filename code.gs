/**
 * =========================================================================
 * PORTAL ARSIP DIGITAL SISWA - GOOGLE DRIVE & SPREADSHEET BACKEND
 * =========================================================================
 * Arsitektur: Zero-Config Google Apps Script (GAS) Enterprise Edition
 * Fitur:
 *  - Pembuatan database otomatis (Sheet & Folder Drive)
 *  - Pencarian berkas berdasarkan NISN (Server-side Indexing)
 *  - Manajemen unggah berkas gambar ke Google Drive (Direct URL lh3)
 *  - Otentikasi hak akses Admin TU & Portal Guru
 *  - Sinkronisasi otomatis konfirmasi Guru ke status Siswa (Benar/Salah)
 * =========================================================================
 */


SHEET_STUDENTS = 'STUDENTS_DATA';
const SHEET_VERIFICATION = 'VERIFIKASI_STATUS';
const SHEET_USERS = 'ADMIN_USERS';
const DRIVE_FOLDER_NAME = 'ARSIP_BERKAS_SISWA_DRIVE';

/**
 * Fungsi inisialisasi otomatis pangkalan data Google Spreadsheet & Drive.
 * Otomatis membuat tab sheet, header kolom, akun default, dan folder Drive.
 */
function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Inisialisasi Sheet STUDENTS_DATA
  let sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheetStudents) {
    sheetStudents = ss.insertSheet(SHEET_STUDENTS);
    sheetStudents.appendRow([
      'id', 'nisn', 'nama_siswa', 'kelas', 'jenis_dokumen',
      'file_id', 'file_url', 'file_name', 'uploaded_at', 'uploaded_by'
    ]);
    sheetStudents.getRange('A1:J1').setFontWeight('bold').setBackground('#0284c7').setFontColor('#ffffff');
    sheetStudents.setFrozenRows(1);
  }

  // 2. Inisialisasi Sheet VERIFIKASI_STATUS
  let sheetVerif = ss.getSheetByName(SHEET_VERIFICATION);
  if (!sheetVerif) {
    sheetVerif = ss.insertSheet(SHEET_VERIFICATION);
    sheetVerif.appendRow([
      'nisn', 'nama_siswa', 'kelas', 'status_siswa', 'waktu_siswa',
      'catatan_siswa', 'status_guru', 'catatan_guru', 'waktu_guru', 'divalidasi_oleh'
    ]);
    sheetVerif.getRange('A1:J1').setFontWeight('bold').setBackground('#0f172a').setFontColor('#38bdf8');
    sheetVerif.setFrozenRows(1);
  }

  // 3. Inisialisasi Sheet ADMIN_USERS (Akun default Admin & Guru)
  let sheetUsers = ss.getSheetByName(SHEET_USERS);
  if (!sheetUsers) {
    sheetUsers = ss.insertSheet(SHEET_USERS);
    sheetUsers.appendRow(['username', 'password', 'nama_lengkap', 'role', 'status', 'created_at']);
    sheetUsers.getRange('A1:F1').setFontWeight('bold').setBackground('#334155').setFontColor('#ffffff');
    sheetUsers.setFrozenRows(1);

    // Tambah akun default awal
    sheetUsers.appendRow(['admin', '123456', 'Administrator TU', 'ADMIN', 'ACTIVE', new Date().toISOString()]);
    sheetUsers.appendRow(['guru', '123456', 'Wali Kelas / Guru Pembina', 'GURU', 'ACTIVE', new Date().toISOString()]);
  }

  // 4. Inisialisasi Folder Google Drive
  getOrCreateDriveFolder();

  return createJsonResponse({
    status: 'success',
    message: 'Setup database dan folder Google Drive berhasil dikonfigurasi!'
  });
}

/**
 * Mendapatkan folder Google Drive penyimpanan arsip atau membuatnya jika belum ada.
 * Mengatur hak akses publik agar dokumen dapat dilihat/diunduh tanpa kendala izin.
 */
function getOrCreateDriveFolder() {
  const folders = DriveApp.getFoldersByName(DRIVE_FOLDER_NAME);
  let folder;
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(DRIVE_FOLDER_NAME);
  }
  
  // Set akses Anyone with link dapat melihat gambar
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (err) {
    Logger.log('Set sharing warning: ' + err.toString());
  }
  
  return folder;
}

/**
 * Web App Entry Point: Menangani HTTP GET (Pencarian NISN, Rekap Data, Health Check)
 */
function doGet(e) {
  try {
    const params = (e && e.parameter) ? e.parameter : {};
    const action = params.action || 'checkHealth';

    if (action === 'checkHealth') {
      return createJsonResponse({ status: 'success', message: 'API Gateway Online', timestamp: new Date().toISOString() });
    }

    if (action === 'setup') {
      return setupDatabase();
    }

    if (action === 'searchNISN') {
      const nisn = (params.nisn || '').trim();
      return handleSearchNISN(nisn);
    }

    if (action === 'getAllDocuments') {
      return handleGetAllDocuments(params);
    }

    if (action === 'getStudentsSummary') {
      return handleGetStudentsSummary(params);
    }

    return createJsonResponse({ status: 'error', message: 'Aksi GET tidak dikenali: ' + action });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  }
}

/**
 * Web App Entry Point: Menangani HTTP POST (Upload File, Login, Konfirmasi Siswa & Guru, CRUD)
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return createJsonResponse({ status: 'error', message: 'Payload data POST tidak ditemukan.' });
    }

    // Parsing payload JSON yang dikirim via text/plain
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    // Route Actions
    switch (action) {
      case 'login':
        return handleLoginAuth(payload.username, payload.password, payload.expectedRole);
      
      case 'uploadDocument':
        return handleUploadDocument(payload);

      case 'editDocument':
        return handleEditDocument(payload);

      case 'deleteDocument':
        return handleDeleteDocument(payload.id);

      case 'studentConfirm':
        return handleStudentConfirmation(payload.nisn, payload.status, payload.note);

      case 'teacherConfirm':
        return handleTeacherConfirmation(payload.nisn, payload.status, payload.note, payload.teacherName);

      case 'resetVerification':
        return handleResetVerification(payload.nisn);

      default:
        return createJsonResponse({ status: 'error', message: 'Aksi POST tidak dikenali: ' + action });
    }
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  }
}

/**
 * Mencari seluruh berkas arsip dan status verifikasi siswa berdasarkan 10 digit NISN.
 */
function handleSearchNISN(nisn) {
  if (!nisn || nisn.length < 5) {
    return createJsonResponse({ status: 'error', message: 'NISN tidak valid.' });
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheetStudents) {
    setupDatabase();
    return createJsonResponse({ status: 'error', message: 'Database baru saja diinisialisasi. Silakan ulangi pencarian.' });
  }

  const values = sheetStudents.getDataRange().getValues();
  if (values.length <= 1) {
    return createJsonResponse({ status: 'success', data: { student: null, documents: [], verification: null } });
  }

  const documents = [];
  let studentProfile = null;

  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const rowNISN = String(row[1]).trim();

    if (rowNISN === nisn) {
      if (!studentProfile) {
        studentProfile = {
          nisn: rowNISN,
          nama_siswa: row[2],
          kelas: row[3]
        };
      }

      documents.push({
        id: row[0],
        nisn: rowNISN,
        nama_siswa: row[2],
        kelas: row[3],
        jenis_dokumen: row[4],
        file_id: row[5],
        file_url: row[6],
        file_name: row[7],
        uploaded_at: row[8],
        uploaded_by: row[9]
      });
    }
  }

  // Ambil data verifikasi dari sheet VERIFIKASI_STATUS
  let verificationRecord = null;
  const sheetVerif = ss.getSheetByName(SHEET_VERIFICATION);
  if (sheetVerif) {
    const verifValues = sheetVerif.getDataRange().getValues();
    for (let j = 1; j < verifValues.length; j++) {
      if (String(verifValues[j][0]).trim() === nisn) {
        verificationRecord = {
          nisn: String(verifValues[j][0]).trim(),
          nama_siswa: verifValues[j][1],
          kelas: verifValues[j][2],
          status_siswa: verifValues[j][3] || 'PENDING',
          waktu_siswa: verifValues[j][4] || '',
          catatan_siswa: verifValues[j][5] || '',
          status_guru: verifValues[j][6] || 'PENDING',
          catatan_guru: verifValues[j][7] || '',
          waktu_guru: verifValues[j][8] || '',
          divalidasi_oleh: verifValues[j][9] || ''
        };
        break;
      }
    }
  }

  return createJsonResponse({
    status: 'success',
    data: {
      student: studentProfile,
      documents: documents,
      verification: verificationRecord
    }
  });
}

/**
 * Mengambil daftar arsip dokumen untuk Portal Admin TU & Portal Guru.
 * Mendukung filter kelas, jenis dokumen, dan query pencarian.
 */
function handleGetAllDocuments(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_STUDENTS);
  if (!sheet) {
    return createJsonResponse({ status: 'success', data: { documents: [], total: 0 } });
  }

  const values = sheet.getDataRange().getValues();
  if (values.length <= 1) {
    return createJsonResponse({ status: 'success', data: { documents: [], total: 0 } });
  }

  const filterKelas = params.kelas || 'ALL';
  const query = (params.query || '').toLowerCase().trim();

  const documents = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    const doc = {
      id: row[0],
      nisn: String(row[1]).trim(),
      nama_siswa: row[2],
      kelas: row[3],
      jenis_dokumen: row[4],
      file_id: row[5],
      file_url: row[6],
      file_name: row[7],
      uploaded_at: row[8],
      uploaded_by: row[9]
    };

    // Filter Kelas
    if (filterKelas !== 'ALL' && doc.kelas !== filterKelas) {
      continue;
    }

    // Search Query (Nama, NISN, Jenis Dokumen)
    if (query) {
      const match = doc.nama_siswa.toLowerCase().includes(query) ||
                    doc.nisn.includes(query) ||
                    doc.jenis_dokumen.toLowerCase().includes(query);
      if (!match) continue;
    }

    documents.push(doc);
  }

  return createJsonResponse({
    status: 'success',
    data: {
      documents: documents,
      total: documents.length
    }
  });
}

/**
 * Mengambil ringkasan data siswa unik beserta status verifikasi (Hasil Siswa / Portal Guru)
 */
function handleGetStudentsSummary(params) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
  const sheetVerif = ss.getSheetByName(SHEET_VERIFICATION);

  if (!sheetStudents) {
    return createJsonResponse({ status: 'success', data: { students: [] } });
  }

  const docValues = sheetStudents.getDataRange().getValues();
  const studentMap = new Map();

  for (let i = 1; i < docValues.length; i++) {
    const nisn = String(docValues[i][1]).trim();
    if (!studentMap.has(nisn)) {
      studentMap.set(nisn, {
        nisn: nisn,
        nama_siswa: docValues[i][2],
        kelas: docValues[i][3],
        total_docs: 1,
        status_siswa: 'PENDING',
        waktu_siswa: '',
        catatan_siswa: '',
        status_guru: 'PENDING',
        catatan_guru: '',
        waktu_guru: '',
        divalidasi_oleh: ''
      });
    } else {
      studentMap.get(nisn).total_docs += 1;
    }
  }

  // Gabungkan dengan data verifikasi
  if (sheetVerif) {
    const verifValues = sheetVerif.getDataRange().getValues();
    for (let j = 1; j < verifValues.length; j++) {
      const nisn = String(verifValues[j][0]).trim();
      if (studentMap.has(nisn)) {
        const item = studentMap.get(nisn);
        item.status_siswa = verifValues[j][3] || 'PENDING';
        item.waktu_siswa = verifValues[j][4] || '';
        item.catatan_siswa = verifValues[j][5] || '';
        item.status_guru = verifValues[j][6] || 'PENDING';
        item.catatan_guru = verifValues[j][7] || '';
        item.waktu_guru = verifValues[j][8] || '';
        item.divalidasi_oleh = verifValues[j][9] || '';
      }
    }
  }

  const result = Array.from(studentMap.values());
  return createJsonResponse({ status: 'success', data: { students: result } });
}

/**
 * Menyimpan file gambar ke Google Drive dan mencatat baris data baru ke Spreadsheet
 */
function handleUploadDocument(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); // Kunci selama maksimal 15 detik

    const nisn = String(payload.nisn).trim();
    const namaSiswa = payload.nama_siswa.trim();
    const kelas = payload.kelas;
    const jenisDokumen = payload.jenis_dokumen;
    const base64Content = payload.base64_data;
    const uploader = payload.uploaded_by || 'admin_tu';

    if (!nisn || !namaSiswa || !base64Content) {
      return createJsonResponse({ status: 'error', message: 'Data NISN, Nama, atau Gambar tidak lengkap.' });
    }

    // Bersihkan header Data URL jika ada
    const cleanBase64 = base64Content.replace(/^data:image\/(png|jpeg|jpg|webp);base64,/, '');
    const decodedBytes = Utilities.base64Decode(cleanBase64);

    // Format nama file
    const safeJenis = jenisDokumen.toLowerCase().replace(/\s+/g, '_');
    const fileName = safeJenis + '_' + nisn + '_' + new Date().getTime() + '.jpg';

    // Simpan ke Google Drive
    const folder = getOrCreateDriveFolder();
    const blob = Utilities.newBlob(decodedBytes, 'image/jpeg', fileName);
    const driveFile = folder.createFile(blob);
    driveFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const fileId = driveFile.getId();
    // Gunakan Direct URL format lh3 untuk penayangan instan di browser
    const fileDirectUrl = 'https://lh3.googleusercontent.com/d/' + fileId;
    const recordId = 'DOC-' + new Date().getFullYear() + '-' + Math.floor(1000 + Math.random() * 9000);
    const nowIso = new Date().toISOString();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    if (!sheetStudents) {
      setupDatabase();
      sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
    }

    sheetStudents.appendRow([
      recordId, nisn, namaSiswa, kelas, jenisDokumen,
      fileId, fileDirectUrl, fileName, nowIso, uploader
    ]);

    return createJsonResponse({
      status: 'success',
      message: 'Berkas ' + jenisDokumen + ' milik ' + namaSiswa + ' berhasil diunggah ke Google Drive!',
      data: {
        id: recordId,
        nisn: nisn,
        file_id: fileId,
        file_url: fileDirectUrl
      }
    });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: 'Gagal mengunggah berkas: ' + error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Memperbarui informasi data arsip berkas (Nama Siswa, NISN, Kelas, Jenis Dokumen)
 */
function handleEditDocument(payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const id = payload.id;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_STUDENTS);
    if (!sheet) return createJsonResponse({ status: 'error', message: 'Sheet tidak ditemukan.' });

    const values = sheet.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow === -1) {
      return createJsonResponse({ status: 'error', message: 'Catatan berkas dengan ID ' + id + ' tidak ditemukan.' });
    }

    // Update kolom: NISN (B), Nama Siswa (C), Kelas (D), Jenis Dokumen (E)
    sheet.getRange(targetRow, 2).setValue(String(payload.nisn).trim());
    sheet.getRange(targetRow, 3).setValue(payload.nama_siswa.trim());
    sheet.getRange(targetRow, 4).setValue(payload.kelas);
    sheet.getRange(targetRow, 5).setValue(payload.jenis_dokumen);

    return createJsonResponse({
      status: 'success',
      message: 'Perubahan data berkas berhasil disimpan!'
    });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Menghapus berkas dari pangkalan data Spreadsheet dan memindahkannya ke Sampah Google Drive
 */
function handleDeleteDocument(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_STUDENTS);
    if (!sheet) return createJsonResponse({ status: 'error', message: 'Sheet tidak ditemukan.' });

    const values = sheet.getDataRange().getValues();
    let targetRow = -1;
    let fileId = '';

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === id) {
        targetRow = i + 1;
        fileId = values[i][5];
        break;
      }
    }

    if (targetRow === -1) {
      return createJsonResponse({ status: 'error', message: 'Berkas tidak ditemukan.' });
    }

    // Hapus dari Google Drive jika ada fileId
    if (fileId) {
      try {
        const file = DriveApp.getFileById(fileId);
        file.setTrashed(true);
      } catch (err) {
        Logger.log('Drive deletion note: ' + err.toString());
      }
    }

    // Hapus baris dari Spreadsheet
    sheet.deleteRow(targetRow);

    return createJsonResponse({
      status: 'success',
      message: 'Berkas arsip berhasil dihapus dari sistem dan Google Drive.'
    });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Mencatat konfirmasi mandiri siswa (Status: BENAR / SALAH dengan Catatan)
 */
function handleStudentConfirmation(nisn, status, note) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const cleanNISN = String(nisn).trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetVerif = ss.getSheetByName(SHEET_VERIFICATION);
    if (!sheetVerif) {
      setupDatabase();
      sheetVerif = ss.getSheetByName(SHEET_VERIFICATION);
    }

    const values = sheetVerif.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === cleanNISN) {
        targetRow = i + 1;
        break;
      }
    }

    const nowIso = new Date().toISOString();

    if (targetRow !== -1) {
      // Perbarui data konfirmasi siswa (Kolom D, E, F)
      sheetVerif.getRange(targetRow, 4).setValue(status);
      sheetVerif.getRange(targetRow, 5).setValue(nowIso);
      sheetVerif.getRange(targetRow, 6).setValue(status === 'SALAH' ? (note || '') : '');
    } else {
      // Ambil nama & kelas dari database siswa
      let studentName = '-';
      let studentClass = '-';
      const sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
      if (sheetStudents) {
        const stVals = sheetStudents.getDataRange().getValues();
        for (let k = 1; k < stVals.length; k++) {
          if (String(stVals[k][1]).trim() === cleanNISN) {
            studentName = stVals[k][2];
            studentClass = stVals[k][3];
            break;
          }
        }
      }

      sheetVerif.appendRow([
        cleanNISN, studentName, studentClass,
        status, nowIso, status === 'SALAH' ? (note || '') : '',
        'PENDING', '', '', ''
      ]);
    }

    return createJsonResponse({
      status: 'success',
      message: 'Status konfirmasi berkas siswa berhasil disimpan dan dikunci!'
    });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Konfirmasi Guru (BENAR / SALAH):
 * Otomatis mensinkronkan dan mengunci status verifikasi siswa secara real-time!
 */
function handleTeacherConfirmation(nisn, status, note, teacherName) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const cleanNISN = String(nisn).trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetVerif = ss.getSheetByName(SHEET_VERIFICATION);
    if (!sheetVerif) {
      setupDatabase();
      sheetVerif = ss.getSheetByName(SHEET_VERIFICATION);
    }

    const values = sheetVerif.getDataRange().getValues();
    let targetRow = -1;

    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === cleanNISN) {
        targetRow = i + 1;
        break;
      }
    }

    const nowIso = new Date().toISOString();
    const validator = teacherName || 'Wali Kelas / Guru';

    if (targetRow !== -1) {
      // 1. Update Kolom Siswa (Otomatis Tersinkron)
      sheetVerif.getRange(targetRow, 4).setValue(status);
      sheetVerif.getRange(targetRow, 5).setValue(nowIso);
      if (status === 'SALAH') {
        sheetVerif.getRange(targetRow, 6).setValue(note || 'Ditetapkan tidak sesuai oleh Guru');
      } else {
        sheetVerif.getRange(targetRow, 6).setValue('');
      }

      // 2. Update Kolom Guru
      sheetVerif.getRange(targetRow, 7).setValue(status);
      sheetVerif.getRange(targetRow, 8).setValue(note || '');
      sheetVerif.getRange(targetRow, 9).setValue(nowIso);
      sheetVerif.getRange(targetRow, 10).setValue(validator);
    } else {
      // Ambil nama & kelas dari database siswa
      let studentName = '-';
      let studentClass = '-';
      const sheetStudents = ss.getSheetByName(SHEET_STUDENTS);
      if (sheetStudents) {
        const stVals = sheetStudents.getDataRange().getValues();
        for (let k = 1; k < stVals.length; k++) {
          if (String(stVals[k][1]).trim() === cleanNISN) {
            studentName = stVals[k][2];
            studentClass = stVals[k][3];
            break;
          }
        }
      }

      sheetVerif.appendRow([
        cleanNISN, studentName, studentClass,
        status, nowIso, status === 'SALAH' ? (note || 'Ditetapkan tidak sesuai oleh Guru') : '',
        status, note || '', nowIso, validator
      ]);
    }

    return createJsonResponse({
      status: 'success',
      message: 'Konfirmasi Guru berhasil disimpan. Status siswa otomatis disinkronkan menjadi ' + status + '!'
    });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Membuka kembali kunci verifikasi siswa agar siswa dapat memvalidasi ulang
 */
function handleResetVerification(nisn) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);

    const cleanNISN = String(nisn).trim();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetVerif = ss.getSheetByName(SHEET_VERIFICATION);
    if (!sheetVerif) return createJsonResponse({ status: 'error', message: 'Sheet verifikasi tidak ditemukan.' });

    const values = sheetVerif.getDataRange().getValues();
    for (let i = 1; i < values.length; i++) {
      if (String(values[i][0]).trim() === cleanNISN) {
        sheetVerif.deleteRow(i + 1);
        break;
      }
    }

    return createJsonResponse({
      status: 'success',
      message: 'Kunci verifikasi untuk NISN ' + cleanNISN + ' telah di-reset.'
    });
  } catch (error) {
    return createJsonResponse({ status: 'error', message: error.toString() });
  } finally {
    lock.releaseLock();
  }
}

/**
 * Validasi otentikasi kredensial pengguna (Admin TU atau Guru)
 */
function handleLoginAuth(username, password, expectedRole) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheetUsers = ss.getSheetByName(SHEET_USERS);
  if (!sheetUsers) {
    setupDatabase();
    sheetUsers = ss.getSheetByName(SHEET_USERS);
  }

  const values = sheetUsers.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    const rowUser = String(values[i][0]).trim();
    const rowPass = String(values[i][1]).trim();
    const rowRole = String(values[i][3]).trim();
    const rowStatus = String(values[i][4]).trim();

    if (rowUser === String(username).trim() && rowPass === String(password).trim()) {
      if (rowStatus !== 'ACTIVE') {
        return createJsonResponse({ status: 'error', message: 'Akun Anda berstatus non-aktif.' });
      }

      if (expectedRole && expectedRole !== 'ALL' && rowRole !== expectedRole) {
        return createJsonResponse({ status: 'error', message: 'Akun tidak memiliki izin untuk peran ini.' });
      }

      return createJsonResponse({
        status: 'success',
        message: 'Login berhasil!',
        data: {
          username: rowUser,
          nama_lengkap: values[i][2],
          role: rowRole
        }
      });
    }
  }

  return createJsonResponse({ status: 'error', message: 'Username atau password salah!' });
}

/**
 * Helper: Mengembalikan respons JSON standar dengan MimeType yang benar
 */
function createJsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
