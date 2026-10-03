/* =========================================================
   E-RAPOR KURIKULUM MERDEKA - GOOGLE APPS SCRIPT BACKEND
   Deploy: Deploy → New Deployment → Web App
   Execute as: Me · Access: Anyone
   ========================================================= */

const CONFIG = {
  MASTER_SS_NAME: 'E-Rapor-DB-Master',
  FOLDER_NAME: 'E-Rapor',
  SECRET: 'CHANGE_THIS_SECRET_KEY_2025',
  SESSION_TTL: 8 * 60 * 60
};

// ================= ENDPOINT =================
function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  try {
    const params = e.parameter || {};
    const body = e.postData ? JSON.parse(e.postData.contents) : {};
    const action = params.action || body.action;

    if (action === 'setup')            return jsonResponse(autoSetupDatabase());
    if (action === 'login')            return jsonResponse(loginUser(body.username, body.password));
    if (action === 'read')             return jsonResponse(readData(body.token, body.sheet, body.filter));
    if (action === 'write')            return jsonResponse(writeData(body.token, body.sheet, body.rows));
    if (action === 'generateDeskripsi')return jsonResponse(generateDeskripsi(body));
    if (action === 'gemini')           return jsonResponse(geminiProxy(body));
    if (action === 'savePDF')          return jsonResponse(savePDFToDrive(body));
    if (action === 'backup')           return jsonResponse(dailyBackup());

    return jsonResponse({ ok: false, message: 'Action tidak dikenali: ' + action });
  } catch (err) {
    return jsonResponse({ ok: false, message: err.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ================= AUTO SETUP =================
function autoSetupDatabase() {
  const log = [];
  const props = PropertiesService.getScriptProperties();

  let ssId = props.getProperty('MASTER_SS_ID');
  let ss;
  if (!ssId) {
    ss = SpreadsheetApp.create(CONFIG.MASTER_SS_NAME);
    ssId = ss.getId();
    props.setProperty('MASTER_SS_ID', ssId);
    log.push('✓ Spreadsheet master dibuat: ' + ssId);
  } else {
    ss = SpreadsheetApp.openById(ssId);
    log.push('✓ Spreadsheet master ditemukan');
  }

  const schemas = {
    'Master_Sekolah':  ['ID','Nama_Sekolah','NPSN','Alamat','Nama_Kepsek','NIP_Kepsek','Logo_URL'],
    'Master_User':     ['User_ID','Username','Password_Hash','Nama_Lengkap','Role','Reference_ID'],
    'Master_Siswa':    ['NISN','NIS','Nama_Siswa','Gender','Kelas_ID','Tempat_Lahir','Tgl_Lahir','Nama_Ortu','Foto'],
    'Master_Mapel_TP': ['Mapel_ID','Nama_Mapel','Kelas_ID','Kode_TP','Deskripsi_TP'],
    'Nilai_Akademik':  ['ID_Nilai','NISN','Mapel_ID','Formatif_Score','Sumatif_Materi','Sumatif_Akhir','Nilai_Akhir','Deskripsi_Auto'],
    'Nilai_P5':        ['ID_P5','NISN','Proyek_ID','Dimensi','Subelemen','Predikat_P5','Catatan_Projek'],
    'Ekskul_Absensi':  ['ID_Ekstra','NISN','Nama_Ekskul','Predikat','Keterangan_Ekskul','Sakit','Izin','Alpha','Catatan_Wali'],
    'Pengaturan':      ['Key','Value','Keterangan']
  };

  Object.keys(schemas).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); log.push('✓ Tab dibuat: ' + name); }
    sh.clear();
    sh.getRange(1, 1, 1, schemas[name].length).setValues([schemas[name]])
      .setFontWeight('bold').setBackground('#0f172a').setFontColor('#ffffff');
    sh.setFrozenRows(1);
  });

  const def = ss.getSheetByName('Sheet1');
  if (def) ss.deleteSheet(def);

  const folderRoot = getOrCreateFolder(CONFIG.FOLDER_NAME);
  ['PDF_Rapor','Foto_Siswa','Aset_Sekolah'].forEach(f => {
    getOrCreateFolder(f, folderRoot);
    log.push('✓ Folder: /' + CONFIG.FOLDER_NAME + '/' + f);
  });

  const userSheet = ss.getSheetByName('Master_User');
  if (userSheet.getLastRow() < 2) {
    userSheet.appendRow(['U001','admin', hashPwd_('admin123'), 'Administrator', 'admin', '-']);
    userSheet.appendRow(['U002','guru',  hashPwd_('guru123'),  'Ibu Sri Wahyuni, S.Pd', 'guru', '-']);
    userSheet.appendRow(['U003','wali',  hashPwd_('wali123'),  'Bapak Hendra Gunawan, S.Pd', 'wali', '-']);
    userSheet.appendRow(['U004','kepsek',hashPwd_('kepsek123'),'Dr. Budi Santoso, M.Pd', 'kepsek', '-']);
    userSheet.appendRow(['U005','siswa', hashPwd_('siswa123'), 'Ahmad Fauzi', 'siswa', '0011223344']);
    log.push('✓ Default users dibuat.');
  }

  return { ok: true, ssId, url: ss.getUrl(), log };
}

function getOrCreateFolder(name, parent) {
  const it = parent ? parent.getFoldersByName(name) : DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return parent ? parent.createFolder(name) : DriveApp.createFolder(name);
}

// ================= AUTH =================
function loginUser(username, password) {
  const ssId = PropertiesService.getScriptProperties().getProperty('MASTER_SS_ID');
  if (!ssId) return { ok: false, message: 'Database belum di-setup. Jalankan autoSetupDatabase.' };

  const sh = SpreadsheetApp.openById(ssId).getSheetByName('Master_User');
  const data = sh.getDataRange().getValues();
  const hash = hashPwd_(password);
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === username && data[i][2] === hash) {
      const token = createToken_(data[i][0], data[i][4]);
      CacheService.getScriptCache().put(token, JSON.stringify({
        userId: data[i][0], role: data[i][4], nama: data[i][3]
      }), CONFIG.SESSION_TTL);
      return { ok: true, token, user: { id: data[i][0], username, role: data[i][4], nama: data[i][3] } };
    }
  }
  return { ok: false, message: 'Username / password salah' };
}

function createToken_(userId, role) {
  const payload = userId + '|' + role + '|' + Date.now();
  const sig = Utilities.computeHmacSha256Signature(payload, CONFIG.SECRET);
  return Utilities.base64EncodeWebSafe(payload + '.' + sig);
}

function validateToken_(token) {
  if (!token) return null;
  const cached = CacheService.getScriptCache().get(token);
  return cached ? JSON.parse(cached) : null;
}

function hashPwd_(pwd) {
  const raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pwd);
  return raw.map(b => ('0' + (b & 0xFF).toString(16)).slice(-2)).join('');
}

// ================= CRUD =================
function readData(token, sheetName, filter) {
  const sess = validateToken_(token);
  if (!sess) return { ok: false, message: 'Session tidak valid / kadaluarsa' };

  const ssId = PropertiesService.getScriptProperties().getProperty('MASTER_SS_ID');
  const sh = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
  if (!sh) return { ok: false, message: 'Sheet tidak ditemukan' };

  const values = sh.getDataRange().getValues();
  const headers = values.shift();
  let rows = values.map(r => {
    const o = {}; headers.forEach((h, i) => o[h] = r[i]); return o;
  });

  if (filter) {
    rows = rows.filter(r => Object.keys(filter).every(k => String(r[k]) === String(filter[k])));
  }
  return { ok: true, sheet: sheetName, headers, rows };
}

function writeData(token, sheetName, rows) {
  const sess = validateToken_(token);
  if (!sess) return { ok: false, message: 'Session tidak valid' };
  if (sess.role === 'siswa' && sheetName !== 'Nilai_Akademik') {
    return { ok: false, message: 'Akses ditolak' };
  }

  const ssId = PropertiesService.getScriptProperties().getProperty('MASTER_SS_ID');
  const sh = SpreadsheetApp.openById(ssId).getSheetByName(sheetName);
  if (!sh) return { ok: false, message: 'Sheet tidak ditemukan' };

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  rows.forEach(r => {
    const arr = headers.map(h => r[h] !== undefined ? r[h] : '');
    const pk = headers[0];
    const existing = sh.getDataRange().getValues().slice(1)
      .findIndex(row => String(row[0]) === String(r[pk]));
    if (existing >= 0) {
      sh.getRange(existing + 2, 1, 1, headers.length).setValues([arr]);
    } else {
      sh.appendRow(arr);
    }
  });
  return { ok: true, message: rows.length + ' baris diproses', count: rows.length };
}

// ================= GENERATOR DESKRIPSI =================
function generateDeskripsi(payload) {
  const { nama_siswa, mapel, nilai_tp } = payload;
  if (!nilai_tp || nilai_tp.length === 0) {
    return { ok: false, message: 'Data TP kosong' };
  }
  const sorted = [...nilai_tp].sort((a, b) => b.nilai - a.nilai);
  const tertinggi = sorted[0];
  const terendah = sorted[sorted.length - 1];
  const narasi = `Menunjukkan penguasaan yang sangat baik dalam ${tertinggi.deskripsi.toLowerCase()}, ` +
                 `namun perlu ditingkatkan dalam ${terendah.deskripsi.toLowerCase()}.`;
  return { ok: true, deskripsi_otomatis: narasi, tp_tertinggi: tertinggi, tp_terendah: terendah };
}

// ================= GEMINI PROXY =================
function geminiProxy(payload) {
  const { apiKey, model = 'gemini-1.5-flash', prompt, temperature = 0.7, maxTokens = 300 } = payload;
  if (!apiKey) return { ok: false, message: 'API Key kosong' };

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
  const body = {
    contents: [{ role: 'user', parts: [{ text: prompt }] }],
    generationConfig: { temperature, maxOutputTokens: maxTokens, topP: 0.95 }
  };

  try {
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(body),
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    const data = JSON.parse(res.getContentText());
    if (code !== 200) return { ok: false, code, message: data.error?.message || 'Gemini error' };
    const text = data.candidates?.[0]?.content?.parts?.[0]?.text?.trim() || '';
    return { ok: true, text, model, usage: data.usageMetadata };
  } catch (e) {
    return { ok: false, message: e.toString() };
  }
}

// ================= PDF =================
function savePDFToDrive(payload) {
  const { filename, base64, nisn } = payload;
  const root = getOrCreateFolder(CONFIG.FOLDER_NAME);
  const pdfFolder = getOrCreateFolder('PDF_Rapor', root);
  const blob = Utilities.newBlob(
    Utilities.base64Decode(base64),
    'application/pdf',
    filename || (nisn + '_rapor.pdf')
  );
  const file = pdfFolder.createFile(blob);
  file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
  return { ok: true, fileId: file.getId(), url: file.getUrl() };
}

// ================= BACKUP =================
function dailyBackup() {
  const ssId = PropertiesService.getScriptProperties().getProperty('MASTER_SS_ID');
  if (!ssId) return { ok: false, message: 'Database belum di-setup' };
  const ss = SpreadsheetApp.openById(ssId);
  const root = getOrCreateFolder(CONFIG.FOLDER_NAME);
  const backupFolder = getOrCreateFolder('Backup', root);
  const copy = DriveApp.getFileById(ssId).makeCopy(
    'Backup_' + ss.getName() + '_' + Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss'),
    backupFolder
  );
  return { ok: true, url: copy.getUrl() };
}

/* =========================================================
   SETUP INSTRUCTION:
   1. Deploy → New Deployment → Web App
      - Execute as: Me
      - Access: Anyone
   2. Copy URL → paste ke assets/app.js API_URL
   3. Jalankan fungsi autoSetupDatabase sekali dari editor
   4. (Opsional) Buat trigger harian untuk dailyBackup
   ========================================================= */