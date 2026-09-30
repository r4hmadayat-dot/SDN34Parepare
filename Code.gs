/**
 * ============================================================
 * KEHADIRAN GURU - GOOGLE APPS SCRIPT
 * File: Code.gs
 * ============================================================
 */

const APP = {
  SHEETS: {
    SCHOOL: 'Data Sekolah',
    GURU: 'DB_GURU',
    ABSEN: 'DB_ABSENSI_GURU',
    PENGAJUAN: 'DB_PENGAJUAN',
    LOG: 'DB_LOG_AKTIVITAS'
  },
  DRIVE_FOLDER: 'ABSENSI GURU - SURAT LAMPIRAN',
  REPORT_FOLDER: 'ABSENSI GURU - REKAP PDF',
  ADMIN_PIN_KEY: 'ADMIN_PIN',
  JAM_MASUK_BATAS: '07:50:00',
  JAM_PULANG_BATAS: '16:50:00'
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Kehadiran Guru - SMP Negeri 2 Soyo Jaya')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .addMetaTag('mobile-web-app-capable', 'yes')
    .addMetaTag('apple-mobile-web-app-capable', 'yes')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ===================== SETUP ===================== */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const structure = {
    'Data Sekolah': ['PARAMETER', 'NILAI'],
    'DB_GURU': ['ID_GURU','NIP','NAMA','JK','JABATAN','NO_WA','KODE_QR','EMAIL','STATUS','PIN_GURU'],
    'DB_ABSENSI_GURU': ['ID_ABSEN','TANGGAL','ID_GURU','NIP','NAMA','JAM_MASUK','JAM_PULANG','STATUS','KETERANGAN','LATITUDE','LONGITUDE','JARAK_METER','AKURASI_METER'],
    'DB_PENGAJUAN': ['ID_PENGAJUAN','TGL_AJU','ID_GURU','NAMA','JENIS','TGL_MULAI','TGL_SELESAI','ALASAN','FILE_LAMPIRAN','STATUS','CATATAN_ADMIN','TGL_PROSES','DISETUJUI_OLEH'],
    'DB_LOG_AKTIVITAS': ['ID_LOG','WAKTU','USER','AKTIVITAS','KETERANGAN']
  };

  Object.keys(structure).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);

    const headers = structure[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
    } else {
      const current = sh.getRange(1,1,1,headers.length).getValues()[0];
      if (current.every(v => v === '')) sh.getRange(1,1,1,headers.length).setValues([headers]);
    }
    // Migrasi header jika sheet lama masih menggunakan struktur 9 kolom.
    if (name === 'DB_ABSENSI_GURU' && sh.getLastColumn() < headers.length) {
      sh.getRange(1,1,1,headers.length).setValues([headers]);
    }
    sh.getRange(1,1,1,headers.length)
      .setFontWeight('bold')
      .setBackground('#1976d2')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
  });

  const school = ss.getSheetByName('Data Sekolah');
  if (school.getLastRow() < 2) {
    school.getRange(2,1,18,2).setValues([
      ['NAMA_SEKOLAH','SMP NEGERI 2 SOYO JAYA'],
      ['NPSN',''],
      ['ALAMAT_SEKOLAH','Desa Lembah Sumara, Kecamatan Soyo Jaya, Kabupaten Morowali Utara, Sulawesi Tengah'],
      ['DESA','Lembah Sumara'],
      ['KECAMATAN','Soyo Jaya'],
      ['KABUPATEN','Morowali Utara'],
      ['PROVINSI','Sulawesi Tengah'],
      ['NAMA_KEPALA_SEKOLAH',''],
      ['NIP_KEPALA_SEKOLAH',''],
      ['TAHUN_PELAJARAN','2026/2027'],
      ['LOGO_SEKOLAH',''],
      ['LATITUDE_SEKOLAH',''],
      ['LONGITUDE_SEKOLAH',''],
      ['RADIUS_ABSENSI','100'],
      ['BATAS_AKURASI_GPS','100'],
      ['EMAIL_SEKOLAH',''],
      ['NO_TELEPON_SEKOLAH',''],
      ['WEBSITE_SEKOLAH','']
    ]);
    school.getRange('A2:A15').setFontWeight('bold');
    school.setColumnWidth(1,230);
    school.setColumnWidth(2,550);
  }

  // Contoh guru hanya dibuat jika DB_GURU masih kosong.
  const guru = ss.getSheetByName('DB_GURU');
  if (guru.getLastRow() < 2) {
    guru.getRange(2,1,1,10).setValues([[
      'GURU-0001','','Nama Guru Contoh','L','Guru Mata Pelajaran','','GURU-0001','','Aktif','1234'
    ]]);
  } else if (guru.getLastColumn() >= 10) {
    // Migrasi akun contoh lama agar dapat diuji dengan login.
    const firstId = String(guru.getRange(2,1).getValue() || '');
    const firstPin = String(guru.getRange(2,10).getValue() || '');
    if (firstId === 'GURU-0001' && !firstPin) guru.getRange(2,10).setValue('1234');
  }

  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty(APP.ADMIN_PIN_KEY)) props.setProperty(APP.ADMIN_PIN_KEY, '123456');

  getOrCreateDriveFolder_();
  formatSheets_();

  return 'Setup selesai. PIN admin awal: 123456 (segera ubah di Pengaturan Admin).';
}

function formatSheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ab = ss.getSheetByName(APP.SHEETS.ABSEN);
  if (ab) {
    ab.getRange('B2:B').setNumberFormat('dd/MM/yyyy');
    ab.getRange('F2:G').setNumberFormat('HH:mm:ss');
  }
  const pg = ss.getSheetByName(APP.SHEETS.PENGAJUAN);
  if (pg) {
    pg.getRange('B2:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
    pg.getRange('F2:G').setNumberFormat('dd/MM/yyyy');
    pg.getRange('L2:L').setNumberFormat('dd/MM/yyyy HH:mm:ss');
  }
  const lg = ss.getSheetByName(APP.SHEETS.LOG);
  if (lg) lg.getRange('B2:B').setNumberFormat('dd/MM/yyyy HH:mm:ss');
}

/* ===================== DATA SEKOLAH ===================== */

function getDataSekolah() {
  const sh = getSheet_(APP.SHEETS.SCHOOL);
  const last = sh.getLastRow();
  const out = {};
  if (last < 2) return out;
  sh.getRange(2,1,last-1,2).getValues().forEach(r => {
    if (r[0]) out[String(r[0])] = r[1] == null ? '' : r[1];
  });
  if (out.LOGO_SEKOLAH) out.LOGO_SEKOLAH = logoClientSource_(out.LOGO_SEKOLAH);
  // Jangan kirim koordinat/radius ke client publik.
  delete out.LATITUDE_SEKOLAH;
  delete out.LONGITUDE_SEKOLAH;
  delete out.RADIUS_ABSENSI;
  delete out.BATAS_AKURASI_GPS;
  delete out.KODE_QR_SEKOLAH;
  return out;
}

function getDataSekolahAdmin(pin) {
  requireAdmin_(pin);
  return getAllSchoolData_();
}

function getAllSchoolData_() {
  const sh = getSheet_(APP.SHEETS.SCHOOL);
  const last = sh.getLastRow();
  const out = {};
  if (last < 2) return out;
  sh.getRange(2,1,last-1,2).getValues().forEach(r => {
    if (r[0]) out[String(r[0])] = r[1] == null ? '' : r[1];
  });
  if (out.LOGO_SEKOLAH) out.LOGO_SEKOLAH = logoClientSource_(out.LOGO_SEKOLAH);
  return out;
}

function updateDataSekolah(data, pin) {
  requireAdmin_(pin);
  data = data || {};
  const sh = getSheet_(APP.SHEETS.SCHOOL);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('Data Sekolah belum tersedia.');

  // LOGO_SEKOLAH dapat dikirim sebagai data URL dari upload file.
  // Logo sekolah bersifat publik, sehingga file disimpan ke Drive dan dibuat dapat dilihat melalui link.
  if (String(data.LOGO_SEKOLAH || '').indexOf('data:image/') === 0) {
    const uploaded = saveSchoolLogo_(data.LOGO_SEKOLAH, String(data.LOGO_FILE_NAME || 'logo-sekolah.png'));
    data.LOGO_SEKOLAH = uploaded.url;
  }

  const rows = sh.getRange(2,1,last-1,2).getValues();
  rows.forEach((r,i) => {
    if (r[0] && Object.prototype.hasOwnProperty.call(data, r[0])) {
      sh.getRange(i+2,2).setValue(data[r[0]]);
    }
  });
  log_('ADMIN','UPDATE_SEKOLAH','Memperbarui data sekolah');
  return getDataSekolah();
}

/* ===================== GURU ===================== */

function getAllGuru() {
  const sh = getSheet_(APP.SHEETS.GURU);
  const last = sh.getLastRow();
  if (last < 2) return [];
  const cols = Math.max(sh.getLastColumn(), 10);
  return sh.getRange(2,1,last-1,cols).getValues()
    .filter(r => r[0] || r[1] || r[2])
    .map(r => ({
      id: String(r[0] || ''),
      nip: String(r[1] || ''),
      nama: String(r[2] || ''),
      jk: String(r[3] || ''),
      jabatan: String(r[4] || ''),
      noWa: String(r[5] || ''),
      qr: String(r[6] || ''),
      email: String(r[7] || ''),
      status: String(r[8] || 'Aktif'),
      pin: String(r[9] || '')
    }));
}

function guruPublic_(g) {
  if (!g) return null;
  return {id:g.id,nip:g.nip,nama:g.nama,jk:g.jk,jabatan:g.jabatan,noWa:g.noWa,email:g.email,status:g.status};
}

function loginGuru(idGuru, pin) {
  const id = String(idGuru || '').trim();
  const pass = String(pin || '').trim();
  if (!id || !pass) throw new Error('Username dan password wajib diisi.');
  const guru = getGuruById_(id);
  if (!guru || String(guru.status).toLowerCase() === 'nonaktif') throw new Error('Akun guru tidak ditemukan atau tidak aktif.');
  if (!guru.pin || guru.pin !== pass) throw new Error('Username atau password salah.');
  const token = Utilities.getUuid();
  CacheService.getScriptCache().put('GURU_SESSION_'+token, guru.id, 21600);
  log_(guru.nama,'LOGIN_GURU','Login berhasil');
  return {token:token,guru:guruPublic_(guru)};
}

function logoutGuru(token) {
  const t = String(token || '');
  if (t) CacheService.getScriptCache().remove('GURU_SESSION_'+t);
  return true;
}

function requireGuruSession_(token) {
  const t = String(token || '');
  if (!t) throw new Error('Sesi login guru tidak ditemukan. Silakan login kembali.');
  const id = CacheService.getScriptCache().get('GURU_SESSION_'+t);
  if (!id) throw new Error('Sesi login telah berakhir. Silakan login kembali.');
  const guru = getGuruById_(id);
  if (!guru || String(guru.status).toLowerCase() === 'nonaktif') throw new Error('Akun guru tidak aktif.');
  return guru;
}

function getGuruProfile(token) {
  return guruPublic_(requireGuruSession_(token));
}

function getGuruByQR(qr) {
  qr = String(qr || '').trim();
  if (!qr) return null;
  const guru = getAllGuru().find(g => g.status.toLowerCase() !== 'nonaktif' && (g.qr === qr || g.id === qr || g.nip === qr));
  return guruPublic_(guru || null);
}

/* ===================== ADMIN: MANAJEMEN USER GURU ===================== */

function getGuruUsers(pin) {
  requireAdmin_(pin);
  return getAllGuru().map(g => ({
    id:g.id,nip:g.nip,nama:g.nama,jk:g.jk,jabatan:g.jabatan,noWa:g.noWa,email:g.email,status:g.status,
    hasPassword:!!g.pin
  }));
}

function validateGuruPayload_(data, isUpdate) {
  data = data || {};
  const id = String(data.id || '').trim();
  const nama = String(data.nama || '').trim();
  const pin = String(data.pin || '').trim();
  if (!id || !nama) throw new Error('ID Guru dan nama wajib diisi.');
  if (!isUpdate && !pin) throw new Error('Password/PIN wajib diisi untuk akun baru.');
  if (pin && !/^\d{4,12}$/.test(pin)) throw new Error('Password/PIN guru harus 4-12 digit angka.');
  return {
    id, nip:String(data.nip||'').trim(), nama, jk:String(data.jk||'').trim(), jabatan:String(data.jabatan||'').trim(),
    noWa:String(data.noWa||'').trim(), email:String(data.email||'').trim(), status:String(data.status||'Aktif').trim() || 'Aktif', pin
  };
}

function createGuruUser(data, pinAdmin) {
  requireAdmin_(pinAdmin);
  const d = validateGuruPayload_(data, false);
  if (getGuruById_(d.id)) throw new Error('ID Guru sudah digunakan.');
  const sh = getSheet_(APP.SHEETS.GURU);
  sh.appendRow([d.id,d.nip,d.nama,d.jk,d.jabatan,d.noWa,'',d.email,d.status,d.pin]);
  log_('ADMIN','TAMBAH_GURU','Menambah akun '+d.id+' - '+d.nama);
  return guruPublic_(getGuruById_(d.id));
}

function updateGuruUser(oldId, data, pinAdmin) {
  requireAdmin_(pinAdmin);
  const old = String(oldId || '').trim();
  const current = getGuruById_(old);
  if (!current) throw new Error('Guru tidak ditemukan.');
  const d = validateGuruPayload_(data, true);
  if (d.id !== old && getGuruById_(d.id)) throw new Error('ID Guru baru sudah digunakan.');
  const sh = getSheet_(APP.SHEETS.GURU);
  const last = sh.getLastRow();
  for (let i=2;i<=last;i++) {
    if (String(sh.getRange(i,1).getValue()) === old) {
      sh.getRange(i,1,1,10).setValues([[d.id,d.nip,d.nama,d.jk,d.jabatan,d.noWa,'',d.email,d.status,d.pin || current.pin]]);
      log_('ADMIN','EDIT_GURU','Mengubah akun '+old+' menjadi '+d.id);
      return guruPublic_(getGuruById_(d.id));
    }
  }
  throw new Error('Data guru tidak ditemukan.');
}

function deleteGuruUser(id, pinAdmin) {
  requireAdmin_(pinAdmin);
  const guru = getGuruById_(id);
  if (!guru) throw new Error('Guru tidak ditemukan.');
  const sh = getSheet_(APP.SHEETS.GURU);
  const last = sh.getLastRow();
  for (let i=2;i<=last;i++) {
    if (String(sh.getRange(i,1).getValue()) === String(id)) {
      sh.deleteRow(i);
      log_('ADMIN','HAPUS_GURU','Menghapus akun '+id+' - '+guru.nama);
      return true;
    }
  }
  throw new Error('Data guru tidak ditemukan.');
}


/* ===================== LOKASI SEKOLAH ===================== */

function getSchoolLocationConfig_() {
  const s = getAllSchoolData_();
  const lat = parseFloat(s.LATITUDE_SEKOLAH);
  const lng = parseFloat(s.LONGITUDE_SEKOLAH);
  const radius = parseFloat(s.RADIUS_ABSENSI || 100);
  const accuracyLimit = parseFloat(s.BATAS_AKURASI_GPS || 100);
  if (!isFinite(lat) || !isFinite(lng)) {
    throw new Error('Koordinat sekolah belum diatur pada menu Data Sekolah.');
  }
  return {
    lat: lat, lng: lng,
    radius: isFinite(radius) ? radius : 100,
    accuracyLimit: isFinite(accuracyLimit) ? accuracyLimit : 100
  };
}

function checkAbsenceLocation(lat, lng, accuracy) {
  return validateLocation_(lat, lng, accuracy);
}

function haversineMeters_(lat1, lon1, lat2, lon2) {
  const R = 6371000;
  const p1 = lat1 * Math.PI / 180;
  const p2 = lat2 * Math.PI / 180;
  const dp = (lat2-lat1) * Math.PI / 180;
  const dl = (lon2-lon1) * Math.PI / 180;
  const a = Math.sin(dp/2)**2 + Math.cos(p1)*Math.cos(p2)*Math.sin(dl/2)**2;
  return R * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1-a));
}

function validateLocation_(lat, lng, accuracy) {
  const c = getSchoolLocationConfig_();
  lat = Number(lat); lng = Number(lng); accuracy = Number(accuracy || 0);
  if (!isFinite(lat) || !isFinite(lng)) throw new Error('Lokasi perangkat tidak valid.');
  if (accuracy > c.accuracyLimit) {
    throw new Error('Akurasi GPS terlalu rendah (' + Math.round(accuracy) + ' m). Aktifkan lokasi presisi lalu coba lagi.');
  }
  const distance = haversineMeters_(lat, lng, c.lat, c.lng);
  if (distance > c.radius) {
    throw new Error('Anda berada di luar area absensi. Jarak sekitar ' + Math.round(distance) +
      ' meter, sedangkan radius yang diizinkan ' + Math.round(c.radius) + ' meter.');
  }
  return {
    ok:true, distance:Math.round(distance), accuracy:Math.round(accuracy),
    schoolLat:c.lat, schoolLng:c.lng, radius:c.radius
  };
}

/* ===================== ABSENSI ===================== */

function getTodayKey_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function getTodayAbsensiGuruSession(token) {
  const guru = requireGuruSession_(token);
  return getTodayAbsensiForGuru(guru.id);
}

function absenMasukGuruSession(token, location) {
  const guru = requireGuruSession_(token);
  return absenMasuk(guru.id, location);
}

function absenPulangGuruSession(token, location) {
  const guru = requireGuruSession_(token);
  return absenPulang(guru.id, location);
}

function getTodayAbsensiForGuru(idGuru) {
  const sh = getSheet_(APP.SHEETS.ABSEN);
  const last = sh.getLastRow();
  if (last < 2) return null;
  const today = getTodayKey_();
  const rows = sh.getRange(2,1,last-1,9).getValues();
  for (let i=rows.length-1; i>=0; i--) {
    const r = rows[i];
    if (String(r[2]) === String(idGuru) && dateKey_(r[1]) === today) {
      return {
        id: r[0], tanggal: dateKey_(r[1]), idGuru: r[2], nip: r[3], nama: r[4],
        jamMasuk: timeText_(r[5]), jamPulang: timeText_(r[6]),
        status: r[7] || 'Hadir', keterangan: r[8] || ''
      };
    }
  }
  return null;
}

function timeToSeconds_(hhmmss) {
  const m = String(hhmmss || '').match(/^(\d{2}):(\d{2}):(\d{2})$/);
  if (!m) return NaN;
  return Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3]);
}

function getJamStatusMasuk_(d) {
  const batas = timeToSeconds_(APP.JAM_MASUK_BATAS);
  const jam = Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm:ss');
  const sekarang = timeToSeconds_(jam);
  return sekarang > batas ? 'Terlambat' : 'Hadir';
}

function addStatusMarker_(status, marker) {
  const base = String(status || 'Hadir').trim();
  const tag = String(marker || '').trim();
  if (!tag) return base;
  if (base.toLowerCase().includes(tag.toLowerCase())) return base;
  return base ? base + ' • ' + tag : tag;
}

function getJamStatusPulang_(d) {
  const batas = timeToSeconds_(APP.JAM_PULANG_BATAS);
  const jam = Utilities.formatDate(d, Session.getScriptTimeZone(), 'HH:mm:ss');
  const sekarang = timeToSeconds_(jam);
  return sekarang < batas ? 'Pulang Cepat' : '';
}

function getAttendanceRules() {
  return {jamMasukBatas: APP.JAM_MASUK_BATAS.slice(0,5), jamPulangBatas: APP.JAM_PULANG_BATAS.slice(0,5)};
}

function absenMasuk(idGuru, location) {
  const guru = getGuruById_(idGuru);
  if (!guru) throw new Error('Guru tidak ditemukan atau tidak aktif.');
  if (!location) throw new Error('Lokasi perangkat diperlukan untuk melakukan absensi.');
  const loc = validateLocation_(location.lat, location.lng, location.accuracy);

  const approved = getApprovedLeaveForDate_(idGuru, new Date());
  if (approved) throw new Error('Hari ini sudah tercatat sebagai ' + approved.jenis + ' yang disetujui.');

  const existing = getTodayAbsensiForGuru(idGuru);
  if (existing && existing.jamMasuk) throw new Error('Anda sudah melakukan absen masuk hari ini pada ' + existing.jamMasuk + '.');

  const sh = getSheet_(APP.SHEETS.ABSEN);
  const now = new Date();

  const statusMasuk = getJamStatusMasuk_(now);
  if (existing) {
    const row = findAbsensiRow_(existing.id);
    sh.getRange(row,6).setValue(now);
    sh.getRange(row,8).setValue(statusMasuk);
    sh.getRange(row,9).setValue(statusMasuk === 'Terlambat' ? 'Absen masuk setelah batas 07:50.' : '');
  } else {
    sh.appendRow([
      generateId_('ABS'),
      now,
      guru.id,
      guru.nip,
      guru.nama,
      now,
      '',
      statusMasuk,
      statusMasuk === 'Terlambat' ? 'Absen masuk setelah batas 07:50.' : '',
      location.lat,
      location.lng,
      loc.distance,
      loc.accuracy
    ]);
  }
  log_(guru.nama,'ABSEN_MASUK','Absen masuk');
  return getTodayAbsensiForGuru(idGuru);
}

function absenPulang(idGuru, location) {
  const guru = getGuruById_(idGuru);
  if (!guru) throw new Error('Guru tidak ditemukan atau tidak aktif.');
  if (!location) throw new Error('Lokasi perangkat diperlukan untuk melakukan absensi.');
  const loc = validateLocation_(location.lat, location.lng, location.accuracy);

  const existing = getTodayAbsensiForGuru(idGuru);
  if (!existing || !existing.jamMasuk) throw new Error('Anda belum melakukan absen masuk hari ini.');
  if (existing.jamPulang) throw new Error('Anda sudah melakukan absen pulang hari ini pada ' + existing.jamPulang + '.');

  const sh = getSheet_(APP.SHEETS.ABSEN);
  const row = findAbsensiRow_(existing.id);
  const now = new Date();
  const early = getJamStatusPulang_(now);
  const updatedStatus = early ? addStatusMarker_(existing.status, early) : existing.status;
  const existingNote = String(existing.keterangan || '').trim();
  const note = early ? (existingNote ? existingNote + ' ' : '') + 'Absen pulang sebelum 16:50.' : existingNote;
  sh.getRange(row,7).setValue(now);
  sh.getRange(row,8).setValue(updatedStatus);
  sh.getRange(row,9).setValue(note);
  sh.getRange(row,10,1,4).setValues([[location.lat, location.lng, loc.distance, loc.accuracy]]);
  log_(guru.nama,'ABSEN_PULANG','Absen pulang; status '+updatedStatus+'; jarak '+loc.distance+' m; akurasi '+loc.accuracy+' m');
  return getTodayAbsensiForGuru(idGuru);
}

function findAbsensiRow_(id) {
  const sh = getSheet_(APP.SHEETS.ABSEN);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('Data absensi tidak ditemukan.');
  const vals = sh.getRange(2,1,last-1,1).getValues();
  for (let i=0;i<vals.length;i++) if (String(vals[i][0]) === String(id)) return i+2;
  throw new Error('Data absensi tidak ditemukan.');
}

/* ===================== IZIN / CUTI / SAKIT ===================== */

function submitPengajuanGuruSession(token, payload) {
  const guru = requireGuruSession_(token);
  payload = payload || {};
  payload.idGuru = guru.id;
  return submitPengajuan(payload);
}

function getPengajuanGuruSession(token) {
  const guru = requireGuruSession_(token);
  return getPengajuanGuru(guru.id);
}

function getRekapGuruSession(token, bulan) {
  const guru = requireGuruSession_(token);
  return buildRekap_(guru.id, bulan);
}

function submitPengajuan(payload) {
  if (!payload || !payload.idGuru) throw new Error('Data guru tidak valid.');
  const guru = getGuruById_(payload.idGuru);
  if (!guru) throw new Error('Guru tidak ditemukan.');
  const jenis = String(payload.jenis || '').trim();
  const mulai = String(payload.tglMulai || '').trim();
  const selesai = String(payload.tglSelesai || '').trim();
  const alasan = String(payload.alasan || '').trim();

  if (!jenis || !mulai || !selesai || !alasan) throw new Error('Jenis, tanggal, dan alasan wajib diisi.');
  if (new Date(mulai) > new Date(selesai)) throw new Error('Tanggal mulai tidak boleh setelah tanggal selesai.');

  let fileUrl = '';
  if (payload.file && payload.file.data) {
    fileUrl = saveBase64File_(payload.file.data, payload.file.name || 'lampiran', payload.file.mimeType || 'application/octet-stream', guru.id);
  }

  const sh = getSheet_(APP.SHEETS.PENGAJUAN);
  sh.appendRow([
    generateId_('PGJ'),
    new Date(),
    guru.id,
    guru.nama,
    jenis,
    new Date(mulai),
    new Date(selesai),
    alasan,
    fileUrl,
    'Menunggu',
    '',
    '',
    ''
  ]);

  log_(guru.nama,'AJUKAN_'+jenis.toUpperCase(),'Pengajuan '+jenis+' '+mulai+' s/d '+selesai);
  return {ok:true, message:'Pengajuan berhasil dikirim dan menunggu persetujuan admin.'};
}

function getPengajuanGuru(idGuru) {
  const sh = getSheet_(APP.SHEETS.PENGAJUAN);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2,1,last-1,13).getValues()
    .filter(r => String(r[2]) === String(idGuru))
    .reverse()
    .map(pengajuanToObj_);
}

function getAllPengajuan(pin) {
  requireAdmin_(pin);
  return getAllPengajuanInternal_();
}

function getAllPengajuanInternal_() {
  const sh = getSheet_(APP.SHEETS.PENGAJUAN);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2,1,last-1,13).getValues().reverse().map(pengajuanToObj_);
}

function prosesPengajuan(idPengajuan, keputusan, catatan, admin, pinAdmin) {
  requireAdmin_(pinAdmin);
  if (!['Disetujui','Ditolak'].includes(keputusan)) throw new Error('Keputusan tidak valid.');
  const sh = getSheet_(APP.SHEETS.PENGAJUAN);
  const last = sh.getLastRow();
  if (last < 2) throw new Error('Pengajuan tidak ditemukan.');
  const vals = sh.getRange(2,1,last-1,13).getValues();
  for (let i=0;i<vals.length;i++) {
    if (String(vals[i][0]) === String(idPengajuan)) {
      const row = i+2;
      sh.getRange(row,10,1,4).setValues([[keputusan, catatan || '', new Date(), admin || 'Admin']]);
      log_(admin || 'Admin','PROSES_PENGAJUAN',idPengajuan+' -> '+keputusan);
      return pengajuanToObj_(sh.getRange(row,1,1,13).getValues()[0]);
    }
  }
  throw new Error('Pengajuan tidak ditemukan.');
}

function pengajuanToObj_(r) {
  return {
    id: r[0], tglAju: dateTimeText_(r[1]), idGuru: r[2], nama: r[3],
    jenis: r[4], tglMulai: dateKey_(r[5]), tglSelesai: dateKey_(r[6]),
    alasan: r[7], file: r[8], status: r[9], catatan: r[10],
    tglProses: dateTimeText_(r[11]), disetujuiOleh: r[12]
  };
}

function getApprovedLeaveForDate_(idGuru, date) {
  const all = getPengajuanGuru(idGuru);
  const d = dateKey_(date);
  return all.find(p => p.status === 'Disetujui' && d >= p.tglMulai && d <= p.tglSelesai) || null;
}

/* ===================== REKAP ===================== */

function getRekapGuru(idGuru, bulan) {
  const guru = getGuruById_(idGuru);
  if (!guru) throw new Error('Guru tidak ditemukan.');
  return buildRekap_(idGuru, bulan);
}

function buildRekap_(idGuru, bulan) {
  const sh = getSheet_(APP.SHEETS.ABSEN);
  const last = sh.getLastRow();
  const map = {};
  if (last >= 2) {
    sh.getRange(2,1,last-1,9).getValues().forEach(r => {
      if (String(r[2]) !== String(idGuru)) return;
      const d = dateKey_(r[1]);
      if (!d.startsWith(bulan)) return;
      map[d] = {
        tanggal:d, masuk:timeText_(r[5]), pulang:timeText_(r[6]),
        status:r[7] || 'Hadir', keterangan:r[8] || ''
      };
    });
  }

  const pg = getPengajuanGuru(idGuru);
  pg.filter(p => p.status === 'Disetujui').forEach(p => {
    let d = new Date(p.tglMulai+'T00:00:00');
    const end = new Date(p.tglSelesai+'T00:00:00');
    while (d <= end) {
      const key = Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      if (key.startsWith(bulan)) {
        map[key] = {tanggal:key, masuk:'', pulang:'', status:p.jenis, keterangan:p.alasan};
      }
      d.setDate(d.getDate()+1);
    }
  });

  const rows = Object.values(map).sort((a,b)=>a.tanggal.localeCompare(b.tanggal));
  const summary = {hadir:0, terlambat:0, pulangCepat:0, izin:0, sakit:0, cuti:0, dinas:0, lainnya:0};
  rows.forEach(r => {
    const s = String(r.status).toLowerCase();
    if (s.includes('terlambat')) summary.terlambat++;
    else if (s.includes('hadir')) summary.hadir++;
    else if (s.includes('izin')) summary.izin++;
    else if (s.includes('sakit')) summary.sakit++;
    else if (s.includes('cuti')) summary.cuti++;
    else if (s.includes('dinas')) summary.dinas++;
    else summary.lainnya++;
    if (s.includes('pulang cepat')) summary.pulangCepat++;
  });
  return {bulan, summary, rows};
}

function getRekapAdmin(bulan, pin) {
  requireAdmin_(pin);
  return getRekapAdminInternal_(bulan);
}

function getRekapAdminInternal_(bulan) {
  return getAllGuru().map(g => {
    const r = buildRekap_(g.id, bulan);
    return {id:g.id,nama:g.nama,nip:g.nip,jabatan:g.jabatan,summary:r.summary};
  });
}

function getAbsensiHariIni(pin) {
  requireAdmin_(pin);
  return getAbsensiHariIniInternal_();
}

function getAbsensiHariIniInternal_() {
  const today = getTodayKey_();
  const sh = getSheet_(APP.SHEETS.ABSEN);
  const rows = [];
  if (sh.getLastRow() >= 2) {
    sh.getRange(2,1,sh.getLastRow()-1,9).getValues().forEach(r => {
      if (dateKey_(r[1]) === today) rows.push({
        id:r[0], idGuru:r[2], nip:r[3], nama:r[4], masuk:timeText_(r[5]),
        pulang:timeText_(r[6]), status:r[7], keterangan:r[8]
      });
    });
  }
  const guru = getAllGuru();
  const existing = {};
  rows.forEach(r=>existing[String(r.idGuru)] = true);
  return guru.map(g => {
    const a = rows.find(x => String(x.idGuru) === String(g.id));
    return a || {idGuru:g.id,nip:g.nip,nama:g.nama,masuk:'',pulang:'',status:'Belum Absen',keterangan:''};
  });
}

function getAdminPrintData(bulan, pin) {
  requireAdmin_(pin);
  const month = bulan || getTodayKey_().slice(0,7);
  const school = getAllSchoolData_();
  const rows = getRekapAdminInternal_(month);
  return {school:school, bulan:month, rows:rows};
}

/**
 * Membuat rekap absensi guru sebagai PDF A4 landscape yang tersimpan di Drive.
 * PDF dibuat melalui Google Docs agar hasilnya benar-benar file PDF, bukan
 * halaman about:blank/HTML yang hanya perlu dicetak manual.
 */
/**
 * Jalankan sekali oleh pemilik script setelah fitur PDF ditambahkan.
 * Tujuannya memunculkan/menetapkan izin Google Docs, Drive, dan UrlFetch
 * sebelum tombol PDF dipakai dari Web App.
 */
function setupPDFPermission() {
  const doc = DocumentApp.create('TEMP - Otorisasi PDF Absensi Guru');
  const id = doc.getId();
  doc.getBody().appendParagraph('Otorisasi PDF Kehadiran Guru.');
  doc.saveAndClose();
  try { DriveApp.getFileById(id).setTrashed(true); } catch(e) {}
  try { UrlFetchApp.fetch('https://www.google.com/generate_204', {muteHttpExceptions:true}); } catch(e) {}
  getOrCreateReportFolder_();
  return 'Otorisasi PDF selesai. Tombol PDF siap digunakan dari Web App.';
}

function generateRekapPDF(bulan, pin) {
  requireAdmin_(pin);
  const month = String(bulan || getTodayKey_().slice(0,7));
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('Periode rekap tidak valid.');

  const school = getSchoolDataRawForReport_();
  const rows = getRekapAdminInternal_(month);
  const monthLabel = monthLabelId_(month);
  const schoolName = String(school.NAMA_SEKOLAH || 'SEKOLAH').trim();
  const reportName = sanitizeReportName_('REKAP ABSENSI GURU - ' + schoolName + ' - ' + monthLabel + '.pdf');
  const tempTitle = 'TEMP REKAP ABSENSI GURU - ' + schoolName + ' - ' + monthLabel + ' - ' + Date.now();
  let doc = null;
  try {
    doc = DocumentApp.create(tempTitle);
    const body = doc.getBody();
    body.clear();
    try {
      body.setPageWidth(841.89);   // A4 landscape (pt)
      body.setPageHeight(595.28);
      body.setMarginTop(22);
      body.setMarginBottom(22);
      body.setMarginLeft(28);
      body.setMarginRight(28);
    } catch (e) {}

    // ===== KOP SURAT =====
    const kop = body.appendTable([['','']]);
    kop.setBorderWidth(0);
    kop.setColumnWidth(0, 90);
    kop.setColumnWidth(1, 680);
    const logoCell = kop.getCell(0,0);
    const textCell = kop.getCell(0,1);
    logoCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
    textCell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);

    const logoBlob = getSchoolLogoBlobForReport_(school.LOGO_SEKOLAH);
    if (logoBlob) {
      const img = logoCell.appendImage(logoBlob);
      const maxW = 56, maxH = 56;
      const iw = img.getWidth(), ih = img.getHeight();
      if (iw > 0 && ih > 0) {
        if (iw >= ih) { img.setWidth(maxW); img.setHeight(Math.round(ih * maxW / iw)); }
        else { img.setHeight(maxH); img.setWidth(Math.round(iw * maxH / ih)); }
      }
    }

    const instansi = textCell.appendParagraph(('PEMERINTAH KABUPATEN ' + String(school.KABUPATEN || '').toUpperCase()).trim());
    instansi.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    instansi.setSpacingAfter(1);
    instansi.editAsText().setFontFamily('Arial').setFontSize(9.5).setBold(true);

    const dinas = textCell.appendParagraph('DINAS PENDIDIKAN');
    dinas.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    dinas.setSpacingAfter(1);
    dinas.editAsText().setFontFamily('Arial').setFontSize(10.5).setBold(true);

    const nama = textCell.appendParagraph(schoolName.toUpperCase());
    nama.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    nama.setSpacingAfter(2);
    nama.editAsText().setFontFamily('Arial').setFontSize(12.5).setBold(true);

    const npsn = school.NPSN ? ('NPSN: ' + school.NPSN) : '';
    const alamatParts = [school.ALAMAT_SEKOLAH, school.DESA, school.KECAMATAN, school.KABUPATEN, school.PROVINSI].filter(Boolean);
    const contactParts = [];
    if (school.EMAIL_SEKOLAH) contactParts.push('Email: ' + school.EMAIL_SEKOLAH);
    if (school.WEBSITE_SEKOLAH) contactParts.push('Website: ' + school.WEBSITE_SEKOLAH);
    const detailText = [npsn, alamatParts.join(', '), contactParts.join(' | ')].filter(Boolean).join('\n');
    const detail = textCell.appendParagraph(detailText);
    detail.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    detail.setSpacingAfter(0);
    detail.editAsText().setFontFamily('Arial').setFontSize(7.5).setBold(false);

    body.appendParagraph('').setSpacingAfter(0);
    const rule = body.appendHorizontalRule();
    body.appendParagraph('').setSpacingAfter(0);

    // ===== JUDUL =====
    const title = body.appendParagraph('REKAPITULASI ABSENSI GURU');
    title.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    title.setSpacingAfter(2);
    title.editAsText().setFontFamily('Arial').setFontSize(12).setBold(true);

    const period = body.appendParagraph('PERIODE: ' + monthLabel.toUpperCase());
    period.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    period.setSpacingAfter(10);
    period.editAsText().setFontFamily('Arial').setFontSize(8.5).setBold(true);

    // ===== TABEL =====
    const tableData = [['NO','NAMA GURU','NIP','HADIR','TERLAMBAT','PULANG CEPAT','IZIN','SAKIT','CUTI','DINAS']];
    rows.forEach((r,i)=> tableData.push([
      String(i+1), String(r.nama || '-'), String(r.nip || '-'),
      String(r.summary.hadir || 0), String(r.summary.terlambat || 0),
      String(r.summary.pulangCepat || 0), String(r.summary.izin || 0), String(r.summary.sakit || 0),
      String(r.summary.cuti || 0), String(r.summary.dinas || 0)
    ]));
    if (rows.length === 0) tableData.push(['-','Tidak ada data guru','-','0','0','0','0','0','0']);

    const table = body.appendTable(tableData);
    table.setBorderWidth(0.75);
    table.setBorderColor('#7f8c8d');
    const widths = [30, 190, 130, 46, 60, 72, 44, 44, 44, 44];
    widths.forEach((w,i)=>{ try{table.setColumnWidth(i,w)}catch(e){} });
    const header = table.getRow(0);
    for (let c=0;c<header.getNumCells();c++) {
      const cell=header.getCell(c);
      cell.setBackgroundColor('#eaf2fb');
      cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
      const par=cell.getChild(0).asParagraph();
      par.setAlignment(c<3?DocumentApp.HorizontalAlignment.CENTER:DocumentApp.HorizontalAlignment.CENTER);
      par.setSpacingBefore(1); par.setSpacingAfter(1);
      par.editAsText().setFontFamily('Arial').setFontSize(7.5).setBold(true);
    }
    for (let r=1;r<table.getNumRows();r++) {
      const row=table.getRow(r);
      for (let c=0;c<row.getNumCells();c++) {
        const cell=row.getCell(c);
        cell.setVerticalAlignment(DocumentApp.VerticalAlignment.CENTER);
        if (r%2===0) cell.setBackgroundColor('#f9fbfd');
        const par=cell.getChild(0).asParagraph();
        par.setAlignment(c===1?DocumentApp.HorizontalAlignment.LEFT:DocumentApp.HorizontalAlignment.CENTER);
        par.setSpacingBefore(1); par.setSpacingAfter(1);
        par.editAsText().setFontFamily('Arial').setFontSize(7.3);
      }
    }

    // ===== CATATAN / TOTAL =====
    const totalGuru = rows.length;
    const totalPar = body.appendParagraph('Jumlah Guru: ' + totalGuru + '    |    Dokumen dibuat otomatis oleh Kehadiran Guru.');
    totalPar.setSpacingBefore(7);
    totalPar.setSpacingAfter(0);
    totalPar.editAsText().setFontFamily('Arial').setFontSize(7.3).setForegroundColor('#5f6b7a');

    // ===== TANDA TANGAN =====
    body.appendParagraph('').setSpacingAfter(0);
    const sigTable = body.appendTable([['', '']]);
    sigTable.setBorderWidth(0);
    sigTable.setColumnWidth(0, 510);
    sigTable.setColumnWidth(1, 260);
    const sig = sigTable.getCell(0,1);
    const cityDate = sig.appendParagraph(formatSignatureDate_(school, new Date()));
    cityDate.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    cityDate.editAsText().setFontFamily('Arial').setFontSize(9);
    const ks = sig.appendParagraph('Kepala Sekolah,');
    ks.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    ks.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true);
    sig.appendParagraph('\n\n');
    const name = sig.appendParagraph(String(school.NAMA_KEPALA_SEKOLAH || '................................................').toUpperCase());
    name.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    name.editAsText().setFontFamily('Arial').setFontSize(9).setBold(true).setUnderline(true);
    const nip = sig.appendParagraph('NIP. ' + String(school.NIP_KEPALA_SEKOLAH || '................................................'));
    nip.setAlignment(DocumentApp.HorizontalAlignment.CENTER);
    nip.editAsText().setFontFamily('Arial').setFontSize(9);

    doc.saveAndClose();
    Utilities.sleep(800);

    const exportUrl = 'https://docs.google.com/document/d/' + doc.getId() + '/export?format=pdf';
    const response = UrlFetchApp.fetch(exportUrl, {
      headers: {Authorization: 'Bearer ' + ScriptApp.getOAuthToken()},
      muteHttpExceptions: true
    });
    if (response.getResponseCode() !== 200) {
      throw new Error('Gagal membuat PDF. HTTP ' + response.getResponseCode() + '. Jalankan fungsi ini lagi setelah memberikan izin Google Docs/Drive.');
    }

    const folder = getOrCreateReportFolder_();
    const pdfFile = folder.createFile(response.getBlob().setName(reportName));
    try { pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
    log_('ADMIN','GENERATE_REKAP_PDF','Membuat ' + reportName);

    return {
      ok: true,
      fileId: pdfFile.getId(),
      fileName: reportName,
      viewUrl: 'https://drive.google.com/file/d/' + pdfFile.getId() + '/view',
      downloadUrl: 'https://drive.google.com/uc?export=download&id=' + pdfFile.getId()
    };
  } finally {
    if (doc) {
      try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch(e) {}
    }
  }
}

function getSchoolDataRawForReport_() {
  const sh = getSheet_(APP.SHEETS.SCHOOL);
  const out = {};
  if (sh.getLastRow() < 2) return out;
  sh.getRange(2,1,sh.getLastRow()-1,2).getValues().forEach(r=>{
    if (r[0]) out[String(r[0])] = r[1] == null ? '' : r[1];
  });
  return out;
}

function getSchoolLogoBlobForReport_(value) {
  const s = String(value || '').trim();
  if (!s) return null;
  try {
    const m = s.match(/^data:(image\/[a-z0-9.+-]+);base64,(.+)$/i);
    if (m) return Utilities.newBlob(Utilities.base64Decode(m[2]), m[1], 'logo-sekolah');
    const m1 = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    const m2 = s.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
    const id = m1 ? m1[1] : (m2 ? m2[1] : '');
    if (id) return DriveApp.getFileById(id).getBlob();
    if (/^https?:\/\//i.test(s)) return UrlFetchApp.fetch(s).getBlob();
  } catch(e) {}
  return null;
}

function getOrCreateReportFolder_() {
  const props = PropertiesService.getScriptProperties();
  const key = 'REPORT_FOLDER_ID';
  const id = props.getProperty(key);
  if (id) {
    try { return DriveApp.getFolderById(id); } catch(e) {}
  }
  const folders = DriveApp.getFoldersByName(APP.REPORT_FOLDER);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(APP.REPORT_FOLDER);
  props.setProperty(key, folder.getId());
  return folder;
}

function monthLabelId_(month) {
  const names=['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const m=String(month).match(/^(\d{4})-(\d{2})$/);
  if (!m) return month;
  const idx=parseInt(m[2],10)-1;
  return idx>=0&&idx<12 ? names[idx]+' '+m[1] : month;
}

function formatSignatureDate_(school, d) {
  const tz = Session.getScriptTimeZone() || 'Asia/Makassar';
  const day = Utilities.formatDate(d, tz, 'dd');
  const month = Utilities.formatDate(d, tz, 'MM');
  const year = Utilities.formatDate(d, tz, 'yyyy');
  const names = ['Januari','Februari','Maret','April','Mei','Juni','Juli','Agustus','September','Oktober','November','Desember'];
  const idx = parseInt(month, 10) - 1;

  // Tempat tanda tangan WAJIB mengambil data dari sheet "Data Sekolah",
  // parameter KECAMATAN. Tidak menggunakan nama daerah hard-code.
  const kecamatan = String((school && school.KECAMATAN) || '').trim();
  const tempat = kecamatan ? kecamatan : '................................';

  return tempat + ', ' + day + ' ' + names[idx] + ' ' + year;
}

function sanitizeReportName_(name) {
  return String(name || 'Rekap Absensi Guru.pdf').replace(/[\\/:*?"<>|]/g,'_');
}

/* ===================== ADMIN ===================== */

function verifyAdmin(pin) {
  const saved = PropertiesService.getScriptProperties().getProperty(APP.ADMIN_PIN_KEY) || '123456';
  return String(pin) === String(saved);
}

function requireAdmin_(pin) {
  if (!verifyAdmin(pin)) throw new Error('Akses admin ditolak. PIN admin tidak valid.');
  return true;
}

function changeAdminPin(oldPin, newPin) {
  if (!verifyAdmin(oldPin)) throw new Error('PIN lama salah.');
  if (!/^\d{4,12}$/.test(String(newPin))) throw new Error('PIN baru harus 4-12 digit angka.');
  PropertiesService.getScriptProperties().setProperty(APP.ADMIN_PIN_KEY, String(newPin));
  return true;
}

function getAdminDashboard(bulan, pin) {
  requireAdmin_(pin);
  const today = getAbsensiHariIniInternal_();
  const counts = {total:getAllGuru().length,hadir:0,terlambat:0,pulangCepat:0,izin:0,sakit:0,cuti:0,belum:0};
  today.forEach(r => {
    const s = String(r.status || '').toLowerCase();
    if (s === 'belum absen') counts.belum++;
    else if (s.includes('terlambat')) counts.terlambat++;
    else if (s.includes('izin')) counts.izin++;
    else if (s.includes('sakit')) counts.sakit++;
    else if (s.includes('cuti')) counts.cuti++;
    else if (s.includes('hadir')) counts.hadir++;
    if (s.includes('pulang cepat')) counts.pulangCepat++;
  });
  return {counts, today, rekap:getRekapAdminInternal_(bulan || getTodayKey_().slice(0,7)), pengajuan:getAllPengajuanInternal_().filter(p=>p.status==='Menunggu')};
}

function saveSchoolLogo_(dataUrl, fileName) {
  const match = String(dataUrl || '').match(/^data:(image\/(?:png|jpe?g|webp));base64,(.+)$/i);
  if (!match) throw new Error('Logo harus berupa PNG, JPG/JPEG, atau WEBP.');
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 2 * 1024 * 1024) throw new Error('Ukuran logo maksimal 2 MB.');

  const safe = String(fileName || 'logo-sekolah.png').replace(/[^\w.\-() ]/g,'_');
  const ext = (match[1].split('/')[1] || 'png').toLowerCase().replace('jpeg','jpg');
  const finalName = 'LOGO_SEKOLAH_' + Date.now() + '.' + ext;
  const folder = getOrCreateDriveFolder_();
  const file = folder.createFile(Utilities.newBlob(bytes, match[1], finalName));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  const directUrl = 'https://drive.google.com/uc?export=view&id=' + file.getId();
  return {id:file.getId(), url:directUrl};
}

function normalizeLogoUrl_(url) {
  const s = String(url || '').trim();
  if (!s) return '';
  const m1 = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  const m2 = s.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  const id = m1 ? m1[1] : (m2 ? m2[1] : '');
  return id ? ('https://drive.google.com/thumbnail?id=' + id + '&sz=w300') : s;
}

/**
 * Menghasilkan sumber gambar yang lebih andal untuk browser HP.
 * Logo kecil dikirim sebagai data URI sehingga tidak bergantung pada
 * kemampuan browser membuka URL Drive. Untuk file lebih besar, gunakan
 * thumbnail Drive agar respons aplikasi tetap ringan.
 */
function logoClientSource_(url) {
  const s = String(url || '').trim();
  if (!s) return '';
  if (s.indexOf('data:image/') === 0) return s;
  const m1 = s.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  const m2 = s.match(/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/);
  const id = m1 ? m1[1] : (m2 ? m2[1] : '');
  if (!id) return s;
  try {
    const cache = CacheService.getScriptCache();
    const key = 'LOGO_DATA_'+id;
    const cached = cache.get(key);
    if (cached) return cached;
    const file = DriveApp.getFileById(id);
    const blob = file.getBlob();
    const bytes = blob.getBytes();
    // Batasi data URI agar tidak membuat respons terlalu besar.
    if (bytes.length <= 400 * 1024) {
      const mime = blob.getContentType() || 'image/png';
      const data = 'data:'+mime+';base64,'+Utilities.base64Encode(bytes);
      try { cache.put(key, data, 21600); } catch (e) {}
      return data;
    }
  } catch (e) {
    // Fallback ke thumbnail Drive jika akses blob gagal.
  }
  return 'https://drive.google.com/thumbnail?id=' + id + '&sz=w300';
}

/* ===================== DRIVE ===================== */

function getOrCreateDriveFolder_() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('DRIVE_FOLDER_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch(e) {}
  }
  const folders = DriveApp.getFoldersByName(APP.DRIVE_FOLDER);
  const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(APP.DRIVE_FOLDER);
  props.setProperty('DRIVE_FOLDER_ID', folder.getId());
  return folder;
}

function saveBase64File_(dataUrl, fileName, mimeType, idGuru) {
  const match = String(dataUrl).match(/^data:([^;]+);base64,(.+)$/);
  if (!match) throw new Error('Format file tidak valid.');
  const bytes = Utilities.base64Decode(match[2]);
  if (bytes.length > 8 * 1024 * 1024) throw new Error('Ukuran file maksimal 8 MB.');
  const safe = String(fileName).replace(/[^\w.\-() ]/g,'_');
  const file = getOrCreateDriveFolder_().createFile(Utilities.newBlob(bytes, mimeType, Date.now()+'_'+idGuru+'_'+safe));
  return file.getUrl();
}

/* ===================== UTILITAS ===================== */

function getSheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Sheet "'+name+'" belum dibuat. Jalankan setupDatabase() terlebih dahulu.');
  return sh;
}

function getGuruById_(id) {
  return getAllGuru().find(g => String(g.id) === String(id));
}

function generateId_(prefix) {
  return prefix+'-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMddHHmmss')+'-'+Math.floor(1000+Math.random()*9000);
}

function dateKey_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  const d = new Date(value);
  return isNaN(d) ? String(value).slice(0,10) : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function timeText_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value)) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm:ss');
  }
  return String(value);
}

function dateTimeText_(value) {
  if (!value) return '';
  const d = new Date(value);
  return isNaN(d) ? String(value) : Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');
}

function log_(user, activity, detail) {
  try {
    getSheet_(APP.SHEETS.LOG).appendRow([generateId_('LOG'),new Date(),user,activity,detail]);
  } catch(e) {}
}
