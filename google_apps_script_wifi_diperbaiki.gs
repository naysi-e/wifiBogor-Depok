/**
 * Pendaftaran WiFi: simpan ke Google Sheets lalu kirim notifikasi email ke admin.
 * Pasang script ini dari spreadsheet tujuan: Ekstensi > Apps Script.
 */

const SHEET_NAME = 'Pendaftaran';
const ADMIN_EMAIL = 'NeisyaZaqia@gmail.com';

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    const ss = getSpreadsheet_();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Waktu Pendaftaran', 'Nama Lengkap', 'Nomor WhatsApp',
        'Kota/Kabupaten', 'Kecamatan', 'Kelurahan/Desa',
        'Alamat Lengkap', 'Paket WiFi', 'Catatan', 'Sumber'
      ]);
      sheet.setFrozenRows(1);
    }

    const p = (e && e.parameter) ? e.parameter : {};
    const data = {
      waktu: new Date(),
      nama: safe_(p.nama),
      telepon: safe_(p.telepon),
      kota: safe_(p.kota),
      kecamatan: safe_(p.kecamatan),
      kelurahan: safe_(p.kelurahan || '-'),
      alamat: safe_(p.alamat),
      paket: safe_(p.paket),
      catatan: safe_(p.catatan || '-'),
      sumber: safe_(p.sumber || 'Website WiFi')
    };

    if (!data.nama || !data.telepon || !data.kota ||
        !data.kecamatan || !data.alamat || !data.paket) {
      throw new Error('Data wajib belum lengkap.');
    }

    // Simpan data lebih dahulu.
    sheet.appendRow([
      data.waktu, data.nama, data.telepon, data.kota,
      data.kecamatan, data.kelurahan, data.alamat,
      data.paket, data.catatan, data.sumber
    ]);
    SpreadsheetApp.flush();

    // Kirim email notifikasi setelah data tersimpan.
    const time = Utilities.formatDate(
      data.waktu,
      Session.getScriptTimeZone(),
      'dd/MM/yyyy HH:mm:ss'
    );

    const subject = 'Pendaftar pemasangan WiFi baru dari website';
    const body = [
      'Ada pendaftar pemasangan WiFi baru dari website.',
      '',
      'Waktu pendaftaran: ' + time,
      'Nama pelanggan: ' + data.nama,
      'Nomor WhatsApp: ' + data.telepon,
      'Kota/Kabupaten: ' + data.kota,
      'Kecamatan: ' + data.kecamatan,
      'Kelurahan/Desa: ' + data.kelurahan,
      'Alamat: ' + data.alamat,
      'Paket WiFi: ' + data.paket,
      'Catatan: ' + data.catatan,
      'Sumber: ' + data.sumber,
      '',
      'Google Sheets: ' + ss.getUrl()
    ].join('\n');

    MailApp.sendEmail({
      to: ADMIN_EMAIL,
      subject: subject,
      body: body
    });

    return json_({ ok: true, saved: true, emailSent: true });
  } catch (err) {
    console.error('Pendaftaran gagal: ' + (err && err.stack ? err.stack : err));
    return json_({
      ok: false,
      saved: false,
      error: String(err && err.message ? err.message : err)
    });
  } finally {
    try {
      lock.releaseLock();
    } catch (_) {}
  }
}

function getSpreadsheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'Spreadsheet tidak ditemukan. Buka Apps Script dari Ekstensi > Apps Script pada spreadsheet tujuan.'
    );
  }
  return ss;
}

function doGet() {
  return ContentService.createTextOutput('Endpoint pendaftaran WiFi aktif.');
}

// Jalankan manual sekali untuk menyiapkan tab dan meminta izin spreadsheet.
function testSetup() {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      'Waktu Pendaftaran', 'Nama Lengkap', 'Nomor WhatsApp',
      'Kota/Kabupaten', 'Kecamatan', 'Kelurahan/Desa',
      'Alamat Lengkap', 'Paket WiFi', 'Catatan', 'Sumber'
    ]);
    sheet.setFrozenRows(1);
  }

  Logger.log('Spreadsheet siap: ' + ss.getUrl());
  return 'Spreadsheet siap digunakan.';
}

// Jalankan manual untuk menguji pengiriman email ke admin.
function testEmail() {
  MailApp.sendEmail({
    to: ADMIN_EMAIL,
    subject: 'Tes notifikasi pendaftaran WiFi',
    body:
      'Halo Admin,\n\n' +
      'Ini email percobaan dari sistem pendaftaran WiFi.\n' +
      'Jika email ini diterima, pengiriman email dari Apps Script berhasil.\n\n' +
      'Alamat tujuan: ' + ADMIN_EMAIL
  });

  Logger.log('Email percobaan dikirim ke ' + ADMIN_EMAIL);
  return 'Email percobaan diminta dikirim ke ' + ADMIN_EMAIL;
}

function safe_(value) {
  let text = String(value == null ? '' : value).trim();
  // Cegah input pelanggan ditafsirkan sebagai formula spreadsheet.
  if (/^[=+@-]/.test(text)) text = "'" + text;
  return text;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
