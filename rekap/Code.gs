/**
 * Penerima unggahan dari web Keuangan GCC -> Google Drive.
 *
 * Yang dilakukan skrip ini:
 *  1. Mengubah isi base64 menjadi file asli (Word / Excel / PDF) dengan tipe MIME yang benar,
 *     sehingga bisa dibuka di Google Docs, Sheets, atau pratinjau Drive.
 *  2. (Opsional) Membuat salinan Google Docs / Google Sheets di folder yang sama
 *     bila web mengirim convert = true. Memerlukan layanan "Drive API" (lihat langkah di bawah).
 *
 * LANGKAH PASANG:
 *  - Buka script.google.com, buka proyek yang sekarang dipakai, ganti seluruh isi Code.gs dengan file ini.
 *  - Untuk salinan Google Docs/Sheets: menu Layanan (ikon +) di kiri -> pilih "Drive API" -> Tambahkan.
 *  - Deploy -> Kelola deployment -> ikon pensil -> Versi: "Versi baru" -> Deploy.
 *    (Pakai deployment yang sama agar URL /exec tidak berubah. Setujui izin akses bila diminta.)
 *  - Akses: Jalankan sebagai "Saya", yang dapat mengakses "Siapa saja".
 */

var MIME_BY_EXT = {
  docx: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
  xlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  pdf: 'application/pdf'
};

var GOOGLE_MIME = {
  document: 'application/vnd.google-apps.document',
  spreadsheet: 'application/vnd.google-apps.spreadsheet'
};

function doGet() {
  return jsonOut_({ status: 'success', message: 'Penerima unggahan Keuangan GCC aktif' });
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (!data.folderId || !data.fileData) throw new Error('Data unggahan tidak lengkap');

    var ext = String(data.ext || '').toLowerCase();
    if (!ext) {
      var m = String(data.fileName || '').match(/\.([a-z0-9]+)$/i);
      ext = m ? m[1].toLowerCase() : '';
    }
    var mime = MIME_BY_EXT[ext] || data.mimeType || 'application/octet-stream';

    var nama = String(data.fileName || ('Rekapan.' + ext));
    if (ext && nama.toLowerCase().slice(-(ext.length + 1)) !== '.' + ext) nama += '.' + ext;

    var bytes = Utilities.base64Decode(data.fileData);
    if (!bytes || !bytes.length) throw new Error('Isi file kosong');
    if (data.size && Math.abs(bytes.length - Number(data.size)) > 0) {
      throw new Error('Ukuran file tidak cocok (diterima ' + bytes.length + ' byte, seharusnya ' + data.size + ')');
    }

    var folder = DriveApp.getFolderById(data.folderId);
    var blob = Utilities.newBlob(bytes, mime, nama);
    var file = folder.createFile(blob);

    var hasil = { status: 'success', fileId: file.getId(), url: file.getUrl(), converted: null };

    if (data.convert && GOOGLE_MIME[data.googleType]) {
      try {
        hasil.converted = konversiKeGoogle_(blob, data.folderId, nama, GOOGLE_MIME[data.googleType]);
      } catch (err) {
        hasil.convertError = String(err);   // file asli tetap tersimpan
      }
    }
    return jsonOut_(hasil);
  } catch (err) {
    return jsonOut_({ status: 'error', message: String(err) });
  }
}

function konversiKeGoogle_(blob, folderId, nama, googleMime) {
  if (typeof Drive === 'undefined') {
    throw new Error('Layanan Drive API belum ditambahkan di proyek Apps Script');
  }
  var dasar = nama.replace(/\.[^.]+$/, '');
  if (Drive.Files.create) {            // Drive API v3
    var f3 = Drive.Files.create({ name: dasar, mimeType: googleMime, parents: [folderId] }, blob);
    return { id: f3.id, url: 'https://drive.google.com/open?id=' + f3.id };
  }
  var f2 = Drive.Files.insert(           // Drive API v2
    { title: dasar, mimeType: googleMime, parents: [{ id: folderId }] }, blob, { convert: true });
  return { id: f2.id, url: 'https://drive.google.com/open?id=' + f2.id };
}

function jsonOut_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
