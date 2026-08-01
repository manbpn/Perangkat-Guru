/**
 * Backend Google Apps Script untuk Ruang Guru MAN Balikpapan
 *
 * WAJIB ada 3 tab di spreadsheet ini, dengan nama PERSIS seperti ini
 * (huruf besar/kecil harus sama), masing-masing header di baris 1,
 * SATU KATA PER KOLOM (jangan digabung dalam satu sel):
 *
 * 1) Tab "Data" -> kolom A sampai K:
 *    id | nama | nip | jabatan | mapel | link | status | createdAt | verifiedBy | verifiedAt | rejectReason
 *
 * 2) Tab "Verifikator" -> kolom A sampai F:
 *    id | nama | kode | status | createdAt | updatedAt
 *
 * 3) Tab "LoginLog" -> kolom A sampai C:
 *    waktu | peran | nama
 */

const SHEET_GURU = 'Data';
const SHEET_VERIFIKATOR = 'Verifikator';
const SHEET_LOGIN = 'LoginLog';

function sheet_(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}

function rowsToObjects_(sheet) {
  if (!sheet) return [];
  const values = sheet.getDataRange().getValues();
  const headerRow = values.shift();
  if (!headerRow) return [];
  const headers = headerRow.map(h => h.toString().trim());
  return values
    .filter(row => row[0] !== '')
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function findRowIndexById_(sheet, id) {
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) return i + 1; // nomor baris 1-based
  }
  return -1;
}

function doGet(e) {
  const type = ((e.parameter && e.parameter.type) || 'guru').toLowerCase();
  let data;
  if (type === 'verifikator') {
    data = rowsToObjects_(sheet_(SHEET_VERIFIKATOR));
  } else if (type === 'loginlog') {
    data = rowsToObjects_(sheet_(SHEET_LOGIN));
  } else {
    data = rowsToObjects_(sheet_(SHEET_GURU));
  }
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, data }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const action = body.action;
  let result = { ok: true };

  if (action === 'create') {
    sheet_(SHEET_GURU).appendRow([
      body.id, body.nama, body.nip, body.jabatan, body.mapel, body.link,
      'pending', new Date().toISOString(), '', '', ''
    ]);

  } else if (action === 'update') {
    const sheet = sheet_(SHEET_GURU);
    const row = findRowIndexById_(sheet, body.id);
    if (row > 0) {
      // Kolom isi data (dipakai saat Guru merevisi datanya sendiri)
      if (body.nama !== undefined) sheet.getRange(row, 2).setValue(body.nama);
      if (body.nip !== undefined) sheet.getRange(row, 3).setValue(body.nip);
      if (body.jabatan !== undefined) sheet.getRange(row, 4).setValue(body.jabatan);
      if (body.mapel !== undefined) sheet.getRange(row, 5).setValue(body.mapel);
      if (body.link !== undefined) sheet.getRange(row, 6).setValue(body.link);
      // Kolom status verifikasi (dipakai saat Verifikator/Admin menyetujui/menolak)
      if (body.status !== undefined) sheet.getRange(row, 7).setValue(body.status);
      if (body.verifiedBy !== undefined) sheet.getRange(row, 9).setValue(body.verifiedBy);
      if (body.verifiedAt !== undefined) sheet.getRange(row, 10).setValue(body.verifiedAt);
      if (body.rejectReason !== undefined) sheet.getRange(row, 11).setValue(body.rejectReason);
    }

  } else if (action === 'delete') {
    const sheet = sheet_(SHEET_GURU);
    const row = findRowIndexById_(sheet, body.id);
    if (row > 0) sheet.deleteRow(row);

  } else if (action === 'login') {
    sheet_(SHEET_LOGIN).appendRow([new Date().toISOString(), body.peran || '', body.nama || '']);

  } else if (action === 'createVerifikator') {
    sheet_(SHEET_VERIFIKATOR).appendRow([
      body.id, body.nama, body.kode, 'aktif', new Date().toISOString(), ''
    ]);

  } else if (action === 'resetVerifikatorKode') {
    const sheet = sheet_(SHEET_VERIFIKATOR);
    const row = findRowIndexById_(sheet, body.id);
    if (row > 0) {
      sheet.getRange(row, 3).setValue(body.kode || '');
      sheet.getRange(row, 6).setValue(new Date().toISOString());
    }

  } else if (action === 'toggleVerifikatorStatus') {
    const sheet = sheet_(SHEET_VERIFIKATOR);
    const row = findRowIndexById_(sheet, body.id);
    if (row > 0) {
      sheet.getRange(row, 4).setValue(body.status || 'aktif');
      sheet.getRange(row, 6).setValue(new Date().toISOString());
    }

  } else {
    result = { ok: false, error: 'Aksi tidak dikenal' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}