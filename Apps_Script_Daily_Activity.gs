/**
 * CITA MOBILE — penerima data Daily Activity → sheet "Input" di Daily Monitoring HSE Site Sandai.
 *
 * Pemasangan: buka Google Sheet (hasil "Save as Google Sheets") > Extensions > Apps Script >
 * hapus isi lama, tempel kode ini > ganti SECRET > Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone) > salin URL Web App.
 *
 * Jika script dibuat dari script.google.com (bukan dari menu Ekstensi di dalam Sheet), isi SPREADSHEET_ID.
 * Yang ditulis HANYA kolom kuning di sheet Input: B Tanggal, C Nama Personil, F Kegiatan,
 * H Jam Mulai, I Jam Selesai, K Status, L Keterangan. Kolom abu-abu (rumus) tidak disentuh.
 * ID entri aplikasi disimpan di kolom AX ("ID App") supaya edit/hapus bisa menemukan barisnya.
 */
const SECRET = 'cita-cc896bdc7e5b';   // harus sama dengan DAILY_SYNC_SECRET di index.html
const SPREADSHEET_ID = '1TY9rPfQQsMgYpmsR0FR1IL_f0buTIw9eiezxTR2pMfs'; // salin dari alamat Google Sheet: bagian antara /d/ dan /edit
const SHEET_NAME = 'Input';
const FIRST_ROW = 5;          // baris data pertama
const ID_COL = 50;            // kolom AX
const COL = { tanggal:2, nama:3, kegiatan:6, mulai:8, selesai:9, status:11, ket:12 };

const VERSI = 'v7';
// Buka URL Web App di browser: kalau tampil teks ini, deployment sudah memakai kode terbaru.
function doGet(){
  return ContentService.createTextOutput('CITA Daily Activity sync aktif · ' + VERSI);
}

function doPost(e){
  try{
    const req = JSON.parse(e.postData.contents);
    if(req.secret !== SECRET) return out_({ok:false, error:'secret salah'});
    const lock = LockService.getScriptLock(); lock.waitLock(60000);
    try{
      const ss = SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(SHEET_NAME);
      if(!sh) return out_({ok:false, error:'sheet Input tidak ditemukan'});
      if(sh.getMaxColumns() < ID_COL) sh.insertColumnsAfter(sh.getMaxColumns(), ID_COL - sh.getMaxColumns());
      if(!sh.getRange(FIRST_ROW-1, ID_COL).getValue()) sh.getRange(FIRST_ROW-1, ID_COL).setValue('ID App');

      if(req.action === 'bulk') return out_(bulk_(ss, sh, req.items || []));

      const n = sh.getMaxRows() - FIRST_ROW + 1;
      const tgl = sh.getRange(FIRST_ROW, COL.tanggal, n, 1).getValues();
      let last = FIRST_ROW - 1;                       // baris terakhir yang berisi tanggal
      for(let i=n-1;i>=0;i--){ if(tgl[i][0] !== '' && tgl[i][0] !== null){ last = FIRST_ROW + i; break; } }
      const ids = last >= FIRST_ROW ? sh.getRange(FIRST_ROW, ID_COL, last-FIRST_ROW+1, 1).getValues() : [];
      let row = -1;
      for(let i=0;i<ids.length;i++){ if(String(ids[i][0]) === String(req.id)){ row = FIRST_ROW + i; break; } }

      if(req.action === 'delete'){
        if(row > 0) removeRow_(sh, row, last);
        return out_({ok:true});
      }
      const r = req.row;
      if(row < 0) row = last + 1;
      const tz = ss.getSpreadsheetTimeZone();
      ensureFormulas_(sh, row, last);   // pastikan kolom rumus (No, Jabatan, Section, Frekuensi, Lama, helper) ada di baris ini
      setVal_(sh.getRange(row, COL.tanggal), Utilities.parseDate(r['Tanggal'], tz, 'yyyy-MM-dd'));
      setVal_(sh.getRange(row, COL.nama), r['Nama Personil']);
      SpreadsheetApp.flush();   // beri waktu rumus Jabatan / daftar Kegiatan terhitung dulu
      setVal_(sh.getRange(row, COL.kegiatan), r['Kegiatan']);
      setVal_(sh.getRange(row, COL.mulai), frac_(r['Jam Mulai']));
      setVal_(sh.getRange(row, COL.selesai), frac_(r['Jam Selesai']));
      setVal_(sh.getRange(row, COL.status), r['Status']);
      setVal_(sh.getRange(row, COL.ket), r['Keterangan']);
      sh.getRange(row, ID_COL).setValue(req.id);
      return out_({ok:true, row:row});
    } finally { lock.releaseLock(); }
  }catch(err){ return out_({ok:false, error:String(err)}); }
}

// Kirim banyak entri sekaligus (dipakai tombol "Kirim Ulang Semua"): baris dicari sekali, lalu semua entri ditulis berurutan.
function bulk_(ss, sh, items){
  const tz = ss.getSpreadsheetTimeZone();
  const n = sh.getMaxRows() - FIRST_ROW + 1;
  const tgl = sh.getRange(FIRST_ROW, COL.tanggal, n, 1).getValues();
  let last = FIRST_ROW - 1;
  for(let i=n-1;i>=0;i--){ if(tgl[i][0] !== '' && tgl[i][0] !== null){ last = FIRST_ROW + i; break; } }
  const map = {};
  if(last >= FIRST_ROW){
    const ids = sh.getRange(FIRST_ROW, ID_COL, last-FIRST_ROW+1, 1).getValues();
    for(let i=0;i<ids.length;i++){ if(ids[i][0] !== '') map[String(ids[i][0])] = FIRST_ROW + i; }
  }
  let done = 0; const errors = [];
  items.forEach(function(it){
    try{
      const r = it.row; let row = map[String(it.id)];
      if(!row){ row = last + 1; ensureFormulas_(sh, row, last); }
      setVal_(sh.getRange(row, COL.tanggal), Utilities.parseDate(r['Tanggal'], tz, 'yyyy-MM-dd'));
      setVal_(sh.getRange(row, COL.nama), r['Nama Personil']);
      SpreadsheetApp.flush();
      setVal_(sh.getRange(row, COL.kegiatan), r['Kegiatan']);
      setVal_(sh.getRange(row, COL.mulai), frac_(r['Jam Mulai']));
      setVal_(sh.getRange(row, COL.selesai), frac_(r['Jam Selesai']));
      setVal_(sh.getRange(row, COL.status), r['Status']);
      setVal_(sh.getRange(row, COL.ket), r['Keterangan']);
      sh.getRange(row, ID_COL).setValue(it.id);
      if(!map[String(it.id)]){ map[String(it.id)] = row; last = row; }
      done++;
    }catch(err){ errors.push(String(it.id) + ': ' + String(err)); }
  });
  return {ok: errors.length === 0, done: done, total: items.length, error: errors.slice(0,3).join(' | ')};
}

// Hapus entri: data di bawahnya digeser naik supaya tidak ada baris kosong (rumus kolom abu-abu tetap utuh).
function removeRow_(sh, row, last){
  const cols = [COL.tanggal, COL.nama, COL.kegiatan, COL.mulai, COL.selesai, COL.status, COL.ket, ID_COL];
  const cnt = last - row;
  cols.forEach(c=>{
    if(cnt > 0){
      const below = sh.getRange(row+1, c, cnt, 1);
      below.copyTo(sh.getRange(row, c, cnt, 1), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
    }
    sh.getRange(last, c).clearContent();
  });
}
// Salin rumus dari baris data terakhir ke baris baru HANYA jika sel rumus di baris baru masih kosong.
// Kolom rumus: A (No), D, E, G, J, dan M:AV (helper). Kolom input kuning tidak disalin.
function ensureFormulas_(sh, row, last){
  const src = last >= FIRST_ROW ? last : FIRST_ROW;
  if(src === row) return;
  const cols = [[1,1],[4,5],[7,7],[10,10],[13,48]];
  cols.forEach(c=>{
    const n = c[1]-c[0]+1;
    const target = sh.getRange(row, c[0], 1, n);
    const has = target.getFormulas()[0].some(f=>f !== '');
    if(has) return;
    const source = sh.getRange(src, c[0], 1, n);
    if(source.getFormulas()[0].some(f=>f !== '')){
      source.copyTo(target, SpreadsheetApp.CopyPasteType.PASTE_FORMULA, false);
    }
  });
}

// Tulis nilai ke sel dengan melepas aturan validasi sebentar (dropdown), lalu memasangnya kembali.
function setVal_(range, value, fmt){
  const dv = range.getDataValidation();
  if(dv) range.clearDataValidations();
  range.setValue(value);
  if(fmt) range.setNumberFormat(fmt);
  if(dv) range.setDataValidation(dv);
}
function frac_(hhmm){
  const p = String(hhmm).split(':');
  return (Number(p[0])*60 + Number(p[1])) / 1440;
}
function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

// ===== UJI COBA (jalankan dari editor Apps Script, pilih nama fungsi di sebelah tombol "Jalankan") =====
// tesKirim: menulis 1 baris contoh ke sheet Input. Pertama kali dijalankan akan meminta otorisasi (Izinkan).
function tesKirim(){
  const e = { postData:{ contents: JSON.stringify({
    secret: SECRET, action:'upsert', id:'tes-manual-1',
    row:{ 'Tanggal':'2026-10-08', 'Nama Personil':'Idola Kulana', 'Kegiatan':'P5M', 'Jam Mulai':'07:00', 'Jam Selesai':'07:30', 'Status':'Selesai', 'Keterangan':'TES dari Apps Script' }
  }) } };
  Logger.log(doPost(e).getContent());
}
// tesHapus: menghapus baris contoh tadi.
function tesHapus(){
  const e = { postData:{ contents: JSON.stringify({ secret: SECRET, action:'delete', id:'tes-manual-1' }) } };
  Logger.log(doPost(e).getContent());
}
