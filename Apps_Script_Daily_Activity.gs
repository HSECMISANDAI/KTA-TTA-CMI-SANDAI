/**
 * CITA MOBILE — penerima data Daily Activity → sheet "Input" di Daily Monitoring HSE Site Sandai.
 *
 * Pemasangan: buka Google Sheet (hasil "Save as Google Sheets") > Extensions > Apps Script >
 * hapus isi lama, tempel kode ini > ganti SECRET > Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone) > salin URL Web App.
 *
 * Yang ditulis HANYA kolom kuning di sheet Input: B Tanggal, C Nama Personil, F Kegiatan,
 * H Jam Mulai, I Jam Selesai, K Status, L Keterangan. Kolom abu-abu (rumus) tidak disentuh.
 * ID entri aplikasi disimpan di kolom AX ("ID App") supaya edit/hapus bisa menemukan barisnya.
 */
const SECRET = 'GANTI_DENGAN_KODE_RAHASIA';   // harus sama dengan DAILY_SYNC_SECRET di index.html
const SHEET_NAME = 'Input';
const FIRST_ROW = 5;          // baris data pertama
const ID_COL = 50;            // kolom AX
const COL = { tanggal:2, nama:3, kegiatan:6, mulai:8, selesai:9, status:11, ket:12 };

function doPost(e){
  try{
    const req = JSON.parse(e.postData.contents);
    if(req.secret !== SECRET) return out_({ok:false, error:'secret salah'});
    const lock = LockService.getScriptLock(); lock.waitLock(25000);
    try{
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sh = ss.getSheetByName(SHEET_NAME);
      if(!sh) return out_({ok:false, error:'sheet Input tidak ditemukan'});
      if(sh.getMaxColumns() < ID_COL) sh.insertColumnsAfter(sh.getMaxColumns(), ID_COL - sh.getMaxColumns());
      if(!sh.getRange(FIRST_ROW-1, ID_COL).getValue()) sh.getRange(FIRST_ROW-1, ID_COL).setValue('ID App');

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
      sh.getRange(row, COL.tanggal).setValue(Utilities.parseDate(r['Tanggal'], tz, 'yyyy-MM-dd')).setNumberFormat('yyyy-mm-dd');
      sh.getRange(row, COL.nama).setValue(r['Nama Personil']);
      sh.getRange(row, COL.kegiatan).setValue(r['Kegiatan']);
      sh.getRange(row, COL.mulai).setValue(frac_(r['Jam Mulai'])).setNumberFormat('HH:mm');
      sh.getRange(row, COL.selesai).setValue(frac_(r['Jam Selesai'])).setNumberFormat('HH:mm');
      sh.getRange(row, COL.status).setValue(r['Status']);
      sh.getRange(row, COL.ket).setValue(r['Keterangan']);
      sh.getRange(row, ID_COL).setValue(req.id);
      return out_({ok:true, row:row});
    } finally { lock.releaseLock(); }
  }catch(err){ return out_({ok:false, error:String(err)}); }
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
function frac_(hhmm){
  const p = String(hhmm).split(':');
  return (Number(p[0])*60 + Number(p[1])) / 1440;
}
function out_(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
