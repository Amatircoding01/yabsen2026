const SHIFT_START = "07:30:00"; 

function doGet() {
  setupDatabase(); 
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Dashboard HRIS KOBE')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName('Master_Karyawan')) {
    let sheet = ss.insertSheet('Master_Karyawan');
    sheet.appendRow(['NRPP', 'Nama Karyawan', 'Gol', 'Status', 'Jabatan', 'dept', 'Divisi', 'Perusahaan']);
    sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#d9ead3");
  }
  if (!ss.getSheetByName('Raw_Kehadiran')) {
    let sheet = ss.insertSheet('Raw_Kehadiran');
    sheet.appendRow(['Tanggal', 'NRPP', 'Nama', 'Jam Masuk', 'Jam Keluar', 'Status']);
    sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#c9daf8");
  }
  if (!ss.getSheetByName('Rekap_Bulanan')) {
    let sheet = ss.insertSheet('Rekap_Bulanan');
    sheet.appendRow(['Bulan-Tahun', 'NRPP', 'Nama', 'Divisi', 'H (Hadir)', 'S (Sakit)', 'C (Cuti)', 'PG (Potong Gaji/Alpa)', 'ST', 'Visit Customer']);
    sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#fff2cc");
  }
}

function convertSmartDate(value) {
  if (!value) return value;
  if (typeof value === 'string' && (value.includes('/') || value.includes('-'))) return value;
  let serial = parseFloat(value);
  if (!isNaN(serial)) {
    const excelDate = new Date((serial - 25569) * 86400 * 1000); 
    const day = ("0" + excelDate.getDate()).slice(-2);
    const month = ("0" + (excelDate.getMonth() + 1)).slice(-2);
    const year = excelDate.getFullYear();
    return `${day}/${month}/${year}`;
  }
  return value;
}

function timeToSeconds(timeStr) {
  if(!timeStr || typeof timeStr !== 'string') return 0;
  let parts = timeStr.trim().split(':');
  if(parts.length < 2) return 0;
  let h = parseInt(parts[0]) || 0; let m = parseInt(parts[1]) || 0; let s = parseInt(parts[2]) || 0;
  return h * 3600 + m * 60 + s;
}

function secondsToTimeStr(secs) {
  if (secs <= 0) return "00:00:00";
  let h = Math.floor(secs / 3600);
  let m = Math.floor((secs % 3600) / 60);
  let s = secs % 60;
  return `${h.toString().padStart(2,'0')}:${m.toString().padStart(2,'0')}:${s.toString().padStart(2,'0')}`;
}

function getWorkdaysByDate(startStr, endStr) {
  if(!startStr || !endStr) return 0;
  let sParts = startStr.split('-'); let eParts = endStr.split('-');
  let d1 = new Date(parseInt(sParts[0]), parseInt(sParts[1])-1, parseInt(sParts[2]));
  let d2 = new Date(parseInt(eParts[0]), parseInt(eParts[1])-1, parseInt(eParts[2]));
  let count = 0; let cur = new Date(d1);
  while(cur <= d2) {
    let wd = cur.getDay();
    if(wd !== 0 && wd !== 6) count++;
    cur.setDate(cur.getDate() + 1);
  }
  return count;
}

function isDateInRange(tglSheet, startStr, endStr) {
  if(!tglSheet || !startStr || !endStr) return false;
  let parts = tglSheet.toString().split('/');
  if(parts.length !== 3) return false;
  let d = parts[0].padStart(2, '0'); let m = parts[1].padStart(2, '0'); let y = parts[2];
  let yyyyMmDd = `${y}-${m}-${d}`;
  return (yyyyMmDd >= startStr && yyyyMmDd <= endStr);
}

function processExcelUpload(dataRows) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
    const sheetMaster = ss.getSheetByName('Master_Karyawan');
    
    // 1. Ambil semua NRPP dari Master_Karyawan untuk filter
    const masterData = sheetMaster.getDataRange().getValues();
    const daftarNrppMaster = masterData.map(row => row[0].toString().trim());

    let dataSiapTulis = [];
    let jumlahDiskip = 0;

    // 2. Iterasi data dari Excel (mulai baris ke-2 jika ada header)
    for (let i = 1; i < dataRows.length; i++) {
      let row = dataRows[i];
      if (!row || row.length < 2 || !row[0]) continue; 

      let nrppExcel = row[0].toString().trim();
      
      // --- PROSES FILTER DISINI ---
      if (!daftarNrppMaster.includes(nrppExcel)) {
        jumlahDiskip++; // Lewati jika NRPP tidak ada di Master_Karyawan
        continue; 
      }
      // ----------------------------

      let nama = row[1] ? row[1].toString().trim() : "";
      let signIn = row[2] ? row[2].toString().trim() : "";
      let signOut = row[3] ? row[3].toString().trim() : "";
      let status = row[4] ? row[4].toString().trim() : "H";
      let rawDate = row[5] ? row[5] : "";
      let tanggalRapih = convertSmartDate(rawDate);

      dataSiapTulis.push([tanggalRapih, nrppExcel, nama, signIn, signOut, status]);
    }

    if (dataSiapTulis.length > 0) {
      sheetRaw.getRange(sheetRaw.getLastRow() + 1, 1, dataSiapTulis.length, 6).setValues(dataSiapTulis);
      updateRekapanBulanan(dataSiapTulis); 
      
      let msg = `${dataSiapTulis.length} data berhasil masuk.`;
      if(jumlahDiskip > 0) msg += ` (${jumlahDiskip} data karyawan lain otomatis difilter/dibuang).`;
      
      return { success: true, message: msg };
    } else { 
      return { success: false, message: "Tidak ada data yang cocok dengan Master Karyawan." }; 
    }
  } catch (error) { 
    return { success: false, message: error.message }; 
  }
}

function updateRekapanBulanan(dataBaru) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetRekap = ss.getSheetByName('Rekap_Bulanan');
  const sheetMaster = ss.getSheetByName('Master_Karyawan');
  
  let masterValues = sheetMaster.getDataRange().getDisplayValues();
  let kamusDivisi = {};
  for(let i=1; i<masterValues.length; i++) { 
    if(masterValues[i] && masterValues[i][0]) { kamusDivisi[masterValues[i][0].toString().trim()] = masterValues[i][6] || "-"; }
  }
  
  let hitungHadir = {};
  dataBaru.forEach(row => {
    let tgl = row[0]; if(!tgl || typeof tgl !== 'string') return;
    let bulanTahun = tgl.substring(3, 10); 
    let nrpp = row[1]; let nama = row[2];
    let key = bulanTahun + "_" + nrpp;
    if(!hitungHadir[key]) { hitungHadir[key] = { bulan: bulanTahun, nrpp: nrpp, nama: nama, divisi: (kamusDivisi[nrpp]||"-"), hadir: 0 }; }
    if(row[5] === 'H') hitungHadir[key].hadir += 1;
  });
  
  let barisBaru = [];
  for (let key in hitungHadir) {
    let d = hitungHadir[key];
    barisBaru.push([d.bulan, d.nrpp, d.nama, d.divisi, d.hadir, 0, 0, 0, 0, 0]); 
  }
  if(barisBaru.length > 0) sheetRekap.getRange(sheetRekap.getLastRow() + 1, 1, barisBaru.length, 10).setValues(barisBaru);
}

function getDaftarKaryawan() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master_Karyawan');
    if(!sheet) return { error: "Sheet Master_Karyawan tidak ditemukan." };
    const data = sheet.getDataRange().getDisplayValues();
    let hasil = []; let ptList = [];
    for(let i=1; i<data.length; i++) {
      let row = data[i];
      if(row && row[0]) {
        let pt = (row[7] || "").toString().trim();
        hasil.push({nrpp: row[0].toString().trim(), nama: row[1] || "-", pt: pt});
        if(pt && !ptList.includes(pt)) ptList.push(pt);
      }
    }
    if(ptList.length === 0) ptList.push("PT. KOBE BOGA UTAMA");
    return { success: true, data: hasil, perusahaanList: ptList };
  } catch (err) { return { error: err.message }; }
}

function getAttendanceForEdit(dateStr, ptFilter) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const masterData = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
  let employees = [];
  for(let i=1; i<masterData.length; i++){
    let row = masterData[i];
    if(!row || !row[0]) continue;
    let ptEmp = (row[7] || "").toString().trim();
    if(ptFilter === "ALL" || ptEmp === ptFilter) {
      employees.push({nrpp: row[0].toString().trim(), nama: row[1]});
    }
  }
  
  let parts = dateStr.split('-');
  let tglFormatted = `${parts[2]}/${parts[1]}/${parts[0]}`;
  const rawData = ss.getSheetByName('Raw_Kehadiran').getDataRange().getDisplayValues();
  
  let attMap = {};
  for(let i=1; i<rawData.length; i++){
    if(rawData[i][0] === tglFormatted) {
       attMap[rawData[i][1].toString().trim()] = { masuk: rawData[i][3], keluar: rawData[i][4], status: rawData[i][5] };
    }
  }
  
  employees.forEach(emp => {
    let att = attMap[emp.nrpp];
    emp.masuk = att ? att.masuk : "";
    emp.keluar = att ? att.keluar : "";
    emp.status = att ? att.status : "-";
  });
  return employees;
}

function saveAbsensiManual(tglInput, nrpp, masuk, keluar, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
    const rawData = sheetRaw.getDataRange().getDisplayValues();
    let parts = tglInput.split('-');
    let tglFormatted = `${parts[2]}/${parts[1]}/${parts[0]}`;

    const masterData = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
    let nama = "-";
    for(let i=1; i<masterData.length; i++) {
      let row = masterData[i];
      if(row && row[0] && row[0].toString().trim() === nrpp.toString().trim()) { nama = row[1] || "-"; break; }
    }

    let foundRowIndex = -1;
    for(let i=1; i<rawData.length; i++) {
      let row = rawData[i];
      if(row && row[0] === tglFormatted && row[1].toString().trim() === nrpp.toString().trim()) {
        foundRowIndex = i + 1; break;
      }
    }

    if(foundRowIndex > -1) {
      sheetRaw.getRange(foundRowIndex, 4).setValue(masuk);
      sheetRaw.getRange(foundRowIndex, 5).setValue(keluar);
      sheetRaw.getRange(foundRowIndex, 6).setValue(status);
    } else {
      sheetRaw.appendRow([tglFormatted, nrpp, nama, masuk, keluar, status]);
    }
    return { success: true, message: `Disimpan` };
  } catch(e) { return { success: false, message: e.message }; }
}

function deleteAbsensiManual(tglInput, nrpp) {
  try {
    const sheetRaw = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Raw_Kehadiran');
    const rawData = sheetRaw.getDataRange().getDisplayValues();
    let parts = tglInput.split('-');
    let tglFormatted = `${parts[2]}/${parts[1]}/${parts[0]}`;
    
    let foundRowIndex = -1;
    for(let i=1; i<rawData.length; i++) {
      if(rawData[i][0] === tglFormatted && rawData[i][1].toString().trim() === nrpp.toString().trim()) {
        foundRowIndex = i + 1; break;
      }
    }
    if(foundRowIndex > -1) {
      sheetRaw.deleteRow(foundRowIndex);
      return { success: true, message: "Dihapus" };
    }
    return { success: false, message: "Tidak ada data" };
  } catch(e) { return { success: false, message: e.message }; }
}

// -------------------------------------------------------------------------
// PERUBAHAN CROSSTAB: Tambah Summary Kalkulasi & Waktu Keterlambatan
// -------------------------------------------------------------------------
function getCrosstabRekapData(startStr, endStr, perusahaan) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let d1 = new Date(startStr); let d2 = new Date(endStr);
  let datesArr = [];
  const namaHari = ["Minggu","Senin","Selasa","Rabu","Kamis","Jumat","Sabtu"];
  const namaBulan = ["Januari","Februari","Maret","April","Mei","Juni","Juli","Agustus","September","Oktober","November","Desember"];

  let cur = new Date(d1);
  while(cur <= d2) {
    let y = cur.getFullYear(); let m = (cur.getMonth()+1).toString().padStart(2, '0'); let d = cur.getDate().toString().padStart(2, '0');
    let rawDate = `${d}/${m}/${y}`; 
    let displayDate = `${namaHari[cur.getDay()]}, ${d} ${namaBulan[cur.getMonth()]} ${y}`;
    datesArr.push({ raw: rawDate, display: displayDate });
    cur.setDate(cur.getDate() + 1);
  }

  const masterData = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
  let employeesMap = {};
  let employees = [];
  for(let i=1; i<masterData.length; i++) {
    let row = masterData[i];
    if(!row || !row[0]) continue;
    let pt = (row[7] || "").toString().trim(); 
    if(perusahaan === "ALL" || pt === perusahaan) {
      let nrpp = row[0].toString().trim();
      let empObj = { 
        nrpp: nrpp, nama: row[1] || "-", 
        H: 0, S: 0, C: 0, ST: 0, VC: 0, PG: 0, lateSec: 0 
      };
      employees.push(empObj);
      employeesMap[nrpp] = empObj;
    }
  }

  const rawKehadiran = ss.getSheetByName('Raw_Kehadiran').getDataRange().getDisplayValues();
  let attMap = {}; 
  let stdStartSec = timeToSeconds(SHIFT_START); 

  for(let i=1; i<rawKehadiran.length; i++) {
    let row = rawKehadiran[i]; 
    if(!row || !row[0] || !row[1]) continue;
    let tgl = row[0].toString().trim(); 
    let nrpp = row[1].toString().trim();
    let masuk = row[3] || "";
    let keluar = row[4] || "";
    let status = (row[5] || "H").toString().trim().toUpperCase();

    // Matriks data
    if(!attMap[nrpp]) attMap[nrpp] = {};
    attMap[nrpp][tgl] = { in: masuk||"-", out: keluar||"-", stat: status||"-" };

    // Agregasi Summary di Crosstab
    if (isDateInRange(tgl, startStr, endStr) && employeesMap[nrpp]) {
        if(status === 'H') employeesMap[nrpp].H++;
        else if(status === 'S') employeesMap[nrpp].S++;
        else if(status === 'C') employeesMap[nrpp].C++;
        else if(status === 'ST') employeesMap[nrpp].ST++;
        else if(status === 'PG') employeesMap[nrpp].PG++;
        else if(status === 'VC') employeesMap[nrpp].VC++;

        if (masuk) {
           let mSec = timeToSeconds(masuk);
           if (mSec > stdStartSec) {
               employeesMap[nrpp].lateSec += (mSec - stdStartSec);
           }
        }
    }
  }

  // Format detik menjadi String HH:MM:SS
  employees.forEach(emp => {
     emp.lateStr = secondsToTimeStr(emp.lateSec);
  });

  let totalWorkdays = getWorkdaysByDate(startStr, endStr);

  return { dates: datesArr, employees: employees, attendance: attMap, workdays: totalWorkdays };
}

function getMultiRekapSummary(nrppArray, startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const raw = ss.getSheetByName('Raw_Kehadiran').getDataRange().getDisplayValues();
  const masterData = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
  
  let masterDict = {};
  for(let i=1; i<masterData.length; i++) { 
    let row = masterData[i];
    if(row && row[0]) masterDict[row[0].toString().trim()] = { nama: row[1]||"", div: row[6]||"-" }; 
  }
  
  let result = {};
  let stdStartSec = timeToSeconds(SHIFT_START); 
  
  for(let i=1; i<raw.length; i++) {
    let row = raw[i];
    if(!row || !row[0] || !row[1]) continue;
    let tgl = row[0].toString().trim(); let nrpp = row[1].toString().trim(); let status = (row[5]||"H").toString().trim().toUpperCase();
    
    if (isDateInRange(tgl, startDateStr, endDateStr) && nrppArray.includes(nrpp)) {
      if(!result[nrpp]) { 
        result[nrpp] = { nrpp: nrpp, nama: (masterDict[nrpp]?.nama||row[2]||"-"), div: (masterDict[nrpp]?.div||"-"), H:0, S:0, C:0, PG:0, totalLate: 0 }; 
      }
      if(status==='H') result[nrpp].H++;
      else if(status==='S') result[nrpp].S++;
      else if(status==='C') result[nrpp].C++;
      else if(status==='PG') result[nrpp].PG++;
      
      let masuk = row[3];
      if(masuk) {
        let mSec = timeToSeconds(masuk);
        if(mSec > stdStartSec) {
          result[nrpp].totalLate += (mSec - stdStartSec);
        }
      }
    }
  }
  
  return Object.values(result).map(r => [
    `${startDateStr} s/d ${endDateStr}`, r.nrpp, r.nama, r.div, r.H, r.S, r.C, r.PG, secondsToTimeStr(r.totalLate)
  ]);
}

function getLaporanIndividuDynamicByDate(nrpp, startDateStr, endDateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const masterData = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
  let profil = { nrpp: nrpp, nama: "-", gol: "-", status: "-", jabatan: "-", dept: "-", divisi: "-", pt: "" };
  for(let i=1; i<masterData.length; i++) {
    let row = masterData[i];
    if(!row || !row[0]) continue;
    if(row[0].toString().trim() === nrpp.toString().trim()) {
      profil.nama = row[1] || "-"; profil.gol = row[2] || "-";
      profil.status = row[3] || "-"; profil.jabatan = row[4] || "-";
      profil.dept = row[5] || "-"; profil.divisi = row[6] || "-";
      profil.pt = row[7] || "";
      break;
    }
  }

  const sheetRaw = ss.getSheetByName('Raw_Kehadiran');
  let history = [];
  const namaHariList = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];

  if (sheetRaw) {
    const rawData = sheetRaw.getDataRange().getDisplayValues();
    for (let i=1; i<rawData.length; i++) {
       let row = rawData[i];
       if (!row || !row[0] || !row[1]) continue;
       if (row[1].toString().trim() !== nrpp.toString().trim()) continue;
       if (!isDateInRange(row[0], startDateStr, endDateStr)) continue;
       
       let masukTeks = (row[3] || "").toString().trim(); 
       let keluarTeks = (row[4] || "").toString().trim();
       let stat = (row[5] || "H").toString().trim().toUpperCase();
       
       let parts = row[0].split('/');
       if(parts.length !== 3) continue;

       let dObj = new Date(parseInt(parts[2]), parseInt(parts[1])-1, parseInt(parts[0]));
       let dayOfWeek = dObj.getDay(); 
       let namaHari = namaHariList[dayOfWeek];
       
       let stdStartSec = timeToSeconds(SHIFT_START); 
       let mSec = timeToSeconds(masukTeks);
       let kSec = timeToSeconds(keluarTeks);
       
       let strTerlambat = "-"; let strDurasi = "-";

       if (masukTeks && mSec > 0) {
           let terlambatSec = (mSec > stdStartSec) ? (mSec - stdStartSec) : 0;
           strTerlambat = secondsToTimeStr(terlambatSec); 
       }
       if (masukTeks && keluarTeks && kSec > 0 && mSec > 0) {
           let durasiSec = (kSec > mSec) ? (kSec - mSec) : 0;
           strDurasi = secondsToTimeStr(durasiSec);
       }

       history.push({ 
         hari: namaHari, tgl: row[0], masuk: masukTeks || '-', keluar: keluarTeks || '-', 
         terlambat: strTerlambat, durasi: strDurasi, status: stat 
       });
    }
  }

  let summary = { H: 0, S: 0, C: 0, ST: 0, PG: 0, VC: 0, TotalHariKerja: 0 };
  summary.H = history.filter(r => r.status === 'H').length;
  summary.S = history.filter(r => r.status === 'S').length;
  summary.C = history.filter(r => r.status === 'C').length;
  summary.ST = history.filter(r => r.status === 'ST').length;
  summary.PG = history.filter(r => r.status === 'PG').length;
  summary.VC = history.filter(r => r.status === 'VC').length;

  let sParts = startDateStr.split('-');
  let mmYYYY = sParts[1] + '/' + sParts[0];

  const sheetRekap = ss.getSheetByName('Rekap_Bulanan');
  if (sheetRekap) {
    const rekapData = sheetRekap.getDataRange().getDisplayValues();
    for(let i=1; i<rekapData.length; i++) {
      let row = rekapData[i];
      if(!row || !row[0] || !row[1]) continue;
      if(row[0] === mmYYYY && row[1].toString().trim() === nrpp.toString().trim()) {
        summary.S = Math.max(summary.S, parseInt(row[5]) || 0);
        summary.C = Math.max(summary.C, parseInt(row[6]) || 0);
        summary.PG = Math.max(summary.PG, parseInt(row[7]) || 0);
        summary.ST = Math.max(summary.ST, parseInt(row[8]) || 0);
        summary.VC = Math.max(summary.VC, parseInt(row[9]) || 0);
        break;
      }
    }
  }
  
  summary.TotalHariKerja = getWorkdaysByDate(startDateStr, endDateStr);
  return { profil: profil, data: history, summary: summary };
}

function getVariabelGajiDataByDate(startDateStr, endDateStr, filterNrpp, filterPt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let result = {};
  const masterData = ss.getSheetByName('Master_Karyawan').getDataRange().getDisplayValues();
  let masterDict = {};
  for(let i=1; i<masterData.length; i++) { 
    let row = masterData[i];
    if(!row || !row[0]) continue;
    masterDict[row[0].toString().trim()] = { nama: row[1]||"", dept: row[5]||"-", pt: row[7]||"" }; 
  }

  const raw = ss.getSheetByName('Raw_Kehadiran').getDataRange().getDisplayValues();
  for(let i=1; i<raw.length; i++) {
    let row = raw[i];
    if(!row || !row[0] || !row[1]) continue;
    let tgl = row[0].toString().trim(); let nrpp = row[1].toString().trim(); let status = (row[5]||"H").toString().trim().toUpperCase();
    
    if (isDateInRange(tgl, startDateStr, endDateStr)) {
      if (filterNrpp && filterNrpp !== 'ALL' && nrpp !== filterNrpp.trim()) continue;
      let ptKaryawan = masterDict[nrpp]?.pt || "";
      if (filterPt && filterPt !== 'ALL' && ptKaryawan !== filterPt.trim()) continue;
      
      if(!result[nrpp]) { result[nrpp] = { nrpp: nrpp, nama: (masterDict[nrpp]?.nama || row[2] || "-"), dept: (masterDict[nrpp]?.dept || "-"), totalHadir: 0 }; }
      if(status === 'H') result[nrpp].totalHadir++;
    }
  }
  return Object.values(result).map(r => [r.nrpp, r.nama, r.dept, r.totalHadir]);
}
