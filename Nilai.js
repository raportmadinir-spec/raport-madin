// ==========================================
// FILE: Nilai.js
// FUNGSI: Modul Input Nilai UTS & UAS Santri
// ==========================================
function getInitNilai() {
  try {
    const sheetKelas = DB.getSheet('mst_kelas');
    const dataKelasRaw = sheetKelas.getDataRange().getDisplayValues();
    const headKelas = dataKelasRaw.length > 0 ? dataKelasRaw[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_')) : [];
    let kelasList = [];
    for(let i=1; i<dataKelasRaw.length; i++) {
      let obj = {};
      headKelas.forEach((h, idx) => obj[h] = dataKelasRaw[i][idx]);
      if((obj.deleted_at === "" || !obj.deleted_at) && (obj.is_active === 'TRUE' || obj.is_active === true)) {
        kelasList.push({ id_kelas: obj.id_kelas, nama_kelas: obj.nama_kelas, tingkat: obj.tingkat });
      }
    }
    const sheetMapel = DB.getSheet('mst_mapel');
    const dataMapelRaw = sheetMapel.getDataRange().getDisplayValues();
    const headMapel = dataMapelRaw.length > 0 ? dataMapelRaw[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_')) : [];
    let mapelList = [];
    for(let i=1; i<dataMapelRaw.length; i++) {
      let obj = {};
      headMapel.forEach((h, idx) => obj[h] = dataMapelRaw[i][idx]);
      if((obj.deleted_at === "" || !obj.deleted_at) && (obj.is_active === 'TRUE' || obj.is_active === true)) {
        mapelList.push({ id_mapel: obj.id_mapel, kode_mapel: obj.kode_mapel, nama_mapel: obj.nama_mapel });
      }
    }
    return { status: "success", kelas: kelasList, mapel: mapelList };
  } catch(e) { return { status: "error", message: e.toString() }; }
}

function getSiswaDanNilai(id_tapel, id_kelas, id_mapel) {
  try {
    // 1. Ambil Master Siswa (Untuk Nama & NIS)
    const sheetSiswa = DB.getSheet('mst_siswa');
    const dataSiswa = sheetSiswa.getDataRange().getDisplayValues();
    const headS = dataSiswa[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let mapSiswa = {};
    for(let i=1; i<dataSiswa.length; i++) {
      let idS = dataSiswa[i][headS.indexOf('id_siswa')];
      mapSiswa[idS] = {
        nis: dataSiswa[i][headS.indexOf('nis')],
        nama: dataSiswa[i][headS.indexOf('nama_lengkap')],
        jk: dataSiswa[i][headS.indexOf('jk')]
      };
    }
    // 2. Ambil Data Siswa di Kelas Tersebut (Dari Plotting)
    const sheetPlot = DB.getSheet('trx_kelas_siswa');
    let dataPlot = sheetPlot.getDataRange().getDisplayValues();
    const headP = dataPlot[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let siswaDiKelas = [];
    for(let i=1; i<dataPlot.length; i++) {
      if(dataPlot[i][headP.indexOf('id_tapel')].toString().trim() === id_tapel.toString().trim() && 
         dataPlot[i][headP.indexOf('id_kelas')].toString().trim() === id_kelas.toString().trim()) {
         siswaDiKelas.push(dataPlot[i][headP.indexOf('id_siswa')].toString().trim());
      }
    }
    if(siswaDiKelas.length === 0) {
      return { status: "error", message: "Belum ada santri yang di-plot ke kelas ini." };
    }
    // 3. Ambil Data Nilai yang Sudah Ada (Jika ada)
    const sheetNilai = DB.getSheet('trx_nilai');
    let allNilai = sheetNilai.getDataRange().getDisplayValues();
    
    // Auto-Repair Header trx_nilai
    if(allNilai.length === 0 || allNilai[0].length < 7) {
      sheetNilai.getRange(1, 1, 1, 7).setValues([['id_nilai', 'id_tapel', 'id_kelas', 'id_mapel', 'id_siswa', 'nilai_uts', 'nilai_uas']]);
      allNilai = sheetNilai.getDataRange().getDisplayValues();
    }
    const headN = allNilai[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let mapNilaiTersimpan = {};
    for(let i=1; i<allNilai.length; i++) {
      if(allNilai[i][headN.indexOf('id_tapel')].toString().trim() === id_tapel.toString().trim() &&
         allNilai[i][headN.indexOf('id_kelas')].toString().trim() === id_kelas.toString().trim() &&
         allNilai[i][headN.indexOf('id_mapel')].toString().trim() === id_mapel.toString().trim()) {
         
         mapNilaiTersimpan[allNilai[i][headN.indexOf('id_siswa')]] = {
           uts: allNilai[i][headN.indexOf('nilai_uts')],
           uas: allNilai[i][headN.indexOf('nilai_uas')]
         };
      }
    }
    // 4. Gabungkan Data (Santri + Nilai)
    let resultData = [];
    siswaDiKelas.forEach(idS => {
      if(mapSiswa[idS]) {
        let utsVal = mapNilaiTersimpan[idS] ? mapNilaiTersimpan[idS].uts : "";
        let uasVal = mapNilaiTersimpan[idS] ? mapNilaiTersimpan[idS].uas : "";
        resultData.push({
          id_siswa: idS,
          nis: mapSiswa[idS].nis,
          nama: mapSiswa[idS].nama,
          jk: mapSiswa[idS].jk,
          nilai_uts: utsVal,
          nilai_uas: uasVal
        });
      }
    });
    // Urutkan berdasarkan Nama Abjad
    resultData.sort((a, b) => a.nama.localeCompare(b.nama));
    return { status: "success", data: resultData };
  } catch(e) { return { status: "error", message: e.toString() }; }
}

function saveNilaiKolektif(id_tapel, id_kelas, id_mapel, arr_nilai) {
  try {
    const sheet = DB.getSheet('trx_nilai');
    let allData = sheet.getDataRange().getValues();
    
    if (allData.length === 0 || allData[0].length < 7) {
      sheet.getRange(1, 1, 1, 7).setValues([['id_nilai', 'id_tapel', 'id_kelas', 'id_mapel', 'id_siswa', 'nilai_uts', 'nilai_uas']]);
      allData = sheet.getDataRange().getValues();
    }
    let headers = allData[0].map(h => h.toString().toLowerCase().trim().replace(/\s+/g, '_'));
    let idxId = headers.indexOf('id_nilai');
    let idxTapel = headers.indexOf('id_tapel');
    let idxKelas = headers.indexOf('id_kelas');
    let idxMapel = headers.indexOf('id_mapel');
    let idxSiswa = headers.indexOf('id_siswa');
    let idxUts = headers.indexOf('nilai_uts');
    let idxUas = headers.indexOf('nilai_uas');
    
    // HAPUS NILAI LAMA UNTUK KOMBINASI INI (WIPE)
    for (let i = allData.length - 1; i > 0; i--) {
      if (allData[i][idxTapel].toString().trim() === id_tapel.toString().trim() && 
          allData[i][idxKelas].toString().trim() === id_kelas.toString().trim() &&
          allData[i][idxMapel].toString().trim() === id_mapel.toString().trim()) {
        sheet.deleteRow(i + 1);
      }
    }
    // INSERT NILAI BARU (Hanya yang diisi)
    if (arr_nilai && arr_nilai.length > 0) {
      let newRows = [];
      arr_nilai.forEach(item => {
        // Simpan meskipun nilainya 0. Hanya abaikan jika benar-benar kosong string.
        if(item.uts !== "" || item.uas !== "") { 
          let row = new Array(headers.length).fill("");
          row[idxId] = Utilities.getUuid();
          row[idxTapel] = id_tapel;
          row[idxKelas] = id_kelas;
          row[idxMapel] = id_mapel;
          row[idxSiswa] = item.id_siswa;
          row[idxUts] = item.uts;
          row[idxUas] = item.uas;
          newRows.push(row);
        }
      });
      
      if(newRows.length > 0) {
        sheet.getRange(sheet.getLastRow() + 1, 1, newRows.length, headers.length).setValues(newRows);
      }
    }
    
    return { status: "success", message: "Nilai UTS & UAS berhasil disimpan!" };
  } catch(e) { return { status: "error", message: e.toString() }; }
}