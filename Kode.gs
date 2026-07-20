const SPREADSHEET_ID = '1f1vepNb9QgsglpR_ZqESjhrJcP49KJxRLkSR3uTad5c';
const SHEET_USERS = 'Users';
const SHEET_KELAS = 'Kelas';
const SHEET_SISWA = 'Siswa';
const SHEET_JADWAL = 'Jam_Mengajar';
const SHEET_ABSENSI = 'Absensi';
const SHEET_PENILAIAN = 'Penilaian';
const SHEET_INFORMASI = 'Informasi';
const SHEET_JURNAL = 'Jurnal_Mengajar';

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('SakuGuru - Portal Administrasi Sekolah')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getSheet(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    
    if (sheetName === SHEET_USERS) {
      sheet.appendRow(['NIP', 'Nama', 'Password', 'Role', 'Data_Mengajar']);
      sheet.appendRow(['admin', 'Administrator Sekolah', '@Smaga123', 'admin', '']);
      sheet.appendRow(['guru', 'Budi Santoso, S.Pd', '@Smaga123', 'guru', '']);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f4f6");
    }
    else if (sheetName === SHEET_KELAS) {
      sheet.appendRow(['ID_Kelas', 'Nama_Kelas', 'Tingkat', 'Nama_Wali_Kelas', 'NIP_Wali_Kelas']);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f4f6");
    }
    else if (sheetName === SHEET_SISWA) {
      sheet.appendRow(['NISN', 'Nama_Siswa', 'Jenis_Kelamin', 'ID_Kelas', 'No_HP_Ortu']);
      sheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f4f6");
    }
    else if (sheetName === SHEET_JADWAL) {
      sheet.appendRow(['ID_Jadwal', 'NIP_Guru', 'Nama_Guru', 'Tanggal', 'Jam_Ke', 'Waktu_Mengajar', 'Mata_Pelajaran']);
      sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#f3f4f6");
    }
    else if (sheetName === SHEET_ABSENSI) {
      sheet.appendRow(['ID_Absensi', 'Tanggal', 'Jam_Ke', 'ID_Kelas', 'Mata_Pelajaran', 'NIP_Guru', 'Nama_Guru', 'Data_Kehadiran']);
      sheet.getRange("A1:H1").setFontWeight("bold").setBackground("#f3f4f6");
    }
    else if (sheetName === SHEET_PENILAIAN) {
      sheet.appendRow(['ID_Penilaian', 'Tanggal', 'Jam_Ke', 'ID_Kelas', 'Mata_Pelajaran', 'Materi_Ajar', 'Jenis_Penilaian', 'NIP_Guru', 'Nama_Guru', 'Data_Nilai']);
      sheet.getRange("A1:J1").setFontWeight("bold").setBackground("#f3f4f6");
    }
    else if (sheetName === SHEET_INFORMASI) {
      sheet.appendRow(['ID_Info', 'Tanggal', 'Isi_Informasi', 'NIP_Pembuat']);
      sheet.appendRow(['INF-1', new Date().toISOString().split('T')[0], 'Pastikan pengisian Absensi, Penilaian, dan Jurnal Mengajar dilakukan secara berkala.', 'admin']);
      sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#f3f4f6");
    }
    else if (sheetName === SHEET_JURNAL) {
      sheet.appendRow(['ID_Jurnal', 'Tanggal', 'Jam_Ke', 'ID_Kelas', 'Mata_Pelajaran', 'Materi_Pembelajaran', 'Evaluasi_Catatan', 'NIP_Guru', 'Nama_Guru']);
      sheet.getRange("A1:I1").setFontWeight("bold").setBackground("#f3f4f6");
    }
  }
  return sheet;
}

function verifyLogin(nip, password) {
  try {
    const sheet = getSheet(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      let dbNip = data[i][0].toString().trim();
      let dbName = data[i][1].toString().trim();
      let dbPass = data[i][2].toString().trim();
      let dbRole = data[i][3].toString().trim().toLowerCase();
      let dbDataMengajar = data[i][4] ? data[i][4].toString() : '[]';

      if (dbNip.toLowerCase() === nip.toString().trim().toLowerCase() && dbPass === password.toString().trim()) {
        return { success: true, user: { nip: dbNip, name: dbName, role: dbRole, dataMengajar: dbDataMengajar } };
      }
    }
    return { success: false, message: 'NIP tidak terdaftar atau Kata Sandi salah.' };
  } catch (error) {
    return { success: false, message: 'Gagal membaca Database: ' + error.message };
  }
}

function getTeachers() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const teachersMap = {};

    // 1. Ambil guru terdaftar dari sheet Users
    const userSheet = ss.getSheetByName(SHEET_USERS);
    const userData = userSheet ? userSheet.getDataRange().getValues() : [];
    for (let i = 1; i < userData.length; i++) {
      let dbNip = userData[i][0].toString().trim();
      let dbName = userData[i][1].toString().trim();
      let dbRole = userData[i][3].toString().trim().toLowerCase();
      let dbDataMengajar = userData[i][4] ? userData[i][4].toString() : '[]';

      if (dbRole === 'guru' && dbNip !== '') {
        teachersMap[dbNip.toLowerCase()] = {
          nip: dbNip,
          name: dbName,
          role: 'guru',
          dataMengajar: dbDataMengajar
        };
      }
    }

    // 2. Ambil guru yang ditugaskan sebagai Wali Kelas di sheet Kelas
    const kelasSheet = ss.getSheetByName(SHEET_KELAS);
    const kelasData = kelasSheet ? kelasSheet.getDataRange().getValues() : [];
    for (let i = 1; i < kelasData.length; i++) {
      let classId = kelasData[i][0].toString().trim();
      let waliName = kelasData[i][3].toString().trim();
      let waliNip = kelasData[i][4].toString().trim();

      if (waliNip !== '' && waliName !== '') {
        const key = waliNip.toLowerCase();
        if (!teachersMap[key]) {
          // Buat profil guru sementara dengan tugas mengajar otomatis sebagai wali kelas
          const autoTeaching = JSON.stringify([{ mapel: 'Wali Kelas', kelas: [classId] }]);
          teachersMap[key] = {
            nip: waliNip,
            name: waliName,
            role: 'guru',
            dataMengajar: autoTeaching
          };
        }
      }
    }

    // 3. Ambil guru dari daftar jadwal mengajar aktif
    const jadwalSheet = ss.getSheetByName(SHEET_JADWAL);
    const jadwalData = jadwalSheet ? jadwalSheet.getDataRange().getValues() : [];
    for (let i = 1; i < jadwalData.length; i++) {
      let nip = jadwalData[i][1].toString().trim();
      let name = jadwalData[i][2].toString().trim();

      if (nip !== '' && name !== '') {
        const key = nip.toLowerCase();
        if (!teachersMap[key]) {
          teachersMap[key] = {
            nip: nip,
            name: name,
            role: 'guru',
            dataMengajar: '[]'
          };
        }
      }
    }

    const result = Object.values(teachersMap);
    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveTeacherOnboarding(nip, teachingData) {
  try {
    const sheet = getSheet(SHEET_USERS);
    const data = sheet.getDataRange().getValues();
    const jsonString = JSON.stringify(teachingData);
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === nip.toString().trim().toLowerCase()) {
        sheet.getRange(i + 1, 5).setValue(jsonString); 
        return { success: true, message: 'Data tugas mengajar guru berhasil diperbarui.' };
      }
    }
    return { success: false, message: 'User Guru tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan: ' + error.message };
  }
}

function getGlobalPrintSettings() {
  try {
    const props = PropertiesService.getScriptProperties();
    return {
      success: true,
      data: {
        namaKepsek: props.getProperty('PRINT_NAMA_KEPSEK') || 'Dr. H. Ahmad Yani, M.Pd.',
        nipKepsek: props.getProperty('PRINT_NIP_KEPSEK') || '197508122000031002',
        marginTop: props.getProperty('PRINT_MARGIN_TOP') || '20',
        marginBottom: props.getProperty('PRINT_MARGIN_BOTTOM') || '20',
        marginLeft: props.getProperty('PRINT_MARGIN_LEFT') || '20',
        marginRight: props.getProperty('PRINT_MARGIN_RIGHT') || '20'
      }
    };
  } catch(e) {
    return { 
      success: true, 
      data: { 
        namaKepsek: 'Dr. H. Ahmad Yani, M.Pd.', 
        nipKepsek: '197508122000031002', 
        marginTop: '20', 
        marginBottom: '20', 
        marginLeft: '20', 
        marginRight: '20' 
      } 
    };
  }
}

function saveGlobalPrintSettings(settings) {
  try {
    const props = PropertiesService.getScriptProperties();
    props.setProperty('PRINT_NAMA_KEPSEK', settings.namaKepsek || '');
    props.setProperty('PRINT_NIP_KEPSEK', settings.nipKepsek || '');
    props.setProperty('PRINT_MARGIN_TOP', settings.marginTop || '20');
    props.setProperty('PRINT_MARGIN_BOTTOM', settings.marginBottom || '20');
    props.setProperty('PRINT_MARGIN_LEFT', settings.marginLeft || '20');
    props.setProperty('PRINT_MARGIN_RIGHT', settings.marginRight || '20');
    return { success: true, message: 'Pengaturan cetak instansi global berhasil disimpan.' };
  } catch(e) {
    return { success: false, message: e.message };
  }
}

function getClasses() {
  try {
    const sheet = getSheet(SHEET_KELAS);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if(data[i][0] !== '') { 
        result.push({
          idKelas: data[i][0].toString().trim(),
          namaKelas: data[i][1].toString().trim(),
          tingkat: data[i][2].toString().trim(),
          namaWali: data[i][3].toString().trim(),
          nipWali: data[i][4].toString().trim()
        });
      }
    }
    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveClass(classData, action, oldId) {
  try {
    const sheet = getSheet(SHEET_KELAS);
    const data = sheet.getDataRange().getValues();
    const rowData = [classData.idKelas, classData.namaKelas, classData.tingkat, classData.namaWali, classData.nipWali];

    if (action === 'add') {
      const isExist = data.some(row => row[0].toString().trim().toLowerCase() === classData.idKelas.toLowerCase());
      if (isExist) return { success: false, message: `Kelas ${classData.namaKelas} sudah ada!` };
      sheet.appendRow(rowData);
      return { success: true, message: `Kelas berhasil ditambahkan.` };
    } else if (action === 'edit') {
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString().trim().toLowerCase() === oldId.trim().toLowerCase()) { rowIndex = i + 1; break; }
      }
      if (rowIndex > -1) {
        if (oldId.toLowerCase() !== classData.idKelas.toLowerCase() && data.some(row => row[0].toString().trim().toLowerCase() === classData.idKelas.toLowerCase())) {
          return { success: false, message: `Kelas ${classData.namaKelas} sudah dipakai!` };
        }
        sheet.getRange(rowIndex, 1, 1, 5).setValues([rowData]);
        return { success: true, message: `Kelas berhasil diperbarui.` };
      }
      return { success: false, message: 'Data kelas tidak ditemukan.' };
    }
  } catch (error) { return { success: false, message: 'Gagal menyimpan: ' + error.message }; }
}

function deleteClass(classId) {
  try {
    const sheet = getSheet(SHEET_KELAS);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().trim().toLowerCase() === classId.trim().toLowerCase()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: `Kelas berhasil dihapus.` };
      }
    }
    return { success: false, message: 'Kelas tidak ditemukan.' };
  } catch (error) { return { success: false, message: 'Gagal menghapus kelas: ' + error.message }; }
}

function getStudents() {
  try {
    const sheet = getSheet(SHEET_SISWA);
    const data = sheet.getDataRange().getValues();
    const result = [];
    const singleQuoteRegex = new RegExp("'", "g");
    for (let i = 1; i < data.length; i++) {
      if(data[i][0] !== '') { 
        result.push({
          nisn: data[i][0].toString().replace(singleQuoteRegex, "").trim(),
          nama: data[i][1].toString().trim(),
          jk: data[i][2].toString().trim(),
          idKelas: data[i][3].toString().trim(),
          noHp: data[i][4].toString().replace(singleQuoteRegex, "").trim()
        });
      }
    }
    return { success: true, data: result };
  } catch (error) { return { success: false, message: error.message }; }
}

function saveStudent(studentData, action, oldNisn) {
  try {
    const sheet = getSheet(SHEET_SISWA);
    const data = sheet.getDataRange().getValues();
    const singleQuoteRegex = new RegExp("'", "g");

    const formatNoHp = studentData.noHp.startsWith("'") ? studentData.noHp : "'" + studentData.noHp;
    const formatNisn = studentData.nisn.startsWith("'") ? studentData.nisn : "'" + studentData.nisn;
    const rowData = [formatNisn, studentData.nama, studentData.jk, studentData.idKelas, formatNoHp];

    const cleanNisn = studentData.nisn.replace(singleQuoteRegex, "").trim();
    const cleanOldNisn = oldNisn.replace(singleQuoteRegex, "").trim();

    if (action === 'add') {
      if (data.some(row => row[0].toString().replace(singleQuoteRegex, "").trim() === cleanNisn)) {
        return { success: false, message: `Siswa dengan NISN ${cleanNisn} sudah terdaftar!` };
      }
      sheet.appendRow(rowData);
      return { success: true, message: `Siswa berhasil didaftarkan.` };
    } else if (action === 'edit') {
      let rowIndex = -1;
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString().replace(singleQuoteRegex, "").trim() === cleanOldNisn) { rowIndex = i + 1; break; }
      }
      if (rowIndex > -1) {
        if (cleanOldNisn !== cleanNisn && data.some(row => row[0].toString().replace(singleQuoteRegex, "").trim() === cleanNisn)) {
          return { success: false, message: `NISN ${cleanNisn} sudah terpakai oleh siswa lain!` };
        }
        sheet.getRange(rowIndex, 1, 1, 5).setValues([rowData]);
        return { success: true, message: `Data siswa berhasil diperbarui.` };
      }
      return { success: false, message: 'Data siswa tidak ditemukan.' };
    }
  } catch (error) { return { success: false, message: error.message }; }
}

function deleteStudent(nisn) {
  try {
    const sheet = getSheet(SHEET_SISWA);
    const data = sheet.getDataRange().getValues();
    const singleQuoteRegex = new RegExp("'", "g");
    const cleanNisn = nisn.replace(singleQuoteRegex, "").trim();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString().replace(singleQuoteRegex, "").trim() === cleanNisn) {
        sheet.deleteRow(i + 1);
        return { success: true, message: `Data siswa berhasil dihapus.` };
      }
    }
    return { success: false, message: 'Siswa tidak ditemukan.' };
  } catch (error) { return { success: false, message: error.message }; }
}

function getJamMengajar(nip, role) {
  try {
    const sheet = getSheet(SHEET_JADWAL);
    const data = sheet.getDataRange().getValues();
    const result = [];
    
    for (let i = 1; i < data.length; i++) {
      if(data[i][0] !== '') {
        const rowNip = data[i][1].toString().trim();
        if (role.toLowerCase() === 'admin' || rowNip.toLowerCase() === nip.toString().trim().toLowerCase()) {
          result.push({
            id: data[i][0].toString(),
            nip: rowNip,
            namaGuru: data[i][2].toString(),
            tanggal: data[i][3].toString(),
            jamKe: data[i][4].toString(),
            waktuStr: data[i][5].toString(),
            mapel: data[i][6].toString()
          });
        }
      }
    }
    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveJamMengajar(jadwalData, action, oldId) {
  try {
    const sheet = getSheet(SHEET_JADWAL);
    const data = sheet.getDataRange().getValues();
    let idJadwal = oldId;
    if (action === 'add') {
      idJadwal = 'JM-' + new Date().getTime();
    }

    const rowData = [
      idJadwal, 
      jadwalData.nip, 
      jadwalData.namaGuru, 
      jadwalData.tanggal, 
      jadwalData.jamKe, 
      jadwalData.waktuStr, 
      jadwalData.mapel
    ];

    if (action === 'add') {
      sheet.appendRow(rowData);
      return { success: true, message: `Jam mengajar berhasil ditambahkan.` };
    } else if (action === 'edit') {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === oldId.trim()) {
          sheet.getRange(i + 1, 1, 1, 7).setValues([rowData]);
          return { success: true, message: `Data jam mengajar diperbarui.` };
        }
      }
      return { success: false, message: 'Data jadwal tidak ditemukan.' };
    }
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan jadwal: ' + error.message };
  }
}

function deleteJamMengajar(idJadwal) {
  try {
    const sheet = getSheet(SHEET_JADWAL);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === idJadwal.toString()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: `Data jam mengajar berhasil dihapus.` };
      }
    }
    return { success: false, message: 'Data tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal menghapus jadwal: ' + error.message };
  }
}

function getAbsensiSession(tanggal, jamKe, idKelas) {
  try {
    const sheet = getSheet(SHEET_ABSENSI);
    const data = sheet.getDataRange().getValues();
    const searchId = `${tanggal}_${jamKe}_${idKelas}`;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === searchId) {
        return { 
          success: true, 
          exists: true,
          data: JSON.parse(data[i][7].toString() || '[]')
        };
      }
    }
    return { success: true, exists: false, data: [] };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveAbsensiSession(payload) {
  try {
    const sheet = getSheet(SHEET_ABSENSI);
    const data = sheet.getDataRange().getValues();
    const idAbsensi = `${payload.tanggal}_${payload.jamKe}_${payload.idKelas}`;
    const jsonKehadiran = JSON.stringify(payload.dataKehadiran);
    
    const rowData = [
      idAbsensi, 
      payload.tanggal, 
      payload.jamKe, 
      payload.idKelas, 
      payload.mapel, 
      payload.nip, 
      payload.namaGuru, 
      jsonKehadiran
    ];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === idAbsensi) {
        sheet.getRange(i + 1, 1, 1, 8).setValues([rowData]);
        return { success: true, message: `Data absensi kelas ${payload.idKelas} berhasil diperbarui!` };
      }
    }
    
    sheet.appendRow(rowData);
    return { success: true, message: `Data absensi kelas ${payload.idKelas} berhasil disimpan!` };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan absensi: ' + error.message };
  }
}

function getAbsensiForRecap() {
  try {
    const sheet = getSheet(SHEET_ABSENSI);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== '') {
        result.push({
          idAbsensi: data[i][0].toString(),
          tanggal: data[i][1].toString(),
          jamKe: data[i][2].toString(),
          idKelas: data[i][3].toString(),
          mapel: data[i][4].toString(),
          nipGuru: data[i][5].toString(),
          namaGuru: data[i][6].toString(),
          dataKehadiran: JSON.parse(data[i][7].toString() || '[]')
        });
      }
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getPenilaianSession(tanggal, jamKe, idKelas, jenisPenilaian) {
  try {
    const sheet = getSheet(SHEET_PENILAIAN);
    const data = sheet.getDataRange().getValues();
    const searchId = `${tanggal}_${jamKe}_${idKelas}_${jenisPenilaian}`;
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === searchId) {
        return { 
          success: true, 
          exists: true,
          data: JSON.parse(data[i][9].toString() || '[]')
        };
      }
    }
    return { success: true, exists: false, data: [] };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function savePenilaianSession(payload) {
  try {
    const sheet = getSheet(SHEET_PENILAIAN);
    const data = sheet.getDataRange().getValues();
    const idPenilaian = `${payload.tanggal}_${payload.jamKe}_${payload.idKelas}_${payload.jenis}`;
    const jsonNilai = JSON.stringify(payload.dataNilai);
    
    const rowData = [
      idPenilaian, 
      payload.tanggal, 
      payload.jamKe, 
      payload.idKelas, 
      payload.mapel, 
      payload.materi,
      payload.jenis,
      payload.nip, 
      payload.namaGuru, 
      jsonNilai
    ];

    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === idPenilaian) {
        sheet.getRange(i + 1, 1, 1, 10).setValues([rowData]);
        return { success: true, message: `Data penilaian berhasil diperbarui!` };
      }
    }
    
    sheet.appendRow(rowData);
    return { success: true, message: `Data penilaian berhasil disimpan!` };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan penilaian: ' + error.message };
  }
}

function getPenilaianForRecap() {
  try {
    const sheet = getSheet(SHEET_PENILAIAN);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== '') {
        result.push({
          idPenilaian: data[i][0].toString(),
          tanggal: data[i][1].toString(),
          jamKe: data[i][2].toString(),
          idKelas: data[i][3].toString(),
          mapel: data[i][4].toString(),
          materi: data[i][5].toString(),
          jenis: data[i][6].toString(),
          nipGuru: data[i][7].toString(),
          namaGuru: data[i][8].toString(),
          dataNilai: JSON.parse(data[i][9].toString() || '[]')
        });
      }
    }
    return { success: true, data: result };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function getInformasi() {
  try {
    const sheet = getSheet(SHEET_INFORMASI);
    const data = sheet.getDataRange().getValues();
    const result = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== '') {
        result.push({
          id: data[i][0].toString(),
          tanggal: data[i][1].toString(),
          isi: data[i][2].toString(),
          nipPembuat: data[i][3].toString()
        });
      }
    }
    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveInformasi(infoData, action, oldId) {
  try {
    const sheet = getSheet(SHEET_INFORMASI);
    const data = sheet.getDataRange().getValues();
    let idInfo = oldId;
    if (action === 'add') {
      idInfo = 'INF-' + new Date().getTime();
    }
    const rowData = [idInfo, infoData.tanggal, infoData.isi, infoData.nipPembuat];

    if (action === 'add') {
      sheet.appendRow(rowData);
      return { success: true, message: 'Informasi pengumuman berhasil ditambahkan.' };
    } else if (action === 'edit') {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === oldId.trim()) {
          sheet.getRange(i + 1, 1, 1, 4).setValues([rowData]);
          return { success: true, message: 'Informasi pengumuman berhasil diperbarui.' };
        }
      }
      return { success: false, message: 'Informasi tidak ditemukan.' };
    }
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function deleteInformasi(idInfo) {
  try {
    const sheet = getSheet(SHEET_INFORMASI);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === idInfo.toString()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Informasi berhasil dihapus.' };
      }
    }
    return { success: false, message: 'Data pengumuman tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function getJurnal(nip, role) {
  try {
    const sheet = getSheet(SHEET_JURNAL);
    const data = sheet.getDataRange().getValues();
    const result = [];
    const cleanRole = role.toString().trim().toLowerCase();
    
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== '') {
        const rowNip = data[i][7].toString().trim();
        if (cleanRole === 'admin' || rowNip.toLowerCase() === nip.toString().trim().toLowerCase()) {
          result.push({
            id: data[i][0].toString(),
            tanggal: data[i][1].toString(),
            jamKe: data[i][2].toString(),
            idKelas: data[i][3].toString(),
            mapel: data[i][4].toString(),
            materi: data[i][5].toString(),
            evaluasi: data[i][6].toString(),
            nip: rowNip,
            namaGuru: data[i][8].toString()
          });
        }
      }
    }
    return { success: true, data: result };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

function saveJurnal(jurnalData, action, oldId) {
  try {
    const sheet = getSheet(SHEET_JURNAL);
    const data = sheet.getDataRange().getValues();
    let idJurnal = oldId;
    if (action === 'add') {
      idJurnal = 'JR-' + new Date().getTime();
    }
    const rowData = [
      idJurnal,
      jurnalData.tanggal,
      jurnalData.jamKe,
      jurnalData.idKelas,
      jurnalData.mapel,
      jurnalData.materi,
      jurnalData.evaluasi,
      jurnalData.nip,
      jurnalData.namaGuru
    ];

    if (action === 'add') {
      sheet.appendRow(rowData);
      return { success: true, message: 'Jurnal mengajar berhasil disimpan.' };
    } else if (action === 'edit') {
      for (let i = 1; i < data.length; i++) {
        if (data[i][0].toString() === oldId.trim()) {
          sheet.getRange(i + 1, 1, 1, 9).setValues([rowData]);
          return { success: true, message: 'Catatan Jurnal mengajar berhasil diperbarui.' };
        }
      }
      return { success: false, message: 'Data jurnal mengajar tidak ditemukan.' };
    }
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan jurnal: ' + error.message };
  }
}

function deleteJurnal(idJurnal) {
  try {
    const sheet = getSheet(SHEET_JURNAL);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0].toString() === idJurnal.toString()) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Catatan Jurnal berhasil dihapus.' };
      }
    }
    return { success: false, message: 'Data jurnal tidak ditemukan.' };
  } catch (error) {
    return { success: false, message: 'Gagal menghapus jurnal: ' + error.message };
  }
}