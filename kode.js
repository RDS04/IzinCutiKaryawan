const SECRET_KEY = "erwinadmin123";

// ===== KATEGORI CUTI BERDASARKAN GENDER =====
const KATEGORI_CUTI_SEMUA = [
  'Libur',
  'Izin Sakit',
  'Keluarga Kandung Opname/Rawat inap',
  'Cuti Pasca Bersalin',           // Khusus Perempuan
  'Izin Selama Hamil',              // Khusus Perempuan
  'Istri Melahirkan',               // Khusus Laki-laki
  'Cuti Ibadah Haji',
  'Cuti Umroh',
  'Pernikahan Karyawan',
  'Keluarga Kandung Meninggal',
  'Pernikahan Anak Kandung',
  'Aqikah Anak Kandung',
  'Sunat Anak Kandung',
  'Operasi/Sakit Parah',
  'Pengganti Libur'
];

// Fungsi untuk mendapatkan kategori cuti berdasarkan gender
function getKategoriCutiByGender(gender) {
  const gender_lower = (gender || '').trim().toLowerCase();
  
  // Kategori untuk Perempuan - tidak boleh ambil "Istri Melahirkan"
  if (gender_lower === 'perempuan') {
    return KATEGORI_CUTI_SEMUA.filter(k => k !== 'Istri Melahirkan');
  }
  
  // Kategori untuk Laki-laki - tidak boleh ambil "Izin Selama Hamil" dan "Cuti Pasca Bersalin"
  if (gender_lower === 'laki-laki') {
    return KATEGORI_CUTI_SEMUA.filter(k => k !== 'Izin Selama Hamil' && k !== 'Cuti Pasca Bersalin');
  }
  
  // Default: semua kategori
  return KATEGORI_CUTI_SEMUA;
}

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Login').setTitle('Login Sistem Cuti');
}

function getPageContent(role) {
  if (role === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin').getContent();
  } else if (role === 'karyawan') {
    return HtmlService.createHtmlOutputFromFile('KaryawanDashboard').getContent();
  }
  return '<h1>Akses Ditolak. Peran tidak valid.</h1>';
}

// FUNGSI BARU: Untuk redirect ke login page
function getLoginPage() {
  return HtmlService.createHtmlOutputFromFile('Login').getContent();
}

function loginManual(email, password) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Karyawan');
    const data = sheet.getRange("A2:G" + sheet.getLastRow()).getValues();
    let penggunaDitemukan = null;
    for (const row of data) {
      if (row[2].toString().trim().toLowerCase() === email.trim().toLowerCase()) {
        penggunaDitemukan = { email: row[2], hashDisimpan: row[6], peran: row[5] };
        break;
      }
    }
    if (!penggunaDitemukan || !penggunaDitemukan.hashDisimpan) {
      throw new Error("Email atau password salah.");
    }
    const hashInput = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, password + SECRET_KEY)
                        .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
    if (hashInput !== penggunaDitemukan.hashDisimpan) {
      throw new Error("Email atau password salah.");
    }

    const peran = penggunaDitemukan.peran.trim().toLowerCase();
    const token = buatToken(email, peran);
    
    return { sukses: true, peran: peran, token: token };
  } catch (e) {
    return { sukses: false, pesan: e.message };
  }
}


function buatToken(email, peran) {
  const payload = { email: email, peran: peran, exp: new Date().getTime() + (60 * 60 * 1000) }; // Token 1 jam
  const token = Utilities.base64Encode(JSON.stringify(payload));
  CacheService.getUserCache().put(token, JSON.stringify({email: email, peran: peran}), 3600);
  return token;
}

function getEmailFromToken(token) {
   const sessionDataString = CacheService.getUserCache().get(token);
   if (!sessionDataString) return null;
   return JSON.parse(sessionDataString).email;
}


function getDataKaryawan(token) {
  try {
    const email = getEmailFromToken(token);
    if (!email) throw new Error("Sesi tidak valid atau telah berakhir. Silakan login kembali.");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    const sheetCuti = ss.getSheetByName('Database Cuti');
    const dataKaryawanRange = sheetKaryawan.getRange("A2:I" + sheetKaryawan.getLastRow()).getValues();
    let karyawanInfo = null;
    for (const row of dataKaryawanRange) {
      if (row[2].toString().trim().toLowerCase() === email.toLowerCase()) {
        karyawanInfo = { id: row[0], nama: row[1], jatahCuti: row[4], jenisKelamin: row[8] };
        break;
      }
    }
    if (!karyawanInfo) throw new Error("Karyawan tidak ditemukan.");
    
    // ✨ PENTING: Ambil dua jenis jatah yang berbeda
    // 1. Jatah "Libur" (kategori cuti) - dari tabel Libur
    const jatahLibur = getJatahLiburFromTabel(karyawanInfo.id) || 12;
    
    // 2. Jatah Cuti Tahunan (total) - dari kolom J "Jatah Cuti Tahunan"
    const jatahCutiTahunan = getJatahCutiTahunanKaryawan(karyawanInfo.id) || jatahLibur || 12;
    
    Logger.log('ID Karyawan: ' + karyawanInfo.id);
    Logger.log('Jatah Libur (kategori): ' + jatahLibur);
    Logger.log('Jatah Cuti Tahunan (total): ' + jatahCutiTahunan);
    
    const semuaCuti = sheetCuti.getRange("A2:I" + sheetCuti.getLastRow()).getValues();
    const riwayatCuti = semuaCuti.filter(row => row[2] === karyawanInfo.id);
    
    // ✨ Hitung TOTAL cuti yang sudah diambil (dari semua kategori, bukan hanya "Libur")
    let totalCutiDiambilSemua = 0;
    const tahunIni = new Date().getFullYear();
    riwayatCuti.forEach(row => {
      const tahunCuti = new Date(row[3]).getFullYear(); // Ambil tahun dari Tanggal Mulai
      if (row[8] === 'Disetujui' && tahunCuti === tahunIni) {
        totalCutiDiambilSemua += row[5];
      }
    });

    // RUMUS 1: Sisa Cuti Libur = Jatah Libur - Total Cuti Diambil (dari kategori Libur saja)
    let totalLiburDiambil = 0;
    riwayatCuti.forEach(row => {
      const tahunCuti = new Date(row[3]).getFullYear();
      if (row[8] === 'Disetujui' && tahunCuti === tahunIni && row[6] === 'Libur') {
        totalLiburDiambil += row[5];
      }
    });
    const sisaCuti = jatahLibur - totalLiburDiambil;
    
    // RUMUS 2: Sisa Jatah Cuti Tahunan = Jatah Cuti Tahunan - Total Cuti Diambil (dari SEMUA kategori)
    const sisaCutiTahunan = jatahCutiTahunan - totalCutiDiambilSemua;
    
    Logger.log('Total Cuti Diambil (Libur saja): ' + totalLiburDiambil);
    Logger.log('Total Cuti Diambil (SEMUA kategori): ' + totalCutiDiambilSemua);
    Logger.log('Sisa Cuti Libur: ' + sisaCuti);
    Logger.log('Sisa Cuti Tahunan: ' + sisaCutiTahunan);
    
    const riwayatTerformat = riwayatCuti.map(row => ({
        jenis: row[6], 
        tanggal: `${new Date(row[3]).toLocaleDateString('id-ID')} - ${new Date(row[4]).toLocaleDateString('id-ID')}`,
        lama: row[5], 
        status: row[8]
    })).reverse();
    
    // Ambil data Pengganti Libur jika ada
    let penggantiLiburData = null;
    try {
      const penggantiResult = getPenggantiLiburKaryawan(karyawanInfo.id);
      if (penggantiResult && penggantiResult.pengganti) {
        penggantiLiburData = penggantiResult.pengganti;
      }
    } catch (e) {
      Logger.log('Warning: Gagal ambil data pengganti libur - ' + e.message);
    }
    
    return { 
      nama: karyawanInfo.nama, 
      sisaCuti: sisaCuti,                              // Sisa Libur
      sisaCutiTahunan: sisaCutiTahunan,               // Sisa Jatah Cuti Tahunan (BARU)
      riwayat: riwayatTerformat,
      jatahLibur: jatahLibur,                          // Untuk kategori "Libur"
      jatahCutiTahunan: jatahCutiTahunan,             // Untuk display di header (original)
      totalCutiDiambilSemua: totalCutiDiambilSemua,   // Total cuti sudah diambil (SEMUA kategori)
      jenisKelamin: karyawanInfo.jenisKelamin,
      kategoriCutiTersedia: getKategoriCutiByGender(karyawanInfo.jenisKelamin),
      penggantiLibur: penggantiLiburData
    };
  } catch(e) { return { error: e.message }; }
}

// FUNGSI HELPER: Cari kolom berdasarkan header name
function findColumnByHeader(headers, searchTerms) {
  for (let j = 0; j < headers.length; j++) {
    const header = headers[j].toString().trim().toLowerCase();
    for (const term of searchTerms) {
      if (header.includes(term.toLowerCase())) {
        return j;
      }
    }
  }
  return null;
}

// FUNGSI BARU: Ambil jatah "Libur" (kategori cuti) dari tabel Libur di sheet "Data Karyawan"
function getJatahLiburFromTabel(idKaryawan) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    const dataKaryawan = sheetKaryawan.getDataRange().getValues();
    const headers = dataKaryawan[0] || [];
    
    // ✨ Cari kolom "Libur" SPESIFIK (bukan "Jatah Cuti Tahunan")
    // Hanya cari kolom yang EXACTLY match "Libur" atau mengandung "Libur" saja
    let columnJatahLibur = null;
    for (let j = 0; j < headers.length; j++) {
      const header = headers[j].toString().trim().toLowerCase();
      // Cari yang EXACTLY "libur" atau mengandung "libur" tapi TIDAK mengandung "jatah" atau "tahunan"
      if (header === 'libur' || (header.includes('libur') && !header.includes('jatah') && !header.includes('tahunan'))) {
        columnJatahLibur = j;
        break;
      }
    }
    
    if (columnJatahLibur === null) {
      Logger.log('⚠ Kolom "Libur" (kategori) tidak ditemukan');
      return null;
    }
    
    Logger.log('✓ Kolom "Libur" (kategori) ditemukan di kolom ' + (columnJatahLibur + 1) + ' (' + headers[columnJatahLibur] + ')');
    
    // Cari karyawan dan ambil nilai jatah liburnya
    for (let i = 1; i < dataKaryawan.length; i++) {
      const id = dataKaryawan[i][0];
      if (id === idKaryawan || id.toString().trim() === idKaryawan.toString().trim()) {
        const jatahValue = dataKaryawan[i][columnJatahLibur];
        const jatahInt = parseInt(jatahValue) || 0;
        Logger.log('✓ Jatah Libur untuk ' + idKaryawan + ': ' + jatahInt + ' hari');
        return jatahInt;
      }
    }
    
    Logger.log('⚠ Karyawan dengan ID ' + idKaryawan + ' tidak ditemukan');
    return null;
  } catch (e) {
    Logger.log('Error di getJatahLiburFromTabel: ' + e.message);
    return null;
  }
}

// FUNGSI BARU: Ambil "Jatah Cuti Tahunan" dari kolom J (atau header "Jatah Cuti Tahunan")
function getJatahCutiTahunanKaryawan(idKaryawan) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    const dataKaryawan = sheetKaryawan.getDataRange().getValues();
    const headers = dataKaryawan[0] || [];
    
    // ✨ Cari kolom "Jatah Cuti Tahunan" atau default ke kolom J (index 9)
    let columnJatahTahunan = null;
    
    // Cari berdasarkan header name
    for (let j = 0; j < headers.length; j++) {
      const header = headers[j].toString().trim().toLowerCase();
      if (header.includes('jatah cuti tahunan') || header === 'jatah tahunan') {
        columnJatahTahunan = j;
        Logger.log('✓ Kolom "Jatah Cuti Tahunan" ditemukan di kolom ' + (j + 1) + ' (' + headers[j] + ')');
        break;
      }
    }
    
    // Jika tidak ditemukan di header, gunakan kolom J (index 9)
    if (columnJatahTahunan === null) {
      columnJatahTahunan = 9; // Kolom J
      Logger.log('✓ Kolom "Jatah Cuti Tahunan" default ke kolom J (index 9)');
    }
    
    // Ambil nilai untuk karyawan ini
    for (let i = 1; i < dataKaryawan.length; i++) {
      const id = dataKaryawan[i][0];
      if (id === idKaryawan || id.toString().trim() === idKaryawan.toString().trim()) {
        const jatahValue = dataKaryawan[i][columnJatahTahunan];
        const jatahInt = parseInt(jatahValue) || 0;
        Logger.log('✓ Jatah Cuti Tahunan untuk ' + idKaryawan + ': ' + jatahInt + ' hari');
        return jatahInt;
      }
    }
    
    Logger.log('⚠ Karyawan dengan ID ' + idKaryawan + ' tidak ditemukan');
    return null;
  } catch (e) {
    Logger.log('Error di getJatahCutiTahunanKaryawan: ' + e.message);
    return null;
  }
}

// FUNGSI BARU: Ambil jatah cuti berdasarkan ID karyawan dan tahun
function getJatahCutiPerTahun(idKaryawan, tahun) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Coba ambil dari sheet "Cuti Tahunan" jika ada
    let sheetCutiTahunan = null;
    try {
      sheetCutiTahunan = ss.getSheetByName('Cuti Tahunan');
    } catch (e) {
      Logger.log('Sheet "Cuti Tahunan" tidak ditemukan');
    }
    
    if (sheetCutiTahunan) {
      const data = sheetCutiTahunan.getDataRange().getValues();
      // Asumsikan format: Kolom A = ID Karyawan, Kolom B = Tahun, Kolom C = Jatah
      // Atau bisa juga: Kolom A = ID, Kolom B = 2025, Kolom C = 2026, dll
      for (let i = 1; i < data.length; i++) {
        if (data[i][0] === idKaryawan || data[i][0].toString().trim() === idKaryawan.toString().trim()) {
          // Format: Kolom A = ID, Kolom B = Tahun, Kolom C = Jatah
          if (data[i][1] && parseInt(data[i][1]) === tahun) {
            return parseInt(data[i][2]) || null;
          }
          // Format alternatif: Kolom A = ID, Kolom B = 2025, Kolom C = 2026, etc
          const columnIndex = 1 + (tahun - 2025); // Asumsikan tahun mulai dari 2025 di kolom B
          if (columnIndex >= 1 && columnIndex < data[i].length) {
            const value = parseInt(data[i][columnIndex]);
            if (value) return value;
          }
        }
      }
    }
    
    // ✨ PENTING: Gunakan tabel Libur dari sheet "Data Karyawan" untuk kategori LIBUR
    const jatahLibur = getJatahLiburFromTabel(idKaryawan);
    if (jatahLibur !== null) {
      Logger.log('✓ Jatah Libur dari tabel: ' + jatahLibur);
      return jatahLibur;
    }
    
    Logger.log('⚠ Tidak bisa ambil data jatah libur');
    return null;
  } catch (e) {
    Logger.log('Error di getJatahCutiPerTahun: ' + e.message);
    return null;
  }
}


function simpanPengajuanCuti(dataForm, token) {
  try {
    const email = getEmailFromToken(token);
    if (!email) throw new Error("Sesi tidak valid atau telah berakhir. Silakan login kembali.");
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    const sheetCuti = ss.getSheetByName('Database Cuti');
    const dataKaryawanRange = sheetKaryawan.getRange("A2:F" + sheetKaryawan.getLastRow()).getValues();
    let karyawanInfo = null;
    for (const row of dataKaryawanRange) {
      if (row[2].toString().trim().toLowerCase() === email.toLowerCase()) {
        karyawanInfo = { id: row[0], jatahCuti: row[4] };
        break;
      }
    }
    if (!karyawanInfo) throw new Error("Karyawan tidak ditemukan.");
    const semuaCuti = sheetCuti.getRange("A2:I" + sheetCuti.getLastRow()).getValues();
    
    // PERBAIKAN: Hitung sisa cuti hanya untuk tahun ini
    let totalCutiDiambil = 0;
    const tahunIni = new Date().getFullYear();
    semuaCuti.forEach(row => {
      const tahunCuti = new Date(row[3]).getFullYear();
      if (row[2] === karyawanInfo.id && row[8] === 'Disetujui' && tahunCuti === tahunIni) {
        totalCutiDiambil += row[5];
      }
    });

    const sisaCuti = karyawanInfo.jatahCuti - totalCutiDiambil;
    const tglMulai = new Date(dataForm.tglMulai);
    const tglSelesai = new Date(dataForm.tglSelesai);
    const jumlahHariDiajukan = ((tglSelesai - tglMulai) / (1000 * 60 * 60 * 24)) + 1;
    if (jumlahHariDiajukan > sisaCuti) {
      throw new Error(`Pengajuan ditolak. Jumlah hari yang Anda minta (${jumlahHariDiajukan}) melebihi sisa cuti Anda (${sisaCuti}).`);
    }
    const idPengajuan = 'CUTI-' + new Date().getTime();
    sheetCuti.appendRow([idPengajuan, new Date(), karyawanInfo.id, tglMulai, tglSelesai, jumlahHariDiajukan, dataForm.jenisCuti, dataForm.keterangan, 'Menunggu Persetujuan']);
    return 'Pengajuan cuti Anda telah berhasil dikirim!';
  } catch (e) { return 'Error: ' + e.message; }
}

// FUNGSI BARU: Simpan pengajuan cuti dengan kategori
function simpanPengajuanCutiKategori(dataForm, token) {
  try {
    Logger.log('=== simpanPengajuanCutiKategori START ===');
    Logger.log('Input dataForm:', JSON.stringify(dataForm));
    
    const email = getEmailFromToken(token);
    Logger.log('Email dari token:', email);
    if (!email) throw new Error("Sesi tidak valid atau telah berakhir. Silakan login kembali.");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    const sheetCuti = ss.getSheetByName('Database Cuti');
    
    const dataKaryawanRange = sheetKaryawan.getRange("A2:H" + sheetKaryawan.getLastRow()).getValues();
    let karyawanInfo = null;
    
    for (const row of dataKaryawanRange) {
      if (row[2].toString().trim().toLowerCase() === email.toLowerCase()) {
        karyawanInfo = { id: row[0], nama: row[1], jenisKelamin: row[7] };
        Logger.log('Karyawan ditemukan:', karyawanInfo);
        break;
      }
    }
    
    if (!karyawanInfo) throw new Error("Karyawan tidak ditemukan.");
    
    // VALIDASI GENDER - Cek apakah kategori sesuai dengan gender
    const kategoriAllowed = getKategoriCutiByGender(karyawanInfo.jenisKelamin);
    if (!kategoriAllowed.includes(dataForm.kategoriNama)) {
      const jenisKelamin = (karyawanInfo.jenisKelamin || '').trim();
      if (jenisKelamin.toLowerCase() === 'laki-laki') {
        throw new Error(`Kategori "${dataForm.kategoriNama}" tidak tersedia untuk karyawan laki-laki.`);
      } else if (jenisKelamin.toLowerCase() === 'perempuan') {
        throw new Error(`Kategori "${dataForm.kategoriNama}" tidak tersedia untuk karyawan perempuan.`);
      }
      throw new Error(`Kategori "${dataForm.kategoriNama}" tidak tersedia untuk jenis kelamin Anda.`);
    }
    Logger.log('✓ Validasi gender OK');
    
    // Validasi tanggal
    const tglMulai = new Date(dataForm.tglMulai);
    const tglSelesai = new Date(dataForm.tglSelesai);
    Logger.log('Tanggal Mulai:', tglMulai, 'Tanggal Selesai:', tglSelesai);
    
    if (tglMulai > tglSelesai) {
      throw new Error("Tanggal mulai tidak boleh lebih besar dari tanggal selesai.");
    }
    
    // ✨ PENTING: HITUNG HARI CUTI BERDASARKAN ATURAN WEEKEND
    let jumlahHari = 0;
    const kategoriManualInput = ['Izin Sakit', 'Keluarga Kandung Opname/Rawat inap'].includes(dataForm.kategoriNama);
    const kategoriIncludeWeekend = ['Cuti Pasca Bersalin', 'Operasi/Sakit Parah', 'Cuti Ibadah Haji', 'Cuti Umroh'];
    const gunakanKalenderPenuh = kategoriIncludeWeekend.includes(dataForm.kategoriNama);
    
    if (kategoriManualInput) {
      // Kategori manual: gunakan input manual dari pengguna
      jumlahHari = dataForm.jumlahHari || 1;
      Logger.log(dataForm.kategoriNama + ' - Menggunakan input manual:', jumlahHari);
    } else if (gunakanKalenderPenuh) {
      // Kategori tertentu WAJIB penuh termasuk Sabtu/Minggu
      jumlahHari = Math.floor((tglSelesai - tglMulai) / (1000 * 60 * 60 * 24)) + 1;
      Logger.log('Hitung kalender penuh (termasuk weekend):', jumlahHari);
    } else {
      // Default: hitung hari kerja (skip Sabtu, Minggu, libur nasional)
      jumlahHari = hitungHariKerja(tglMulai, tglSelesai);
      Logger.log('Hitung otomatis hari kerja (skip weekend & libur nasional):', jumlahHari);
    }
    
    if (jumlahHari <= 0) {
      throw new Error("Jumlah hari tidak valid. Pilih tanggal yang memiliki hari kerja.");
    }
    
    Logger.log('Jumlah hari final:', jumlahHari);
    
    // Validasi kategori non-cicil dan Libur Kerja - cek riwayat tahun ini
    const tahunIni = new Date().getFullYear();
    const dataCutiRange = sheetCuti.getRange("A2:I" + sheetCuti.getLastRow()).getValues();
    Logger.log('Jumlah data cuti existing:', dataCutiRange.length);
    
    // Untuk kategori non-cicil: check apakah sudah pernah diambil tahun ini
    const kategoriNonCicil = ['Cuti Pasca Bersalin', 'Istri Melahirkan', 'Operasi/Sakit Parah', 
                             'Cuti Ibadah Haji', 'Cuti Umroh', 'Pernikahan Karyawan', 
                             'Keluarga Kandung Meninggal', 'Pernikahan Anak Kandung', 
                             'Aqikah Anak Kandung', 'Sunat Anak Kandung', 'Izin Selama Hamil'];
    
    if (kategoriNonCicil.includes(dataForm.kategoriNama)) {
      Logger.log('Cek kategori non-cicil:', dataForm.kategoriNama);
      for (const row of dataCutiRange) {
        if (row[2] === karyawanInfo.id && row[6] === dataForm.kategoriNama && row[8] === 'Disetujui') {
          const tahunPengajuan = new Date(row[3]).getFullYear();
          if (tahunPengajuan === tahunIni) {
            throw new Error(`Anda sudah mengambil ${dataForm.kategoriNama} tahun ini. Kategori ini hanya bisa diambil sekali dalam setahun.`);
          }
        }
      }
      Logger.log('✓ Validasi kategori non-cicil OK');
    }
    
    // Untuk Libur Kerja: check apakah karyawan sudah mengambil libur kerja ini
    if (dataForm.isLiburKerja) {
      Logger.log('Cek Libur Kerja');
      for (const row of dataCutiRange) {
        if (row[2] === karyawanInfo.id && row[6] === 'Libur Kerja' && row[8] === 'Disetujui') {
          const tglMuaiCuti = new Date(row[3]);
          const tglSelesaiCuti = new Date(row[4]);
          // Cek apakah periode ini sudah terambil
          if (dataForm.liburKerjaId === row[10]) { // Kolom K untuk Libur Kerja ID
            throw new Error("Anda sudah mengambil libur kerja ini.");
          }
        }
      }
      Logger.log('✓ Validasi Libur Kerja OK');
    }
    
    let linkGambarDokter = '';
    
    const idPengajuan = 'CUTI-' + new Date().getTime();
    Logger.log('ID Pengajuan:', idPengajuan);
    
    // Tentukan status: auto-approve untuk Pengganti Libur, biasa untuk yang lain
    let status = 'Menunggu Persetujuan';
    if (dataForm.autoApprove) {
      status = 'Disetujui';
      Logger.log('✓ Auto-approving Pengganti Libur for ' + karyawanInfo.id);
    }
    Logger.log('Status pengajuan:', status);
    
    Logger.log('Append row ke sheet dengan data:');
    Logger.log([
      idPengajuan, 
      new Date(), 
      karyawanInfo.id, 
      tglMulai, 
      tglSelesai, 
      jumlahHari, 
      dataForm.kategoriNama, 
      dataForm.keterangan, 
      status,
      linkGambarDokter,
      dataForm.isPenggantiLibur ? 'PENGGANTI' : (dataForm.liburKerjaId || '')
    ]);
    
    sheetCuti.appendRow([
      idPengajuan, 
      new Date(), 
      karyawanInfo.id, 
      tglMulai, 
      tglSelesai, 
      jumlahHari, 
      dataForm.kategoriNama, 
      dataForm.keterangan, 
      status,
      linkGambarDokter,  // Kolom J: Link Gambar Dokter
      dataForm.isPenggantiLibur ? 'PENGGANTI' : (dataForm.liburKerjaId || '')  // Kolom K: Pengganti Libur marker
    ]);
    
    Logger.log('=== simpanPengajuanCutiKategori SUKSES ===');
    return 'Pengajuan cuti Anda telah berhasil dikirim!';
  } catch (e) { 
    Logger.log('=== ERROR simpanPengajuanCutiKategori ===');
    Logger.log('Error message:', e.message);
    Logger.log('Error stack:', e.stack);
    return 'Error: ' + e.message; 
  }
}

// --- FUNGSI DASHBOARD ADMIN ---

function kelolaKaryawan(payload) {
  const { action, data } = payload;
  if (action === 'TAMBAH') {
    return tambahKaryawan(data);
  } else if (action === 'UPDATE') {
    return updateKaryawan(data);
  }
}

function tambahKaryawan(dataKaryawan) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Karyawan');
    const ids = sheet.getRange("A2:A").getValues().flat();
    if (ids.includes(dataKaryawan.id)) {
      throw new Error('ID Karyawan sudah ada.');
    }
    
    if (!dataKaryawan.peran || !dataKaryawan.passwordSementara) {
      throw new Error('Peran dan Password Sementara wajib diisi.');
    }
    
    if (!dataKaryawan.jenisKelamin) {
      throw new Error('Jenis Kelamin wajib diisi.');
    }

    const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, dataKaryawan.passwordSementara + SECRET_KEY)
                         .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');

    sheet.appendRow([ 
      dataKaryawan.id, 
      dataKaryawan.nama, 
      dataKaryawan.email, 
      new Date(dataKaryawan.tanggalBergabung), 
      dataKaryawan.jatahCuti,
      dataKaryawan.peran,
      hash,
      '',
      dataKaryawan.jenisKelamin
    ]);
    return 'Karyawan baru berhasil ditambahkan.';
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

function updateKaryawan(dataKaryawan) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Karyawan');
    const ids = sheet.getRange("A2:A").getValues().flat();
    const rowIndex = ids.indexOf(dataKaryawan.id);
    if (rowIndex === -1) throw new Error("Karyawan tidak ditemukan.");
    sheet.getRange(rowIndex + 2, 2, 1, 6).setValues([[ dataKaryawan.nama, dataKaryawan.email, new Date(dataKaryawan.tanggalBergabung), dataKaryawan.jatahCuti, '', dataKaryawan.jenisKelamin ]]);
    return "Data karyawan berhasil diperbarui.";
  } catch (e) {
    return "Error: " + e.message;
  }
}

function getKaryawanFullList(searchTerm) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Karyawan');
    let data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    Logger.log('Raw data from sheet:', data);
    Logger.log('Total rows:', data.length);
    
    if (searchTerm && searchTerm.length > 0) {
      const keyword = searchTerm.toLowerCase();
      data = data.filter(row => row[0].toLowerCase().includes(keyword) || row[1].toLowerCase().includes(keyword));
    }
    
    const result = data.map(row => {
      Logger.log('Processing row:', row);
      return {
        id: row[0],
        nama: row[1],
        email: row[2],
        tanggalBergabung: new Date(row[3]).toLocaleDateString('id-ID'),
        jatahCuti: row[4]
      };
    });
    
    Logger.log('Final result:', result);
    return result;
  } catch (e) { 
    Logger.log('Error in getKaryawanFullList:', e);
    return { error: e.toString() }; 
  }
}

function getKaryawanById(idKaryawan) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Karyawan');
    const ids = sheet.getRange("A2:A").getValues().flat();
    const rowIndex = ids.indexOf(idKaryawan);
    if (rowIndex === -1) throw new Error("Karyawan tidak ditemukan.");
    const dataRow = sheet.getRange(rowIndex + 2, 1, 1, sheet.getLastColumn()).getValues()[0];
    const tglBergabung = new Date(dataRow[3]);
    const tglFormatted = tglBergabung.getFullYear() + '-' + ('0' + (tglBergabung.getMonth() + 1)).slice(-2) + '-' + ('0' + tglBergabung.getDate()).slice(-2);
    return { id: dataRow[0], nama: dataRow[1], email: dataRow[2], tanggalBergabung: tglFormatted, jatahCuti: dataRow[4], jenisKelamin: dataRow[8] };
  } catch (e) { return { error: e.message }; }
}

function hapusKaryawan(idKaryawan) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Karyawan');
    const ids = sheet.getRange("A2:A").getValues().flat();
    const rowIndex = ids.indexOf(idKaryawan);
    if (rowIndex === -1) throw new Error("Karyawan tidak ditemukan.");
    sheet.deleteRow(rowIndex + 2);
    return "Karyawan berhasil dihapus.";
  } catch (e) { return "Error: " + e.message; }
}

function getPengajuanCutiList(filter) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCuti = ss.getSheetByName('Database Cuti');
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    let dataCuti = sheetCuti.getRange(2, 1, sheetCuti.getLastRow() - 1, sheetCuti.getLastColumn()).getValues();
    const dataKaryawan = sheetKaryawan.getRange(2, 1, sheetKaryawan.getLastRow() - 1, 2).getValues();
    const namaKaryawanMap = dataKaryawan.reduce((map, row) => { map[row[0]] = row[1]; return map; }, {});
    if (filter) {
      if (filter.nama && filter.nama.length > 0) {
        const keyword = filter.nama.toLowerCase();
        dataCuti = dataCuti.filter(row => (namaKaryawanMap[row[2]] || '').toLowerCase().includes(keyword));
      }
      if (filter.jenisCuti && filter.jenisCuti !== '') {
        dataCuti = dataCuti.filter(row => row[6] === filter.jenisCuti);
      }
      if (filter.status && filter.status !== '') {
        dataCuti = dataCuti.filter(row => row[8] === filter.status);
      }
    }
    const daftarPengajuan = dataCuti.map(row => ({
      idPengajuan: row[0], timestamp: new Date(row[1]).toLocaleDateString('id-ID'), idKaryawan: row[2],
      namaKaryawan: namaKaryawanMap[row[2]] || 'N/A',
      tglMulai: new Date(row[3]).toLocaleDateString('id-ID'),
      tglSelesai: new Date(row[4]).toLocaleDateString('id-ID'),
      jumlahHari: row[5], jenisCuti: row[6], keterangan: row[7], status: row[8],
      linkGambar: row[9] || ''  // Kolom J untuk link dokumen
    }));
    return daftarPengajuan.reverse();
  } catch (e) { return { error: e.toString() }; }
}

function updateStatusCuti(idPengajuan, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCuti = ss.getSheetByName('Database Cuti');
    const ids = sheetCuti.getRange("A2:A").getValues().flat();
    const rowIndex = ids.indexOf(idPengajuan);

    if (rowIndex === -1) {
      throw new Error("Pengajuan Cuti tidak ditemukan.");
    }

    sheetCuti.getRange(rowIndex + 2, 9).setValue(status);

    // ===== LOGIKA UNTUK PENGGANTI LIBUR =====
    // Jika status disetujui dan kategorinya Pengganti Libur, update hari digunakan
    if (status === 'Disetujui') {
      const dataCutiRow = sheetCuti.getRange(rowIndex + 2, 1, 1, 11).getValues()[0];
      const isPenggantiLibur = dataCutiRow[10]; // Kolom K: marker PENGGANTI
      const idKaryawan = dataCutiRow[2]; // Kolom C: ID Karyawan
      const jumlahHari = dataCutiRow[5]; // Kolom F: Jumlah Hari
      
      if (isPenggantiLibur === 'PENGGANTI') {
        Logger.log('Pengganti Libur detected, updating hari digunakan');
        const updated = updateHariDigunakan(idKaryawan, jumlahHari);
        if (updated) {
          Logger.log('✓ Hari digunakan berhasil diupdate');
        } else {
          Logger.log('⚠ Gagal mengupdate hari digunakan, tapi pengajuan tetap disetujui');
        }
      }
      
      // Simpan rekapan cuti otomatis
      simpanRekapanCutiOtomatis(ss);
    }

    // --- LOGIKA PENGIRIMAN EMAIL  ---

    // 1. Ambil detail yang diperlukan untuk email
    const dataCutiRow = sheetCuti.getRange(rowIndex + 2, 1, 1, 9).getValues()[0];
    const idKaryawan = dataCutiRow[2];
    const tanggalMulai = new Date(dataCutiRow[3]).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });
    const tanggalSelesai = new Date(dataCutiRow[4]).toLocaleDateString('id-ID', { day: 'numeric', month: 'long', year: 'numeric' });

    // 2. Cari email karyawan dari sheet Data Karyawan
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    const dataKaryawan = sheetKaryawan.getRange("A2:C" + sheetKaryawan.getLastRow()).getValues();
    let emailKaryawan = null;
    let namaKaryawan = null;

    for (const row of dataKaryawan) {
      if (row[0] === idKaryawan) {
        namaKaryawan = row[1];
        emailKaryawan = row[2];
        break;
      }
    }

    // 3. Kirim email jika email karyawan ditemukan
    if (emailKaryawan) {
      const subjek = `Update Pengajuan Cuti Anda - ${status}`;
      const isiEmail = `
        <p>Halo ${namaKaryawan},</p>
        <p>Pengajuan cuti Anda untuk tanggal <b>${tanggalMulai}</b> hingga <b>${tanggalSelesai}</b> telah di-update dengan status:</p>
        <h2 style="color:${status === 'Disetujui' ? 'green' : 'red'};">${status}</h2>
        <p>Terima kasih.</p>
        <p><em>Email ini dibuat secara otomatis oleh Sistem Cuti.</em></p>
      `;

      MailApp.sendEmail({
        to: emailKaryawan,
        subject: subjek,
        htmlBody: isiEmail
      });
    }

    return `Status pengajuan ${idPengajuan} berhasil diubah menjadi ${status}.`;
  } catch (e) {
    return "Error: " + e.message;
  }
}

// FUNGSI BARU: Update Hari Digunakan saat Pengganti Libur disetujui
function updateHariDigunakan(idKaryawan, jumlahHari) {
  try {
    Logger.log('updateHariDigunakan called for ID: ' + idKaryawan + ', jumlah: ' + jumlahHari);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetPengganti = null;
    
    try {
      sheetPengganti = ss.getSheetByName('Pengganti Libur');
    } catch (e) {
      Logger.log('Pengganti Libur sheet not found');
      return false;
    }
    
    if (!sheetPengganti || sheetPengganti.getLastRow() < 2) {
      Logger.log('Pengganti Libur sheet is empty');
      return false;
    }
    
    const lastRow = sheetPengganti.getLastRow();
    const ids = sheetPengganti.getRange("A2:A" + lastRow).getValues().flat();
    const idToFind = String(idKaryawan).trim();
    let rowIndex = -1;
    
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i]).trim() === idToFind) {
        rowIndex = i;
        break;
      }
    }
    
    if (rowIndex === -1) {
      Logger.log('Employee not found in Pengganti Libur sheet');
      return false;
    }
    
    // Baca data saat ini
    const currentData = sheetPengganti.getRange(rowIndex + 2, 1, 1, 6).getValues()[0];
    const durasiTotal = parseInt(currentData[2]) || 0;
    const hariDigunakan = parseInt(currentData[3]) || 0;
    const newHariDigunakan = hariDigunakan + parseInt(jumlahHari);
    const newSisaHari = durasiTotal - newHariDigunakan;
    
    // Validasi: jangan sampai hari digunakan melebihi durasi total
    if (newHariDigunakan > durasiTotal) {
      Logger.log('Warning: Hari digunakan (' + newHariDigunakan + ') melebihi durasi total (' + durasiTotal + ')');
      return false;
    }
    
    // Update sheet dengan data baru
    sheetPengganti.getRange(rowIndex + 2, 4, 1, 2).setValues([[
      newHariDigunakan,
      newSisaHari
    ]]);
    
    Logger.log('✓ Updated Hari Digunakan from ' + hariDigunakan + ' to ' + newHariDigunakan);
    Logger.log('✓ Updated Sisa Hari: ' + newSisaHari);
    return true;
  } catch (e) {
    Logger.log('Error in updateHariDigunakan: ' + e.message);
    return false;
  }
}

// FUNGSI BARU: Simpan rekapan cuti otomatis ke sheet "Rekapan Cuti"
function simpanRekapanCutiOtomatis(ss) {
  try {
    Logger.log('=== simpanRekapanCutiOtomatis START ===');
    
    const tahunIni = new Date().getFullYear();
    const result = getRekapanCutiDetail();
    
    if (result.error) {
      Logger.log('Error di getRekapanCutiDetail: ' + result.error);
      return;
    }

    const { daftarRekapan, allKategori } = result;
    
    // Cek atau buat sheet "Rekapan Cuti"
    let sheetRekapan = null;
    try {
      sheetRekapan = ss.getSheetByName('Rekapan Cuti');
    } catch (e) {
      Logger.log('Sheet Rekapan Cuti tidak ada, membuat sheet baru...');
      sheetRekapan = ss.insertSheet('Rekapan Cuti');
    }

    // Hapus semua data lama di sheet (kecuali header)
    if (sheetRekapan.getLastRow() > 1) {
      sheetRekapan.deleteRows(2, sheetRekapan.getLastRow() - 1);
    }

    // Header: No | Nama | Kategori1 | Kategori2 | ... | Total Diambil | Jatah Tahun Ini | Sisa
    const header = ['No.', 'Nama Karyawan', ...allKategori, 'Total Diambil', `Jatah Tahun ${tahunIni}`, 'Sisa Cuti'];
    
    // Jika belum ada header, buat
    if (sheetRekapan.getLastRow() === 0) {
      sheetRekapan.appendRow(header);
    } else {
      sheetRekapan.getRange(1, 1, 1, header.length).setValues([header]);
    }

    // Masukkan data rekapan
    (daftarRekapan || []).forEach((item, idx) => {
      const kategoriCounts = allKategori.map(k => item.kategoriCount?.[k] || 0);
      const row = [
        idx + 1,
        item.nama || 'N/A',
        ...kategoriCounts,
        item.totalCutiDiambil || 0,
        item.jatahCutiTahunan || 12,  // Gunakan jatahCutiTahunan
        item.sisaCuti || 0
      ];
      sheetRekapan.appendRow(row);
    });

    // Styling
    const headerRange = sheetRekapan.getRange(1, 1, 1, header.length);
    headerRange.setFontWeight('bold').setBackground('#4285F4').setFontColor('#FFFFFF');
    sheetRekapan.setFrozenRows(1);
    sheetRekapan.autoResizeColumns(1, header.length);

    Logger.log('=== simpanRekapanCutiOtomatis SUKSES ===');
    Logger.log(`Sheet "Rekapan Cuti" berhasil diupdate dengan ${daftarRekapan.length} karyawan`);
    
  } catch (e) {
    Logger.log('=== ERROR simpanRekapanCutiOtomatis ===');
    Logger.log('Error: ' + e.message);
  }
}

function getDetailCuti(idPengajuan) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCuti = ss.getSheetByName('Database Cuti');
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    const idsCuti = sheetCuti.getRange("A2:A").getValues().flat();
    const rowIndexCuti = idsCuti.indexOf(idPengajuan);
    if (rowIndexCuti === -1) throw new Error("Pengajuan Cuti tidak ditemukan.");
    const dataCutiRow = sheetCuti.getRange(rowIndexCuti + 2, 1, 1, sheetCuti.getLastColumn()).getValues()[0];
    const idKaryawan = dataCutiRow[2];
    const idsKaryawan = sheetKaryawan.getRange("A2:A").getValues().flat();
    const rowIndexKaryawan = idsKaryawan.indexOf(idKaryawan);
    if (rowIndexKaryawan === -1) throw new Error("Data Karyawan tidak ditemukan.");
    const dataKaryawanRow = sheetKaryawan.getRange(rowIndexKaryawan + 2, 1, 1, sheetKaryawan.getLastColumn()).getValues()[0];
    const jatahCutiTahunan = dataKaryawanRow[4];

    // PERBAIKAN: Hitung sisa cuti hanya untuk tahun ini
    const semuaCutiKaryawan = sheetCuti.getRange("A2:I" + sheetCuti.getLastRow()).getValues();
    let totalCutiDiambil = 0;
    const tahunIni = new Date().getFullYear();
    semuaCutiKaryawan.forEach(row => {
      const tahunCuti = new Date(row[3]).getFullYear();
      if (row[2] === idKaryawan && row[8] === 'Disetujui' && tahunCuti === tahunIni) {
        totalCutiDiambil += row[5];
      }
    });

    const sisaCuti = jatahCutiTahunan - totalCutiDiambil;
    return {
      nama: dataKaryawanRow[1], idKaryawan: idKaryawan,
      tanggalPengajuan: new Date(dataCutiRow[1]).toLocaleDateString('id-ID'),
      jenisCuti: dataCutiRow[6],
      tanggalCuti: `${new Date(dataCutiRow[3]).toLocaleDateString('id-ID')} - ${new Date(dataCutiRow[4]).toLocaleDateString('id-ID')}`,
      lamaCuti: dataCutiRow[5], keterangan: dataCutiRow[7], sisaCuti: sisaCuti,
      linkGambar: dataCutiRow[9] || ''  // Kolom J untuk link dokumen
    };
  } catch (e) { return { error: e.message }; }
}

function getDashboardData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCuti = ss.getSheetByName('Database Cuti');
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    const dataKaryawan = sheetKaryawan.getRange(2, 1, sheetKaryawan.getLastRow() - 1, 2).getValues();
    const jumlahKaryawan = dataKaryawan.length;
    const namaKaryawanMap = dataKaryawan.reduce((map, row) => { map[row[0]] = row[1]; return map; }, {});
    const hariIni = new Date();
    hariIni.setHours(0, 0, 0, 0);
    const enamBulanLalu = new Date();
    enamBulanLalu.setMonth(enamBulanLalu.getMonth() - 6);
    let pengajuanBaru = 0;
    const sedangCutiSet = new Set();
    let topKaryawanCuti = { nama: 'N/A', jumlah: 0 };
    const frekuensiCuti = {}, trenBulanan = {}, distribusiJenis = {}, cutiMendatang = [];
    if (sheetCuti.getLastRow() < 2) {
      return { jumlahKaryawan, pengajuanBaru, sedangCuti, topKaryawanCuti, trenBulanan, distribusiJenis, cutiMendatang };
    }
    const dataCuti = sheetCuti.getRange(2, 1, sheetCuti.getLastRow() - 1, sheetCuti.getLastColumn()).getValues();
    const toDateOnly = (value) => {
      const d = new Date(value);
      d.setHours(0, 0, 0, 0);
      return d;
    };

    dataCuti.forEach(row => {
      const idKaryawan = row[2], tglPengajuan = new Date(row[1]);
      const tglMulai = toDateOnly(row[3]);
      const tglSelesai = toDateOnly(row[4]);
      const jenisCuti = row[6] || 'Lainnya', status = row[8], jumlahHari = row[5];
      if (status === 'Menunggu Persetujuan') pengajuanBaru++;
      frekuensiCuti[idKaryawan] = (frekuensiCuti[idKaryawan] || 0) + 1;
      if (tglPengajuan >= enamBulanLalu) {
        const bulanKey = tglPengajuan.getFullYear() + '-' + ('0' + (tglPengajuan.getMonth() + 1)).slice(-2);
        trenBulanan[bulanKey] = (trenBulanan[bulanKey] || 0) + 1;
      }
      if (status === 'Disetujui') {
        if (hariIni >= tglMulai && hariIni <= tglSelesai) sedangCutiSet.add(idKaryawan);
        distribusiJenis[jenisCuti] = (distribusiJenis[jenisCuti] || 0) + 1;
        if (tglMulai >= hariIni && cutiMendatang.length < 5) {
          cutiMendatang.push({
            nama: namaKaryawanMap[idKaryawan] || idKaryawan,
            mulai: tglMulai.toLocaleDateString('id-ID', {day: 'numeric', month: 'short'}),
            selesai: tglSelesai.toLocaleDateString('id-ID', {day: 'numeric', month: 'short'}),
            jenis: jenisCuti, lama: jumlahHari
          });
        }
      }
    });
    const sedangCuti = sedangCutiSet.size;
    let topId = null, maxCount = 0;
    for (const id in frekuensiCuti) { if (frekuensiCuti[id] > maxCount) { maxCount = frekuensiCuti[id]; topId = id; } }
    topKaryawanCuti = topId ? { nama: namaKaryawanMap[topId], jumlah: maxCount } : { nama: 'N/A', jumlah: 0 };
    cutiMendatang.sort((a,b) => new Date(a.mulai.split('/').reverse().join('-')) - new Date(b.mulai.split('/').reverse().join('-')));
    return { jumlahKaryawan, pengajuanBaru, sedangCuti, topKaryawanCuti, trenBulanan, distribusiJenis, cutiMendatang };
  } catch (e) { return { error: e.toString() }; }
}

// FUNGSI BARU: Hitung distribusi kategori cuti dengan persentase per karyawan
function getDistribusiKategoriCuti() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCuti = ss.getSheetByName('Database Cuti');
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    
    if (sheetCuti.getLastRow() < 2) {
      return { 
        kategoriData: {
          'Cuti Tahunan': { jumlah: 0, karyawan: 0, persentase: 0 },
          'Cuti Sakit': { jumlah: 0, karyawan: 0, persentase: 0 },
          'Cuti Alasan Penting': { jumlah: 0, karyawan: 0, persentase: 0 }
        },
        totalKaryawan: sheetKaryawan.getLastRow() - 1
      };
    }

    const dataCuti = sheetCuti.getRange(2, 1, sheetCuti.getLastRow() - 1, sheetCuti.getLastColumn()).getValues();
    const totalKaryawan = sheetKaryawan.getLastRow() - 1;

    // Hitung jumlah karyawan yang mengambil cuti per kategori (hanya yang disetujui) - dinamis
    const kategoriMap = {};   // kategori => Set of karyawan
    const kategoriCount = {}; // kategori => total pengajuan disetujui

    dataCuti.forEach(row => {
      const kategori = row[6] || 'Lainnya';
      const status = row[8];
      const idKaryawan = row[2];

      if (status === 'Disetujui') {
        if (!kategoriMap[kategori]) kategoriMap[kategori] = new Set();
        if (!kategoriCount[kategori]) kategoriCount[kategori] = 0;
        kategoriMap[kategori].add(idKaryawan);
        kategoriCount[kategori]++;
      }
    });

    // Hitung persentase
    const kategoriData = {};
    for (const kategori in kategoriMap) {
      const jumlahKaryawanKategori = kategoriMap[kategori].size;
      const persentase = totalKaryawan > 0 ? Math.round((jumlahKaryawanKategori / totalKaryawan) * 100) : 0;
      kategoriData[kategori] = {
        jumlah: kategoriCount[kategori],
        karyawan: jumlahKaryawanKategori,
        persentase: persentase
      };
    }

    return { kategoriData, totalKaryawan };
  } catch (e) { return { error: e.toString() }; }
}

// FUNGSI HELPER: Hitung detail distribusi kategori dengan breakdown per karyawan
function getDetailDistribusiKategoriCuti() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCuti = ss.getSheetByName('Database Cuti');
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    
    const dataCuti = sheetCuti.getRange(2, 1, sheetCuti.getLastRow() - 1, sheetCuti.getLastColumn()).getValues();
    const dataKaryawan = sheetKaryawan.getRange(2, 1, sheetKaryawan.getLastRow() - 1, 2).getValues();
    const namaKaryawanMap = dataKaryawan.reduce((map, row) => { map[row[0]] = row[1]; return map; }, {});
    
    const totalKaryawan = dataKaryawan.length;
    const kategoriDetail = {};
    const karyawanByKategori = {};

    dataCuti.forEach(row => {
      const kategori = row[6] || 'Lainnya';
      const status = row[8];
      const idKaryawan = row[2];
      const namaKaryawan = namaKaryawanMap[idKaryawan] || 'N/A';
      const tglMulai = new Date(row[3]).toLocaleDateString('id-ID');
      const tglSelesai = new Date(row[4]).toLocaleDateString('id-ID');
      const jumlahHari = row[5];

      if (status === 'Disetujui') {
        if (!kategoriDetail[kategori]) {
          kategoriDetail[kategori] = { list: [], total: 0, persentase: 0 };
          karyawanByKategori[kategori] = new Set();
        }
        kategoriDetail[kategori].list.push({
          nama: namaKaryawan,
          id: idKaryawan,
          tanggal: `${tglMulai} - ${tglSelesai}`,
          hari: jumlahHari
        });
        kategoriDetail[kategori].total++;
        karyawanByKategori[kategori].add(idKaryawan);
      }
    });

    // Hitung persentase untuk setiap kategori
    for (const kategori in kategoriDetail) {
      const jumlahKaryawan = karyawanByKategori[kategori].size;
      kategoriDetail[kategori].persentase = totalKaryawan > 0 ? Math.round((jumlahKaryawan / totalKaryawan) * 100) : 0;
      kategoriDetail[kategori].jumlahKaryawan = jumlahKaryawan;
    }

    return { kategoriDetail, totalKaryawan };
  } catch (e) { return { error: e.toString() }; }
}

// FUNGSI BARU: Hitung jumlah karyawan yang sedang cuti per kategori
function getJumlahKaryawanCutiPerKategori() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCuti = ss.getSheetByName('Database Cuti');
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    
    const hariIni = new Date();
    hariIni.setHours(0, 0, 0, 0);
    
    const kategoriData = {};

    const totalKaryawan = sheetKaryawan.getLastRow() - 1;

    if (sheetCuti.getLastRow() < 2) {
      return { kategoriData, totalKaryawan };
    }

    const dataCuti = sheetCuti.getRange(2, 1, sheetCuti.getLastRow() - 1, sheetCuti.getLastColumn()).getValues();
    const toDateOnly = (value) => {
      const d = new Date(value);
      d.setHours(0, 0, 0, 0);
      return d;
    };

    dataCuti.forEach(row => {
      const tglMulai = toDateOnly(row[3]);
      const tglSelesai = toDateOnly(row[4]);
      const kategori = row[6] || 'Lainnya';
      const status = row[8];
      const idKaryawan = row[2];

      // Cek apakah karyawan sedang cuti (tanggal hari ini dalam range tanggal cuti)
      if (status === 'Disetujui' && hariIni >= tglMulai && hariIni <= tglSelesai) {
        if (!kategoriData[kategori]) kategoriData[kategori] = { sedangCuti: 0, karyawan: new Set() };
        kategoriData[kategori].karyawan.add(idKaryawan);
      }
    });

    // Konversi set ke angka
    for (const kategori in kategoriData) {
      const entry = kategoriData[kategori];
      entry.sedangCuti = entry.karyawan ? entry.karyawan.size : 0;
      delete entry.karyawan;
    }

    return { kategoriData, totalKaryawan };
  } catch (e) { return { error: e.toString() }; }
}

// FUNGSI BARU: Ambil rekapan cuti per orang untuk dicetak
function getRekapanCutiPerOrang() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCuti = ss.getSheetByName('Database Cuti');
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    
    const dataCuti = sheetCuti.getRange(2, 1, sheetCuti.getLastRow() - 1, sheetCuti.getLastColumn()).getValues();
    const dataKaryawan = sheetKaryawan.getRange(2, 1, sheetKaryawan.getLastRow() - 1, 5).getValues();
    
    // Buat map nama dan jatah cuti
    const karyawanMap = {};
    dataKaryawan.forEach(row => {
      karyawanMap[row[0]] = {
        nama: row[1],
        email: row[2],
        jatahCuti: row[4]
      };
    });
    
    // Hitung cuti per karyawan per tahun ini
    const tahunIni = new Date().getFullYear();
    const rekapCuti = {};
    
    dataCuti.forEach(row => {
      const idKaryawan = row[2];
      const tglMulai = new Date(row[3]);
      const tglSelesai = new Date(row[4]);
      const tahunCuti = tglMulai.getFullYear();
      const jumlahHari = row[5];
      const kategori = row[6] || 'Lainnya';
      const status = row[8];
      
      // Hanya hitung cuti tahun ini yang sudah disetujui
      if (tahunCuti === tahunIni && status === 'Disetujui') {
        if (!rekapCuti[idKaryawan]) {
          rekapCuti[idKaryawan] = {
            nama: karyawanMap[idKaryawan]?.nama || 'N/A',
            jatahCuti: karyawanMap[idKaryawan]?.jatahCuti || 0,
            totalCutiDiambil: 0,
            riwayatCuti: []
          };
        }
        
        rekapCuti[idKaryawan].totalCutiDiambil += jumlahHari;
        rekapCuti[idKaryawan].riwayatCuti.push({
          tanggalMulai: tglMulai.toLocaleDateString('id-ID'),
          tanggalSelesai: tglSelesai.toLocaleDateString('id-ID'),
          jumlahHari: jumlahHari,
          kategori: kategori
        });
      }
    });
    
    // Ubah ke array dan hitung sisa
    const daftarRekapan = [];
    for (const idKaryawan in rekapCuti) {
      const data = rekapCuti[idKaryawan];
      daftarRekapan.push({
        idKaryawan: idKaryawan,
        nama: data.nama,
        jatahCuti: data.jatahCuti,
        totalCutiDiambil: data.totalCutiDiambil,
        sisaCuti: data.jatahCuti - data.totalCutiDiambil,
        riwayatCuti: data.riwayatCuti
      });
    }
    
    // Sort berdasarkan nama
    daftarRekapan.sort((a, b) => a.nama.localeCompare(b.nama));
    
    return { daftarRekapan, tahun: tahunIni };
  } catch (e) { return { error: e.toString() }; }
}

// FUNGSI: Rekapan cuti detail per karyawan (untuk dashboard admin)
function getRekapanCutiDetail() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCuti = ss.getSheetByName('Database Cuti');
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');

    const tahunIni = new Date().getFullYear();

    // ✨ KATEGORI YANG TIDAK DIHITUNG KE TOTAL CUTI
    const kategoriTidakDihitung = ['Izin Sakit', 'Keluarga Kandung Opname/Rawat inap', 'Pengganti Libur'];

    // Daftar kategori default untuk memastikan kolom selalu muncul meski belum ada data
    const defaultKategori = [
      'Libur',
      'Cuti Pasca Bersalin',
      'Izin Selama Hamil',
      'Istri Melahirkan',
      'Cuti Ibadah Haji',
      'Cuti Umroh',
      'Pernikahan Karyawan',
      'Keluarga Kandung Meninggal',
      'Pernikahan Anak Kandung',
      'Aqikah Anak Kandung',
      'Sunat Anak Kandung',
      'Operasi/Sakit Parah',
      'Izin Sakit',
      'Keluarga Kandung Opname/Rawat inap',
      'Pengganti Libur'
    ];

    // Ambil data karyawan (ID, Nama, Email, TglGabung, Jatah Cuti)
    const dataKaryawan = sheetKaryawan.getRange(2, 1, Math.max(sheetKaryawan.getLastRow() - 1, 0), 5).getValues();
    const karyawanMap = {};
    dataKaryawan.forEach(row => {
      const id = row[0];
      // PERUBAHAN: Ambil jatah cuti tahunan dari fungsi yang mengambil dari kolom J
      // Ini adalah TOTAL jatah cuti tahunan, bukan hanya kategori "Libur"
      const jatahCutiTahunIni = getJatahCutiTahunanKaryawan(id) || 12;
      karyawanMap[id] = {
        nama: row[1] || 'N/A',
        jatahCutiTahunan: jatahCutiTahunIni // GUNAKAN total jatah dari kolom J
      };
      Logger.log('✓ [Rekapan] ID: ' + id + ', Nama: ' + row[1] + ', Jatah Tahunan: ' + jatahCutiTahunIni);
    });

    if (sheetCuti.getLastRow() < 2) {
      return { daftarRekapan: [], tahun: tahunIni, allKategori: [] };
    }

    const dataCuti = sheetCuti.getRange(2, 1, sheetCuti.getLastRow() - 1, sheetCuti.getLastColumn()).getValues();

    const allowedKategori = new Set(defaultKategori);
    const rekapan = {};

    dataCuti.forEach(row => {
      const idKaryawan = row[2];
      const tglMulai = new Date(row[3]);
      const tahunCuti = tglMulai.getFullYear();
      const jumlahHari = row[5];
      let kategori = row[6] || 'Lainnya';
      const status = row[8];

      // MAPPING: Ubah "Cuti Tahunan" menjadi "Libur"
      if (kategori === 'Cuti Tahunan') {
        kategori = 'Libur';
      }

      // Hitung hanya cuti yang disetujui di tahun berjalan
      if (status !== 'Disetujui' || tahunCuti !== tahunIni) return;

      // Abaikan kategori yang tidak ada di daftar allowed
      if (!allowedKategori.has(kategori)) return;

      if (!rekapan[idKaryawan]) {
        rekapan[idKaryawan] = {
          idKaryawan,
          nama: karyawanMap[idKaryawan]?.nama || 'N/A',
          jatahCutiTahunan: karyawanMap[idKaryawan]?.jatahCutiTahunan || 12, // DARI JATAH BERDASARKAN TAHUN
          totalCutiDiambil: 0,
          totalCutiDiambildenganPengecualian: 0, // ✨ Track total ADA cuti (termasuk Izin Sakit & Pengganti Libur)
          totalKategoriDiambil: 0, // ✨ Track JUMLAH KATEGORI yang diambil (bukan jumlah hari)
          kategoriYangDiambil: new Set(), // ✨ Track kategori mana saja yang sudah diambil
          kategoriCount: {}
        };
      }

      // Tambahkan hitungan per kategori (gunakan jumlah hari)
      rekapan[idKaryawan].kategoriCount[kategori] = (rekapan[idKaryawan].kategoriCount[kategori] || 0) + (jumlahHari || 0);
      
      // ✨ TRACK KATEGORI YANG DIAMBIL (untuk menghitung jumlah kategori)
      if (!rekapan[idKaryawan].kategoriYangDiambil.has(kategori)) {
        rekapan[idKaryawan].kategoriYangDiambil.add(kategori);
        rekapan[idKaryawan].totalKategoriDiambil += 1; // INCREMENT jumlah kategori
      }
      
      // ✨ TRACK TOTAL ADA CUTI (untuk pengecekan apakah karyawan punya cuti atau tidak)
      rekapan[idKaryawan].totalCutiDiambildenganPengecualian += jumlahHari || 0;
      
      // ✨ PENTING: Hanya tambahkan ke totalCutiDiambil jika BUKAN kategori yang dikecualikan
      if (!kategoriTidakDihitung.includes(kategori)) {
        rekapan[idKaryawan].totalCutiDiambil += jumlahHari || 0;
      } else {
        Logger.log('⚠ [Rekapan] Kategori "' + kategori + '" TIDAK dihitung ke Total untuk ' + rekapan[idKaryawan].nama);
      }
    });

    const allKategori = [...defaultKategori];

    // Pastikan setiap karyawan punya entry 0 untuk semua kategori yang ada di kolom
    const daftarRekapan = Object.values(rekapan)
      .filter(item => item.totalCutiDiambildenganPengecualian > 0) // ✨ Filter berdasarkan ADA cuti (termasuk Izin Sakit & Pengganti Libur)
      .map(item => {
        const kategoriCountNormalized = { ...item.kategoriCount };
        allKategori.forEach(k => {
          if (kategoriCountNormalized[k] === undefined) kategoriCountNormalized[k] = 0;
        });
        // ✨ Hitung sisa berdasarkan Jatah Awal (tetap tampilkan Jatah Awal di tabel)
        const jatahAwal = (item.jatahCutiTahunan || 12);
        const sisaCutiHitung = jatahAwal - (item.totalCutiDiambil || 0);

        Logger.log('✓ [Rekapan] ' + item.nama + ' | Jatah Awal: ' + jatahAwal + ' - Diambil: ' + (item.totalCutiDiambil || 0) + ' = Sisa: ' + sisaCutiHitung);
        return {
          ...item,
          kategoriCount: kategoriCountNormalized,
          // ✨ TAMPILKAN JATAH AWAL (tidak dikurangi)
          jatahCutiTahunan: jatahAwal,
          // ✨ SISA CUTI dihitung terpisah (Jatah Awal - totalCutiDiambil)
          sisaCuti: sisaCutiHitung,
          // ✨ DISPLAY: Jumlah KATEGORI yang diambil (bukan jumlah hari)
          totalCutiDisplay: item.totalKategoriDiambil
        };
      })
      .sort((a, b) => a.nama.localeCompare(b.nama));

    return { daftarRekapan, tahun: tahunIni, allKategori };
  } catch (e) { return { error: e.toString() }; }
}

// FUNGSI BARU: Tulis sisa cuti ke spreadsheet (kolom "Jml Cuti Pertahun")
function simpanSisaCutiKeSpreadsheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tahunIni = new Date().getFullYear();
    
    // ✨ Ambil data rekapan terbaru (SAMA dengan di Admin Dashboard)
    const dataRekapan = getRekapanCutiDetail();
    if (dataRekapan.error) {
      return { sukses: false, pesan: dataRekapan.error };
    }
    
    const { daftarRekapan, allKategori, tahun } = dataRekapan;
    
    // ✨ Cari atau buat sheet tracking sisa cuti (terpisah dari export sheet)
    const sheetName = `Sisa Cuti ${tahun}`;
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      // Buat sheet baru khusus untuk tracking sisa cuti
      sheet = ss.insertSheet(sheetName);
      Logger.log('✓ [Simpan Sisa Cuti] Sheet baru "' + sheetName + '" telah dibuat');
    } else {
      // ✨ Clear data lama (except header) dan ganti dengan data baru
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.deleteRows(2, lastRow - 1);
      }
      Logger.log('✓ [Simpan Sisa Cuti] Sheet "' + sheetName + '" telah dikosongkan dan dipersiapkan');
    }
    
    // ✨ Buat header SAMA dengan tabel rekapan di Admin Dashboard
    const header = ['No.', 'Nama Karyawan', ...allKategori, 'Total Cuti', 'Jatah Tahunan', 'Sisa Cuti'];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
    
    // Styling header
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f5f5f5');
    sheet.setFrozenRows(1);
    
    // ✨ Populate data ke sheet (SAMA dengan Admin)
    const rows = (daftarRekapan || []).map((item, idx) => {
      const kategoriCounts = allKategori.map(k => item.kategoriCount?.[k] || 0);
      return [
        idx + 1,
        item.nama || 'N/A',
        ...kategoriCounts,
        item.totalCutiDisplay || 0, // ✨ Jumlah KATEGORI yang diambil
        item.jatahCutiTahunan || 0, // ✨ Jatah yang BERKURANG
        item.sisaCuti || 0            // ✨ Sisa Cuti
      ];
    });
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
      Logger.log('✓ [Simpan Sisa Cuti] ' + rows.length + ' baris data telah ditulis ke sheet');
      
      // Auto-resize kolom
      sheet.autoResizeColumns(1, header.length);
    }
    
    Logger.log('✓ [Simpan Sisa Cuti] Proses selesai - Sheet "' + sheetName + '" ter-update dengan ' + rows.length + ' karyawan');
    
    return { 
      sukses: true, 
      pesan: `✓ Berhasil simpan sisa cuti ke sheet "${sheetName}"\n\nTotal ${rows.length} karyawan ter-update dengan data terbaru.`,
      sheetName: sheetName,
      jumlahDiupdate: rows.length,
      tahun: tahun
    };
    
  } catch (e) {
    Logger.log('❌ Error di simpanSisaCutiKeSpreadsheet: ' + e.message);
    return { sukses: false, pesan: '❌ Error: ' + e.message };
  }
}

// Export rekapan cuti ke sheet baru (untuk referensi di Drive)
function exportRekapanCutiToSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const result = getRekapanCutiDetail();

    if (result.error) {
      return { error: result.error };
    }

    const { daftarRekapan, allKategori, tahun } = result;

    // Siapkan nama sheet, hindari duplikasi
    const baseName = `Rekapan Cuti ${tahun}`;
    let sheetName = baseName;
    let counter = 1;
    while (ss.getSheetByName(sheetName)) {
      counter += 1;
      sheetName = `${baseName} (${counter})`;
    }

    const sheet = ss.insertSheet(sheetName);

    // ✨ Header SESUAI dengan tabel di Admin: No. | Nama | [Kategori] | Total Cuti | Jatah Tahunan | Sisa Cuti
    const header = ['No.', 'Nama Karyawan', ...allKategori, 'Total Cuti', 'Jatah Tahunan', 'Sisa Cuti'];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);

    // ✨ Data baris SESUAI dengan display: gunakan totalCutiDisplay, jatahCutiTahunan (yang berkurang), sisaCuti
    const rows = (daftarRekapan || []).map((item, idx) => {
      const kategoriCounts = allKategori.map(k => item.kategoriCount?.[k] || 0);
      return [
        idx + 1,
        item.nama || 'N/A',
        ...kategoriCounts,
        item.totalCutiDisplay || 0, // ✨ Jumlah KATEGORI yang diambil
        item.jatahCutiTahunan || 0, // ✨ Jatah yang BERKURANG
        item.sisaCuti || 0            // ✨ Sisa Cuti
      ];
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    }

    // Styling ringan
    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f5f5f5');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, header.length);

    return { message: `Berhasil export ke sheet "${sheetName}".`, sheetName, url: ss.getUrl() };
  } catch (e) {
    return { error: e.toString() };
  }
}

// ✨ Export rekapan cuti ke Excel (.xlsx) dan kembalikan link download
function exportRekapanCutiToExcel() {
  try {
    const recap = getRekapanCutiDetail();
    if (recap.error) return { error: recap.error };

    const { daftarRekapan, allKategori, tahun } = recap;

    // Buat spreadsheet baru khusus export
    const exportSs = SpreadsheetApp.create(`Rekapan Cuti ${tahun} (Export)`);
    const sheet = exportSs.getActiveSheet();

    // ✨ Header SESUAI dengan tabel di Admin: No. | Nama | [Kategori] | Total Cuti | Jatah Tahunan | Sisa Cuti
    const header = ['No.', 'Nama Karyawan', ...allKategori, 'Total Cuti', 'Jatah Tahunan', 'Sisa Cuti'];
    sheet.getRange(1, 1, 1, header.length).setValues([header]);

    // ✨ Data baris SESUAI dengan display: gunakan totalCutiDisplay, jatahCutiTahunan (yang berkurang), sisaCuti
    const rows = (daftarRekapan || []).map((item, idx) => {
      const kategoriCounts = allKategori.map(k => item.kategoriCount?.[k] || 0);
      return [
        idx + 1,
        item.nama || 'N/A',
        ...kategoriCounts,
        item.totalCutiDisplay || 0, // ✨ Jumlah KATEGORI yang diambil
        item.jatahCutiTahunan || 0, // ✨ Jatah yang BERKURANG
        item.sisaCuti || 0            // ✨ Sisa Cuti
      ];
    });

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, header.length).setValues(rows);
    }

    sheet.getRange(1, 1, 1, header.length).setFontWeight('bold').setBackground('#f5f5f5');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, header.length);

    const fileId = exportSs.getId();
    const downloadUrl = `https://docs.google.com/spreadsheets/d/${fileId}/export?format=xlsx`;

    return {
      message: 'File Excel siap diunduh.',
      downloadUrl,
      fileId
    };
  } catch (e) {
    return { error: e.toString() };
  }
}

// --- FUNGSI HELPER ---

function buatHashPasswordManual() {
  const passwordAsli = "erwin123"; 
  const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, passwordAsli + SECRET_KEY)
                     .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
  Logger.log(hash);
}

function buatHashMassal() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Karyawan');
  const range = sheet.getRange("A2:H" + sheet.getLastRow());
  const values = range.getValues();
  let hashDibuat = 0;
  for (let i = 0; i < values.length; i++) {
    const barisSaatIni = i + 2;
    const passwordSementara = values[i][7];
    const hashDisimpan = values[i][6];
    if (passwordSementara && !hashDisimpan) {
      const hash = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, passwordSementara + SECRET_KEY)
                         .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
      sheet.getRange(barisSaatIni, 7).setValue(hash);
      sheet.getRange(barisSaatIni, 8).clearContent();
      hashDibuat++;
    }
  }
  Browser.msgBox(`Proses Selesai. Sebanyak ${hashDibuat} password hash baru telah dibuat.`);
}

function ubahPasswordKaryawan(data, token) {
  try {
    const email = getEmailFromToken(token);
    if (!email) throw new Error("Sesi tidak valid.");

    const { passwordLama, passwordBaru } = data;
    if (!passwordLama || !passwordBaru) throw new Error("Semua field wajib diisi.");
    if (passwordBaru.length < 6) throw new Error("Password baru minimal 6 karakter.");

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Karyawan');
    const dataKaryawan = sheet.getRange("A2:I" + sheet.getLastRow()).getValues();

    let karyawanDitemukan = null;
    let rowIndex = -1;

    for (let i = 0; i < dataKaryawan.length; i++) {
      if (dataKaryawan[i][2].toString().trim().toLowerCase() === email.toLowerCase()) {
        karyawanDitemukan = { hashDisimpan: dataKaryawan[i][6] };
        rowIndex = i + 2; 
        break;
      }
    }

    if (!karyawanDitemukan) throw new Error("Karyawan tidak ditemukan.");

    const hashPasswordLama = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, passwordLama + SECRET_KEY)
                               .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');
    
    if (hashPasswordLama !== karyawanDitemukan.hashDisimpan) {
      throw new Error("Password lama yang Anda masukkan salah.");
    }

    const hashPasswordBaru = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, passwordBaru + SECRET_KEY)
                               .map(byte => ('0' + (byte & 0xFF).toString(16)).slice(-2)).join('');

    sheet.getRange(rowIndex, 7).setValue(hashPasswordBaru); // Kolom G adalah PasswordHash

    return "Password berhasil diubah. Silakan login kembali dengan password baru Anda.";

  } catch (e) {
    return "Error: " + e.message;
  }
}

// FUNGSI BARU: Ambil kategori cuti tersedia berdasarkan gender
function getKategoriCutiAvailable(token) {
  try {
    const email = getEmailFromToken(token);
    if (!email) throw new Error("Sesi tidak valid atau telah berakhir. Silakan login kembali.");
    
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Karyawan');
    const dataKaryawan = sheet.getRange("A2:H" + sheet.getLastRow()).getValues();
    
    let jenisKelamin = null;
    for (const row of dataKaryawan) {
      if (row[2].toString().trim().toLowerCase() === email.toLowerCase()) {
        jenisKelamin = row[7];
        break;
      }
    }
    
    if (!jenisKelamin) throw new Error("Jenis kelamin tidak ditemukan.");
    
    const kategoriAvailable = getKategoriCutiByGender(jenisKelamin);
    return { kategori: kategoriAvailable, jenisKelamin: jenisKelamin };
  } catch (e) {
    return { error: e.message };
  }
}

// FUNGSI BARU: Hitung jumlah orang yang mengajukan cuti (Tahun ini)
function getJumlahPengajuanCutiKaryawan() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCuti = ss.getSheetByName('Database Cuti');
    
    if (sheetCuti.getLastRow() < 2) {
      return 0;
    }
    
    const dataCuti = sheetCuti.getRange(2, 1, sheetCuti.getLastRow() - 1, sheetCuti.getLastColumn()).getValues();
    const tahunIni = new Date().getFullYear();
    const karyawanYangMengajukan = new Set();
    
    dataCuti.forEach(row => {
      const tahunPengajuan = new Date(row[1]).getFullYear(); // Kolom B: Tanggal Pengajuan
      if (tahunPengajuan === tahunIni) {
        const idKaryawan = row[2]; // Kolom C: ID Karyawan
        karyawanYangMengajukan.add(idKaryawan);
      }
    });
    
    return karyawanYangMengajukan.size;
  } catch (e) {
    return 0;
  }
}

// FUNGSI BARU: Ambil Libur Kerja yang sedang berlaku untuk karyawan
function getLiburKerjaAktif(idKaryawan) {
  try {
    console.log('getLiburKerjaAktif called for:', idKaryawan);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetLibur = null;
    
    try {
      sheetLibur = ss.getSheetByName('Libur Kerja');
    } catch (e) {
      console.log('Sheet Libur Kerja tidak ditemukan');
      return { liburAktif: null };
    }
    
    if (!sheetLibur || sheetLibur.getLastRow() < 2) {
      console.log('Sheet Libur Kerja kosong');
      return { liburAktif: null };
    }
    
    const dataLibur = sheetLibur.getRange(2, 1, sheetLibur.getLastRow() - 1, sheetLibur.getLastColumn()).getValues();
    const hariIni = new Date();
    hariIni.setHours(0, 0, 0, 0);
    
    for (const row of dataLibur) {
      // Sheet: A=ID, B=ID Karyawan, C=Nama Karyawan, D=Tgl Mulai, E=Tgl Selesai, F=Durasi, G=Deskripsi, H=Status
      const idKaryawanLibur = row[1];
      const tglMulai = new Date(row[3]);
      const tglSelesai = new Date(row[4]);
      const durasi = row[5];
      const deskripsi = row[6];
      const status = (row[7] || '').toString().trim().toLowerCase();
      
      tglMulai.setHours(0, 0, 0, 0);
      tglSelesai.setHours(0, 0, 0, 0);
      
      // Cek apakah Libur Kerja untuk karyawan ini, aktif, dan berada dalam periode
      if (idKaryawanLibur === idKaryawan && status === 'active' && hariIni >= tglMulai && hariIni <= tglSelesai) {
        const liburAktif = {
          id: row[0],
          tglMulai: tglMulai.toLocaleDateString('id-ID'),
          tglSelesai: tglSelesai.toLocaleDateString('id-ID'),
          durasi: parseInt(row[5]),
          deskripsi: deskripsi
        };
        console.log('Libur Kerja aktif ditemukan:', liburAktif);
        return { liburAktif: liburAktif };
      }
    }
    
    console.log('Tidak ada Libur Kerja yang aktif untuk periode ini');
    return { liburAktif: null };
  } catch (e) {
    Logger.log('Error in getLiburKerjaAktif: ' + e.message);
    console.error('Error in getLiburKerjaAktif:', e);
    return { liburAktif: null };
  }
}

// FUNGSI BARU: Ambil semua data Libur Kerja
function getAllLiburKerja() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetLibur = null;
    
    try {
      sheetLibur = ss.getSheetByName('Libur Kerja');
    } catch (e) {
      // Sheet belum ada, buat sheet baru dengan kolom baru
      sheetLibur = ss.insertSheet('Libur Kerja');
      sheetLibur.appendRow(['ID', 'ID Karyawan', 'Nama Karyawan', 'Tanggal Mulai', 'Tanggal Selesai', 'Durasi (Hari)', 'Deskripsi', 'Status']);
      return [];
    }
    
    if (sheetLibur.getLastRow() < 2) {
      return [];
    }
    
    const dataLibur = sheetLibur.getRange(2, 1, sheetLibur.getLastRow() - 1, sheetLibur.getLastColumn()).getValues();
    const result = dataLibur.map(row => ({
      id: row[0],
      idKaryawan: row[1],
      namakaryawan: row[2],
      tglMulai: row[3],
      tglSelesai: row[4],
      durasi: row[5],
      deskripsi: row[6],
      status: (row[7] || 'inactive').toString().trim().toLowerCase()
    }));
    
    return result;
  } catch (e) {
    Logger.log('Error in getAllLiburKerja: ' + e.message);
    return [];
  }
}

// FUNGSI BARU: Simpan Libur Kerja baru
function simpanLiburKerja(dataLibur) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetLibur = null;
    
    try {
      sheetLibur = ss.getSheetByName('Libur Kerja');
    } catch (e) {
      sheetLibur = ss.insertSheet('Libur Kerja');
      sheetLibur.appendRow(['ID', 'ID Karyawan', 'Nama Karyawan', 'Tanggal Mulai', 'Tanggal Selesai', 'Durasi (Hari)', 'Deskripsi', 'Status']);
    }
    
    // Ambil nama karyawan dari sheet Data Karyawan
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    const dataKaryawan = sheetKaryawan.getRange(2, 1, sheetKaryawan.getLastRow() - 1, 2).getValues();
    let namaKaryawan = '';
    for (const k of dataKaryawan) {
      if (k[0] === dataLibur.idKaryawan) {
        namaKaryawan = k[1];
        break;
      }
    }
    
    // Generate ID otomatis
    const idLibur = 'LIBUR-' + new Date().getTime();
    
    sheetLibur.appendRow([
      idLibur,
      dataLibur.idKaryawan,
      namaKaryawan,
      new Date(dataLibur.tglMulai),
      new Date(dataLibur.tglSelesai),
      dataLibur.durasi,
      dataLibur.deskripsi,
      dataLibur.status
    ]);
    
    return `Libur Kerja "${dataLibur.deskripsi}" untuk ${namaKaryawan} berhasil ditambahkan!`;
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

// FUNGSI BARU: Hapus Libur Kerja
function hapusLiburKerja(id) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetLibur = ss.getSheetByName('Libur Kerja');
    
    const ids = sheetLibur.getRange("A2:A").getValues().flat();
    const rowIndex = ids.indexOf(id);
    
    if (rowIndex === -1) {
      throw new Error("Libur Kerja dengan ID " + id + " tidak ditemukan.");
    }
    
    sheetLibur.deleteRow(rowIndex + 2);
    return 'Libur Kerja berhasil dihapus!';
  } catch (e) {
    return 'Error: ' + e.message;
  }
}

// ============ FUNGSI PENGGANTI LIBUR (NEW) ============

// FUNGSI BARU: Ambil semua data Pengganti Libur
function getAllPenggantiLibur() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetPengganti = null;
    
    try {
      sheetPengganti = ss.getSheetByName('Pengganti Libur');
    } catch (e) {
      // Sheet belum ada, buat sheet baru dengan struktur baru
      sheetPengganti = ss.insertSheet('Pengganti Libur');
      sheetPengganti.appendRow(['ID Karyawan', 'Nama Karyawan', 'Durasi Total (Hari)', 'Hari Digunakan', 'Sisa Hari', 'Keterangan']);
      Logger.log('Created new Pengganti Libur sheet with new structure');
      return [];
    }
    
    // Cek dan upgrade struktur sheet jika perlu
    const headers = sheetPengganti.getRange(1, 1, 1, sheetPengganti.getLastColumn()).getValues()[0];
    if (headers.length < 6) {
      // Upgrade sheet: tambah kolom baru jika belum ada
      if (headers.length === 4) {
        sheetPengganti.insertColumns(4, 2);
        sheetPengganti.getRange(1, 4).setValue('Hari Digunakan');
        sheetPengganti.getRange(1, 5).setValue('Sisa Hari');
        // Update header kolom 3
        sheetPengganti.getRange(1, 3).setValue('Durasi Total (Hari)');
        Logger.log('Upgraded Pengganti Libur sheet structure');
      }
    }
    
    if (sheetPengganti.getLastRow() < 2) {
      Logger.log('Pengganti Libur sheet is empty (only header)');
      return [];
    }
    
    const dataPengganti = sheetPengganti.getRange(2, 1, sheetPengganti.getLastRow() - 1, 6).getValues();
    const result = dataPengganti.map(row => ({
      idKaryawan: row[0],
      namaKaryawan: row[1],
      durasiTotal: parseInt(row[2]) || 0,
      hariDigunakan: parseInt(row[3]) || 0,
      sisaHari: parseInt(row[4]) || 0,
      keterangan: row[5] || ''
    }));
    
    Logger.log('getAllPenggantiLibur returning ' + result.length + ' records');
    Logger.log('Data: ' + JSON.stringify(result));
    
    return result;
  } catch (e) {
    Logger.log('Error in getAllPenggantiLibur: ' + e.message + ' | Stack: ' + e.stack);
    return [];
  }
}

// FUNGSI BARU: Simpan Pengganti Libur (VERSI SIMPLIFIED)
function simpanPenggantiLibur(dataPengganti) {
  try {
    Logger.log('=== simpanPenggantiLibur MULAI ===');
    Logger.log('Input data:', JSON.stringify(dataPengganti));
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Logger.log('Spreadsheet aktif didapat');
    
    // ===== STEP 1: Validasi input =====
    if (!dataPengganti || !dataPengganti.idKaryawan || !dataPengganti.durasi) {
      const msg = 'Error: Data tidak lengkap. idKaryawan=' + dataPengganti.idKaryawan + ', durasi=' + dataPengganti.durasi;
      Logger.log(msg);
      return msg;
    }
    Logger.log('✓ Input validation OK');
    
    // ===== STEP 2: Ambil sheet Data Karyawan =====
    let sheetKaryawan = null;
    try {
      sheetKaryawan = ss.getSheetByName('Data Karyawan');
      Logger.log('✓ Berhasil get sheet "Data Karyawan"');
    } catch (e) {
      Logger.log('✗ GAGAL get sheet "Data Karyawan": ' + e.message);
      return 'Error: Sheet "Data Karyawan" tidak ditemukan. Pastikan nama sheet PERSIS "Data Karyawan"';
    }
    
    if (!sheetKaryawan) {
      Logger.log('✗ sheetKaryawan adalah NULL');
      return 'Error: Sheet "Data Karyawan" NULL. Pastikan sheet ada di spreadsheet.';
    }
    Logger.log('✓ sheetKaryawan bukan NULL');
    
    // ===== STEP 3: Cek apakah sheet punya data =====
    let lastRowKaryawan = 0;
    try {
      lastRowKaryawan = sheetKaryawan.getLastRow();
      Logger.log('Last row di Data Karyawan: ' + lastRowKaryawan);
    } catch (e) {
      Logger.log('✗ GAGAL getLastRow: ' + e.message);
      return 'Error: Gagal baca Data Karyawan - ' + e.message;
    }
    
    if (lastRowKaryawan < 2) {
      Logger.log('✗ Data Karyawan kosong atau hanya header');
      return 'Error: Sheet "Data Karyawan" tidak punya data (kurang dari 2 rows)';
    }
    Logger.log('✓ Data Karyawan punya data');
    
    // ===== STEP 4: Ambil data karyawan dari sheet =====
    let dataKaryawan = [];
    try {
      dataKaryawan = sheetKaryawan.getRange(2, 1, lastRowKaryawan - 1, 2).getValues();
      Logger.log('✓ Data karyawan retrieved: ' + dataKaryawan.length + ' rows');
    } catch (e) {
      Logger.log('✗ GAGAL read range: ' + e.message);
      return 'Error: Gagal baca data karyawan - ' + e.message;
    }
    
    // ===== STEP 5: Cari nama karyawan =====
    let namaKaryawan = '';
    const idToFind = String(dataPengganti.idKaryawan).trim();
    Logger.log('Mencari ID: "' + idToFind + '"');
    
    for (let i = 0; i < dataKaryawan.length; i++) {
      const sheetId = String(dataKaryawan[i][0]).trim();
      const sheetNama = String(dataKaryawan[i][1]).trim();
      
      if (sheetId === idToFind) {
        namaKaryawan = sheetNama;
        Logger.log('✓ DITEMUKAN di row ' + (i+2) + ': ' + sheetNama);
        break;
      }
    }
    
    if (!namaKaryawan) {
      Logger.log('✗ Karyawan dengan ID "' + idToFind + '" NOT FOUND di sheet');
      return 'Error: ID Karyawan "' + idToFind + '" tidak ditemukan di Data Karyawan';
    }
    Logger.log('✓ Nama karyawan ditemukan: ' + namaKaryawan);
    
    // ===== STEP 6: Ambil atau buat sheet Pengganti Libur =====
    let sheetPengganti = null;
    try {
      sheetPengganti = ss.getSheetByName('Pengganti Libur');
      if (!sheetPengganti) {
        Logger.log('Sheet "Pengganti Libur" belum ada, membuat baru...');
        sheetPengganti = ss.insertSheet('Pengganti Libur');
        sheetPengganti.appendRow(['ID Karyawan', 'Nama Karyawan', 'Durasi Total (Hari)', 'Hari Digunakan', 'Sisa Hari', 'Keterangan']);
        Logger.log('✓ Sheet "Pengganti Libur" berhasil dibuat dengan struktur baru');
      } else {
        Logger.log('✓ Sheet "Pengganti Libur" sudah ada');
      }
    } catch (e) {
      Logger.log('✗ GAGAL manage sheet Pengganti Libur: ' + e.message);
      return 'Error: Gagal akses sheet Pengganti Libur - ' + e.message;
    }
    
    // ===== STEP 7: Cek apakah sudah ada record untuk karyawan ini =====
    let existingRowIndex = -1;
    const lastRowPengganti = sheetPengganti.getLastRow();
    Logger.log('Last row di Pengganti Libur: ' + lastRowPengganti);
    
    if (lastRowPengganti > 1) {
      try {
        const ids = sheetPengganti.getRange("A2:A" + lastRowPengganti).getValues().flat();
        Logger.log('Checking ' + ids.length + ' existing records');
        
        for (let i = 0; i < ids.length; i++) {
          const existingId = String(ids[i]).trim();
          if (existingId === idToFind) {
            existingRowIndex = i;
            Logger.log('✓ Record sudah ada di row ' + (i+2));
            break;
          }
        }
      } catch (e) {
        Logger.log('⚠ Warning saat check existing: ' + e.message);
      }
    }
    
    // ===== STEP 8: Simpan data (UPDATE atau INSERT) =====
    if (existingRowIndex !== -1) {
      // UPDATE
      try {
        const hariDigunakan = dataPengganti.hariDigunakan || 0;
        const sisaHari = dataPengganti.durasi - hariDigunakan;
        sheetPengganti.getRange(existingRowIndex + 2, 1, 1, 6).setValues([[
          dataPengganti.idKaryawan,
          namaKaryawan,
          dataPengganti.durasi,
          hariDigunakan,
          sisaHari,
          dataPengganti.keterangan || ''
        ]]);
        Logger.log('✓ BERHASIL UPDATE record untuk ' + namaKaryawan);
        return 'Pengganti Libur untuk ' + namaKaryawan + ' berhasil diperbarui!';
      } catch (e) {
        Logger.log('✗ GAGAL UPDATE: ' + e.message);
        return 'Error: Gagal update - ' + e.message;
      }
    } else {
      // INSERT
      try {
        const hariDigunakan = dataPengganti.hariDigunakan || 0;
        const sisaHari = dataPengganti.durasi - hariDigunakan;
        sheetPengganti.appendRow([
          dataPengganti.idKaryawan,
          namaKaryawan,
          dataPengganti.durasi,
          hariDigunakan,
          sisaHari,
          dataPengganti.keterangan || ''
        ]);
        Logger.log('✓ BERHASIL INSERT record untuk ' + namaKaryawan);
        return 'Pengganti Libur untuk ' + namaKaryawan + ' (' + dataPengganti.durasi + ' hari) berhasil ditambahkan!';
      } catch (e) {
        Logger.log('✗ GAGAL INSERT: ' + e.message);
        return 'Error: Gagal simpan - ' + e.message;
      }
    }
    
  } catch (e) {
    Logger.log('✗ ERROR UTAMA: ' + e.message + ' | ' + e.stack);
    return 'Error: ' + e.message;
  }
}

// FUNGSI BARU: Hapus Pengganti Libur
function hapusPenggantiLibur(idKaryawan) {
  try {
    Logger.log('hapusPenggantiLibur called for ID: ' + idKaryawan);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetPengganti = null;
    
    try {
      sheetPengganti = ss.getSheetByName('Pengganti Libur');
      if (!sheetPengganti) {
        Logger.log('Pengganti Libur sheet not found');
        return 'Error: Sheet "Pengganti Libur" tidak ditemukan';
      }
    } catch (e) {
      Logger.log('Error finding Pengganti Libur sheet: ' + e.message);
      return 'Error: Gagal mengakses sheet Pengganti Libur - ' + e.message;
    }
    
    const lastRow = sheetPengganti.getLastRow();
    Logger.log('Pengganti Libur sheet has ' + lastRow + ' rows');
    
    if (lastRow < 2) {
      return 'Error: Data Pengganti Libur kosong';
    }
    
    try {
      const ids = sheetPengganti.getRange("A2:A" + lastRow).getValues().flat();
      const idToFind = String(idKaryawan).trim();
      let rowIndex = -1;
      
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i]).trim() === idToFind) {
          rowIndex = i;
          break;
        }
      }
      
      if (rowIndex === -1) {
        Logger.log('Pengganti Libur not found for ID: ' + idKaryawan);
        return 'Error: Pengganti Libur untuk karyawan ini tidak ditemukan.';
      }
      
      sheetPengganti.deleteRow(rowIndex + 2);
      Logger.log('Deleted row ' + (rowIndex + 2));
      return 'Pengganti Libur berhasil dihapus!';
    } catch (e) {
      Logger.log('Error deleting row: ' + e.message);
      return 'Error: Gagal menghapus - ' + e.message;
    }
  } catch (e) {
    Logger.log('Error in hapusPenggantiLibur: ' + e.message);
    return 'Error: ' + e.message;
  }
}

// FUNGSI BARU: Update Pengganti Libur
function updatePenggantiLibur(idKaryawan, durasi, keterangan) {
  try {
    Logger.log('updatePenggantiLibur called for ID: ' + idKaryawan);
    Logger.log('New values - durasi: ' + durasi + ', keterangan: ' + keterangan);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetPengganti = null;
    
    try {
      sheetPengganti = ss.getSheetByName('Pengganti Libur');
      if (!sheetPengganti) {
        Logger.log('Pengganti Libur sheet not found');
        return 'Error: Sheet "Pengganti Libur" tidak ditemukan';
      }
    } catch (e) {
      Logger.log('Error finding Pengganti Libur sheet: ' + e.message);
      return 'Error: Gagal mengakses sheet Pengganti Libur - ' + e.message;
    }
    
    const lastRow = sheetPengganti.getLastRow();
    Logger.log('Pengganti Libur sheet has ' + lastRow + ' rows');
    
    if (lastRow < 2) {
      return 'Error: Data Pengganti Libur kosong';
    }
    
    try {
      const ids = sheetPengganti.getRange("A2:A" + lastRow).getValues().flat();
      const idToFind = String(idKaryawan).trim();
      let rowIndex = -1;
      
      for (let i = 0; i < ids.length; i++) {
        if (String(ids[i]).trim() === idToFind) {
          rowIndex = i;
          break;
        }
      }
      
      if (rowIndex === -1) {
        Logger.log('Pengganti Libur not found for ID: ' + idKaryawan);
        return 'Error: Pengganti Libur untuk karyawan ini tidak ditemukan.';
      }
      
      // Update row: column C (durasi) dan D (keterangan)
      sheetPengganti.getRange(rowIndex + 2, 3).setValue(durasi);           // Kolom C: Durasi
      sheetPengganti.getRange(rowIndex + 2, 4).setValue(keterangan);       // Kolom D: Keterangan
      
      Logger.log('Updated row ' + (rowIndex + 2) + ' - durasi: ' + durasi + ', keterangan: ' + keterangan);
      return 'Pengganti Libur berhasil diperbarui!';
    } catch (e) {
      Logger.log('Error updating row: ' + e.message);
      return 'Error: Gagal memperbarui - ' + e.message;
    }
  } catch (e) {
    Logger.log('Error in updatePenggantiLibur: ' + e.message);
    return 'Error: ' + e.message;
  }
}

// FUNGSI BARU: Ambil data Pengganti Libur untuk karyawan tertentu
function getPenggantiLiburKaryawan(tokenOrId) {
  try {
    // Support both token (dari karyawan) dan idKaryawan (dari admin)
    let idKaryawan = tokenOrId;
    
    // Jika parameter adalah token, ekstrak idKaryawan dari data karyawan
    if (tokenOrId && tokenOrId.length > 50) {
      const email = getEmailFromToken(tokenOrId);
      if (!email) {
        Logger.log('Invalid token in getPenggantiLiburKaryawan');
        return { pengganti: null };
      }
      
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const sheetKaryawan = ss.getSheetByName('Data Karyawan');
      const dataKaryawanRange = sheetKaryawan.getRange("A2:I" + sheetKaryawan.getLastRow()).getValues();
      
      for (const row of dataKaryawanRange) {
        if (row[2].toString().trim().toLowerCase() === email.toLowerCase()) {
          idKaryawan = row[0];
          break;
        }
      }
    }
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetPengganti = null;
    
    try {
      sheetPengganti = ss.getSheetByName('Pengganti Libur');
    } catch (e) {
      return { pengganti: null };
    }
    
    if (!sheetPengganti || sheetPengganti.getLastRow() < 2) {
      return { pengganti: null };
    }
    
    const dataPengganti = sheetPengganti.getRange(2, 1, sheetPengganti.getLastRow() - 1, 6).getValues();
    
    for (const row of dataPengganti) {
      if (row[0] === idKaryawan) {
        // Hitung sisa hari jika belum ada
        let sisaHari = parseInt(row[4]) || 0;
        if (sisaHari === 0) {
          sisaHari = parseInt(row[2]) - (parseInt(row[3]) || 0);
        }
        
        return {
          pengganti: {
            idKaryawan: row[0],
            namaKaryawan: row[1],
            durasiTotal: parseInt(row[2]),
            hariDigunakan: parseInt(row[3]) || 0,
            sisaHari: sisaHari,
            keterangan: row[5] || ''
          }
        };
      }
    }
    
    return { pengganti: null };
  } catch (e) {
    Logger.log('Error in getPenggantiLiburKaryawan: ' + e.message);
    return { pengganti: null };
  }
}

// FUNGSI BARU: Ambil data Pengganti Libur untuk form edit (Admin)
function getPenggantiLiburById(idKaryawan) {
  try {
    Logger.log('getPenggantiLiburById called for ID: ' + idKaryawan);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetPengganti = null;
    let sheetKaryawan = null;
    
    try {
      sheetPengganti = ss.getSheetByName('Pengganti Libur');
      sheetKaryawan = ss.getSheetByName('Data Karyawan');
    } catch (e) {
      return { error: 'Sheet tidak ditemukan' };
    }
    
    if (!sheetPengganti || sheetPengganti.getLastRow() < 2) {
      return { error: 'Data Pengganti Libur tidak ditemukan' };
    }
    
    const dataPengganti = sheetPengganti.getRange(2, 1, sheetPengganti.getLastRow() - 1, 6).getValues();
    
    for (const row of dataPengganti) {
      if (String(row[0]).trim() === String(idKaryawan).trim()) {
        return {
          idKaryawan: row[0],
          namaKaryawan: row[1],
          durasiTotal: parseInt(row[2]),
          hariDigunakan: parseInt(row[3]) || 0,
          sisaHari: parseInt(row[4]) || 0,
          keterangan: row[5] || ''
        };
      }
    }
    
    return { error: 'Data tidak ditemukan untuk ID: ' + idKaryawan };
  } catch (e) {
    Logger.log('Error in getPenggantiLiburById: ' + e.message);
    return { error: 'Error: ' + e.message };
  }
}

// ===== FUNGSI DIAGNOSA =====
function diagnosaSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheets = ss.getSheets();
    
    Logger.log('=== DIAGNOSA SHEET ===');
    Logger.log('Total sheets: ' + sheets.length);
    
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const name = sheet.getName();
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      
      Logger.log('Sheet ' + i + ': "' + name + '" - Rows: ' + lastRow + ', Cols: ' + lastCol);
      
      // Jika ada header, tampilkan nama kolom
      if (lastRow > 0) {
        const headerRange = sheet.getRange(1, 1, 1, lastCol);
        const headerValues = headerRange.getValues()[0];
        Logger.log('  Header: ' + JSON.stringify(headerValues));
      }
    }
    
    // Test akses ke sheet Data Karyawan
    Logger.log('=== TEST AKSES DATA KARYAWAN ===');
    try {
      const sheetKaryawan = ss.getSheetByName('Data Karyawan');
      if (!sheetKaryawan) {
        Logger.log('ERROR: sheetKaryawan is NULL!');
      } else {
        Logger.log('Sheet Data Karyawan found!');
        Logger.log('Last row: ' + sheetKaryawan.getLastRow());
        
        if (sheetKaryawan.getLastRow() > 1) {
          const dataRange = sheetKaryawan.getRange(1, 1, Math.min(3, sheetKaryawan.getLastRow()), 2);
          const data = dataRange.getValues();
          Logger.log('Sample data (max 3 rows):');
          for (let i = 0; i < data.length; i++) {
            Logger.log('  Row ' + (i+1) + ': ' + JSON.stringify(data[i]));
          }
        }
      }
    } catch (e) {
      Logger.log('ERROR accessing Data Karyawan: ' + e.message);
    }
    
    // Test akses ke sheet Pengganti Libur
    Logger.log('=== TEST AKSES PENGGANTI LIBUR ===');
    try {
      const sheetPengganti = ss.getSheetByName('Pengganti Libur');
      if (!sheetPengganti) {
        Logger.log('Sheet Pengganti Libur NOT FOUND - will be created on save');
      } else {
        Logger.log('Sheet Pengganti Libur found!');
        Logger.log('Last row: ' + sheetPengganti.getLastRow());
      }
    } catch (e) {
      Logger.log('Sheet Pengganti Libur not found (normal): ' + e.message);
    }
    
    return 'Diagnosa selesai - lihat Execution Log';
  } catch (e) {
    Logger.log('ERROR in diagnosaSheet: ' + e.message);
    return 'Error: ' + e.message;
  }
}

// ===== FUNGSI UNTUK UPDATE JENIS KELAMIN KARYAWAN =====
function updateJenisKelaminKaryawan() {
  try {
    Logger.log('=== START: updateJenisKelaminKaryawan ===');
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Data Karyawan');
    const data = sheet.getRange("A2:I" + sheet.getLastRow()).getValues();
    
    Logger.log('Total karyawan:', data.length);
    
    // Update jenis kelamin berdasarkan nomor urut
    // No 1-15: Perempuan
    // No 16+: Laki-laki
    for (let i = 0; i < data.length; i++) {
      const nomorUrut = i + 1;
      const jenisKelamin = nomorUrut <= 15 ? 'Perempuan' : 'Laki-laki';
      
      Logger.log(`Row ${i + 2}: ID=${data[i][0]}, Nama=${data[i][1]}, Nomor=${nomorUrut}, Jenis Kelamin=${jenisKelamin}`);
      
      // Kolom I (index 8) adalah jenisKelamin
      sheet.getRange(i + 2, 9).setValue(jenisKelamin);
    }
    
    Logger.log('=== SELESAI: updateJenisKelaminKaryawan ===');
    return `✓ Jenis kelamin telah diupdate untuk ${data.length} karyawan. Karyawan 1-15: Perempuan, Karyawan 16+: Laki-laki`;
  } catch (e) {
    Logger.log('ERROR in updateJenisKelaminKaryawan:', e.message);
    return 'Error: ' + e.message;
  }
}

// ============ FUNGSI TANGGAL CUSTOM CUTI (NEW) ============

// Daftar kategori yang boleh custom tanggal
const KATEGORI_CUSTOM_TANGGAL = [
  'Istri Melahirkan',
  'Pernikahan Karyawan',
  'Keluarga Kandung Meninggal',
  'Pernikahan Anak Kandung',
  'Aqikah Anak Kandung',
  'Sunat Anak Kandung'
];

// FUNGSI: Cek apakah kategori boleh custom tanggal
function isCategoryAllowsCustomDate(kategoriNama) {
  return KATEGORI_CUSTOM_TANGGAL.includes(kategoriNama);
}

// FUNGSI HELPER: Ambil semua libur nasional
function getLiburNasional() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;
    
    try {
      sheet = ss.getSheetByName('Libur Nasional');
    } catch (e) {
      // Sheet belum ada
      return [];
    }
    
    if (sheet.getLastRow() < 2) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    return data.map(row => new Date(row[0]));
  } catch (e) {
    Logger.log('Error in getLiburNasional:', e.message);
    return [];
  }
}

// FUNGSI: Hitung hari kerja (skip Sabtu, Minggu, libur nasional)
function hitungHariKerja(tglMulai, tglSelesai, liburNasional = null) {
  try {
    const mulai = new Date(tglMulai);
    const selesai = new Date(tglSelesai);
    let hariKerja = 0;
    
    // Jika liburNasional null, ambil dari database
    if (liburNasional === null) {
      liburNasional = getLiburNasional();
    }
    
    // Convert libur nasional ke format yang bisa dibandingkan
    const liburSet = new Set();
    liburNasional.forEach(libur => {
      const d = new Date(libur);
      liburSet.add(d.toDateString());
    });
    
    // Loop dari tanggal mulai sampai selesai
    for (let d = new Date(mulai); d <= selesai; d.setDate(d.getDate() + 1)) {
      const dayOfWeek = d.getDay(); // 0=Minggu, 1=Senin, ..., 6=Sabtu
      const dateString = new Date(d).toDateString();
      
      // Skip Minggu (0) dan Sabtu (6)
      // Skip libur nasional
      if (dayOfWeek !== 0 && dayOfWeek !== 6 && !liburSet.has(dateString)) {
        hariKerja++;
      }
    }
    
    return hariKerja;
  } catch (e) {
    Logger.log('Error in hitungHariKerja:', e.message);
    return 0;
  }
}

// FUNGSI: Ambil semua data Tanggal Custom Cuti
function getAllTanggalCustomCuti() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;
    
    try {
      sheet = ss.getSheetByName('Tanggal Custom Cuti');
    } catch (e) {
      // Sheet belum ada, buat baru
      sheet = ss.insertSheet('Tanggal Custom Cuti');
      sheet.appendRow(['ID', 'ID Karyawan', 'Kategori', 'Tanggal', 'Durasi (Hari)', 'Deskripsi', 'Status Aktif', 'Dibuat Tanggal']);
    }
    
    if (sheet.getLastRow() < 2) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    return data.map((row, index) => ({
      id: row[0],
      idKaryawan: row[1],
      kategori: row[2],
      tanggal: new Date(row[3]).toLocaleDateString('id-ID'),
      durasiHari: row[4],
      deskripsi: row[5],
      statusAktif: row[6],
      dibuatTanggal: new Date(row[7]).toLocaleDateString('id-ID'),
      rowIndex: index + 2
    }));
  } catch (e) {
    Logger.log('Error in getAllTanggalCustomCuti:', e.message);
    return [];
  }
}

// FUNGSI: Ambil tanggal custom untuk karyawan tertentu
function getTanggalCustomCutiKaryawan(token) {
  try {
    // Get ID karyawan dari token
    const email = getEmailFromToken(token);
    if (!email) throw new Error("Sesi tidak valid.");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    const dataKaryawanRange = sheetKaryawan.getRange("A2:C" + sheetKaryawan.getLastRow()).getValues();
    let idKaryawan = null;
    
    for (const row of dataKaryawanRange) {
      if (row[2].toString().trim().toLowerCase() === email.toLowerCase()) {
        idKaryawan = row[0];
        break;
      }
    }
    
    if (!idKaryawan) throw new Error("Karyawan tidak ditemukan.");
    
    let sheet = null;
    
    try {
      sheet = ss.getSheetByName('Tanggal Custom Cuti');
    } catch (e) {
      return [];
    }
    
    if (sheet.getLastRow() < 2) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const tahunIni = new Date().getFullYear();
    
    return data.filter(row => {
      const isKaryawanMatch = row[1] === idKaryawan;
      const isTahunIni = new Date(row[3]).getFullYear() === tahunIni;
      const isAktif = row[6] === true || row[6] === 'TRUE' || row[6] === 'Aktif';
      return isKaryawanMatch && isTahunIni && isAktif;
    }).map(row => ({
      id: row[0],
      idKaryawan: row[1],
      kategori: row[2],
      tanggal: new Date(row[3]),
      tanggalFormatted: new Date(row[3]).toLocaleDateString('id-ID'),
      durasiHari: row[4],
      deskripsi: row[5]
    }));
  } catch (e) {
    Logger.log('Error in getTanggalCustomCutiKaryawan:', e.message);
    return [];
  }
}

// FUNGSI: Simpan Tanggal Custom Cuti
function simpanTanggalCustomCuti(dataCustom, token) {
  try {
    Logger.log('=== simpanTanggalCustomCuti START ===');
    Logger.log('Input:', JSON.stringify(dataCustom));
    
    const email = getEmailFromToken(token);
    if (!email) throw new Error("Sesi tidak valid atau telah berakhir.");
    
    // Validasi kategori
    if (!isCategoryAllowsCustomDate(dataCustom.kategori)) {
      throw new Error(`Kategori "${dataCustom.kategori}" tidak memungkinkan custom tanggal.`);
    }
    
    // Validasi tanggal
    const tanggalCustom = new Date(dataCustom.tanggal);
    if (isNaN(tanggalCustom.getTime())) {
      throw new Error("Format tanggal tidak valid.");
    }
    
    // Validasi durasi
    if (!dataCustom.durasiHari || dataCustom.durasiHari < 1) {
      throw new Error("Durasi hari minimal 1.");
    }
    
    // Cari ID karyawan berdasarkan email
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetKaryawan = ss.getSheetByName('Data Karyawan');
    const dataKaryawanRange = sheetKaryawan.getRange("A2:C" + sheetKaryawan.getLastRow()).getValues();
    let idKaryawan = null;
    
    for (const row of dataKaryawanRange) {
      if (row[2].toString().trim().toLowerCase() === email.toLowerCase()) {
        idKaryawan = row[0];
        break;
      }
    }
    
    if (!idKaryawan) throw new Error("Karyawan tidak ditemukan.");
    
    // Ambil atau buat sheet Tanggal Custom Cuti
    let sheet = null;
    try {
      sheet = ss.getSheetByName('Tanggal Custom Cuti');
    } catch (e) {
      sheet = ss.insertSheet('Tanggal Custom Cuti');
      sheet.appendRow(['ID', 'ID Karyawan', 'Kategori', 'Tanggal', 'Durasi (Hari)', 'Deskripsi', 'Status Aktif', 'Dibuat Tanggal']);
    }
    
    // Check apakah sudah ada record untuk kategori ini tahun ini
    if (sheet.getLastRow() > 1) {
      const existingData = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      const tahunIni = new Date().getFullYear();
      
      for (const row of existingData) {
        if (row[1] === idKaryawan && 
            row[2] === dataCustom.kategori && 
            new Date(row[3]).getFullYear() === tahunIni &&
            (row[6] === true || row[6] === 'TRUE' || row[6] === 'Aktif')) {
          throw new Error(`Anda sudah memiliki tanggal custom untuk "${dataCustom.kategori}" tahun ini.`);
        }
      }
    }
    
    // Simpan data
    const idCustom = 'CUSTOM-' + new Date().getTime();
    sheet.appendRow([
      idCustom,
      idKaryawan,
      dataCustom.kategori,
      tanggalCustom,
      dataCustom.durasiHari,
      dataCustom.deskripsi || '',
      true, // Status Aktif
      new Date()
    ]);
    
    Logger.log('=== simpanTanggalCustomCuti SUKSES ===');
    return { sukses: true, pesan: 'Tanggal custom cuti berhasil disimpan!', id: idCustom };
  } catch (e) {
    Logger.log('ERROR simpanTanggalCustomCuti:', e.message);
    return { sukses: false, pesan: 'Error: ' + e.message };
  }
}

// FUNGSI: Update Tanggal Custom Cuti
function updateTanggalCustomCuti(idCustom, dataCustom) {
  try {
    Logger.log('=== updateTanggalCustomCuti START ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Tanggal Custom Cuti');
    
    if (sheet.getLastRow() < 2) {
      throw new Error("Tidak ada data tanggal custom cuti.");
    }
    
    const ids = sheet.getRange("A2:A").getValues().flat();
    const rowIndex = ids.indexOf(idCustom);
    
    if (rowIndex === -1) {
      throw new Error("Tanggal custom cuti tidak ditemukan.");
    }
    
    // Update data
    sheet.getRange(rowIndex + 2, 3, 1, 5).setValues([[
      dataCustom.kategori,
      new Date(dataCustom.tanggal),
      dataCustom.durasiHari,
      dataCustom.deskripsi || '',
      dataCustom.statusAktif !== false
    ]]);
    
    Logger.log('=== updateTanggalCustomCuti SUKSES ===');
    return { sukses: true, pesan: 'Tanggal custom cuti berhasil diperbarui!' };
  } catch (e) {
    Logger.log('ERROR updateTanggalCustomCuti:', e.message);
    return { sukses: false, pesan: 'Error: ' + e.message };
  }
}

// FUNGSI: Hapus Tanggal Custom Cuti
function hapusTanggalCustomCuti(idCustom) {
  try {
    Logger.log('=== hapusTanggalCustomCuti START ===');
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Tanggal Custom Cuti');
    
    if (sheet.getLastRow() < 2) {
      throw new Error("Tidak ada data tanggal custom cuti.");
    }
    
    const ids = sheet.getRange("A2:A").getValues().flat();
    const rowIndex = ids.indexOf(idCustom);
    
    if (rowIndex === -1) {
      throw new Error("Tanggal custom cuti tidak ditemukan.");
    }
    
    sheet.deleteRow(rowIndex + 2);
    
    Logger.log('=== hapusTanggalCustomCuti SUKSES ===');
    return { sukses: true, pesan: 'Tanggal custom cuti berhasil dihapus!' };
  } catch (e) {
    Logger.log('ERROR hapusTanggalCustomCuti:', e.message);
    return { sukses: false, pesan: 'Error: ' + e.message };
  }
}

// FUNGSI: Validasi pengajuan cuti dengan custom tanggal
function validatePengajuanDenganCustomDate(idKaryawan, kategoriNama, tanggalRequest, durasiRequest) {
  try {
    // Jika kategori tidak menggunakan custom tanggal, skip validasi ini
    if (!isCategoryAllowsCustomDate(kategoriNama)) {
      return { valid: true };
    }
    
    // Ambil data custom tanggal karyawan berdasarkan idKaryawan
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = null;
    
    try {
      sheet = ss.getSheetByName('Tanggal Custom Cuti');
    } catch (e) {
      return { valid: true }; // Sheet belum ada, treat sebagai valid
    }
    
    if (sheet.getLastRow() < 2) {
      return { valid: true }; // Tidak ada data, treat sebagai valid
    }
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const tahunIni = new Date().getFullYear();
    
    // Cari matching custom date untuk kategori ini tahun ini
    const matchingCustom = data.find(row => {
      const isKaryawanMatch = row[1] === idKaryawan;
      const isKategoriMatch = row[2] === kategoriNama;
      const isTahunIni = new Date(row[3]).getFullYear() === tahunIni;
      const isAktif = row[6] === true || row[6] === 'TRUE' || row[6] === 'Aktif';
      return isKaryawanMatch && isKategoriMatch && isTahunIni && isAktif;
    });
    
    if (!matchingCustom) {
      // Tidak ada custom date, gunakan tanggal yang diinput
      return { 
        valid: true, 
        usesCustomDate: false,
        tanggal: tanggalRequest,
        durasi: durasiRequest
      };
    }
    
    // Ada custom date, gunakan itu
    return {
      valid: true,
      usesCustomDate: true,
      tanggal: matchingCustom[3],
      tanggalFormatted: new Date(matchingCustom[3]).toLocaleDateString('id-ID'),
      durasi: matchingCustom[4],
      deskripsi: matchingCustom[5]
    };
  } catch (e) {
    Logger.log('Error in validatePengajuanDenganCustomDate:', e.message);
    return { valid: false, error: e.message };
  }
}