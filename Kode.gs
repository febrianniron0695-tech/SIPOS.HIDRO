// ==========================================================
// ROUTER API UNTUK GITHUB PAGES (GET & POST)
// ==========================================================
function doGet(e) {
  try {
    var action = e.parameter.action;
    var result = { success: false, message: "Action tidak dikenal" };

    if (action === 'getUserList') {
      result = getUserList(e.parameter.role, e.parameter.island);
    } else if (action === 'verifyLogin') {
      result = verifyLogin(e.parameter.role, e.parameter.username, e.parameter.password, e.parameter.island);
    } else if (action === 'getAllSheetNames') {
      result = getAllSheetNames(e.parameter.islandName);
    } else if (action === 'getDataBySheetName') {
      result = getDataBySheetName(e.parameter.sheetName, e.parameter.islandName);
    } else if (action === 'getPetugasMonthlyRainfall') {
      result = getPetugasMonthlyRainfall(e.parameter.stationName, e.parameter.islandName);
    } else if (action === 'getAdminSummaryStats') {
      result = getAdminSummaryStats(e.parameter.islandName);
    } else if (action === 'getAdminChartDataByDistrict') {
      result = getAdminChartDataByDistrict(e.parameter.islandName);
    } else if (action === 'getAdminReports') {
      result = getAdminReports(e.parameter.islandName);
    } else if (action === 'getVerifikatorReportList') {
      result = getVerifikatorReportList(e.parameter.islandName, e.parameter.verifikatorKabupaten);
    } else {
      result = { success: true, message: "SI POS HIDRO API Backend Berjalan Normal." };
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    var action = e.parameter.action;
    var postData = JSON.parse(e.postData.contents);
    var result = { success: false, message: "Action POST tidak dikenal" };

    if (action === 'submitReport') {
      result = submitReport(postData);
    } else if (action === 'updateVerifikatorDecision') {
      result = updateVerifikatorDecision(
        postData.rowIndex, 
        postData.statusVerifikasi, 
        postData.catatan, 
        postData.namaVerifikator, 
        postData.verifikatorKab, 
        postData.islandName
      );
    } else if (action === 'deleteMultipleReports') {
      result = deleteMultipleReports(postData.rowIndices, postData.islandName);
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ==========================================================
// KONFIGURASI MULTI-DATABASE SPREADSHEET TIGA PULAU
// ==========================================================
const DB_MAPPING = {
  "pulau timor": "Si_Pos_Hidro_TIMOR_DB",
  "pulau sumba": "Si_Pos_Hidro_SUMBA_DB",
  "pulau flores": "Si_Pos_Hidro_FLORES_DB"
};

function getSpreadsheetByIsland(islandName) {
  try {
    var key = (islandName || "Pulau Timor").toString().toLowerCase().trim();
    var ssName = DB_MAPPING[key] || DB_MAPPING["pulau timor"];
    var files = DriveApp.getFilesByName(ssName);
    if (files.hasNext()) return SpreadsheetApp.open(files.next());
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

function sanitizeName(name) {
  return name ? name.replace(/[\/\?*:[\]]/g, '_').trim() : 'Tanpa_Nama';
}

function normalizeKabName(rawKab) {
  if (!rawKab) return 'Kab. Lainnya';
  var kab = rawKab.toString().toUpperCase().trim();
  if (kab.includes('KUPANG') && !kab.includes('KOTA')) return 'Kab. Kupang';
  if (kab.includes('KOTA KUPANG')) return 'Kota Kupang';
  if (kab.includes('TTS') || kab.includes('TIMOR TENGAH SELATAN')) return 'Kab. TTS';
  if (kab.includes('TTU') || kab.includes('TIMOR TENGAH UTARA')) return 'Kab. TTU';
  if (kab.includes('BELU')) return 'Kab. Belu';
  if (kab.includes('MALAKA')) return 'Kab. Malaka';
  if (kab.includes('ROTE') || kab.includes('NDAO')) return 'Kab. Rote Ndao';
  if (kab.includes('SABU') || kab.includes('RAIJUAN')) return 'Kab. Sabu Raijua';
  if (kab.includes('SUMBA TIMUR')) return 'Kab. Sumba Timur';
  if (kab.includes('SUMBA BARAT DAYA')) return 'Kab. Sumba Barat Daya';
  if (kab.includes('SUMBA BARAT')) return 'Kab. Sumba Barat';
  if (kab.includes('SUMBA TENGAH')) return 'Kab. Sumba Tengah';
  if (kab.includes('FLORES TIMUR')) return 'Kab. Flores Timur';
  if (kab.includes('SIKKA')) return 'Kab. Sikka';
  if (kab.includes('ENDE')) return 'Kab. Ende';
  if (kab.includes('NGADA')) return 'Kab. Ngada';
  if (kab.includes('NAGEKEO')) return 'Kab. Nagekeo';
  if (kab.includes('MANGGARAI BARAT')) return 'Kab. Manggarai Barat';
  if (kab.includes('MANGGARAI TIMUR')) return 'Kab. Manggarai Timur';
  if (kab.includes('MANGGARAI')) return 'Kab. Manggarai';
  if (kab.includes('LEMBATA')) return 'Kab. Lembata';
  if (kab.includes('ALOR')) return 'Kab. Alor';
  return rawKab;
}

function getUserList(role, islandName) {
  try {
    var ss = getSpreadsheetByIsland(islandName);
    var sheet = ss.getSheetByName('USERNAME LOGIN');
    if (!sheet) return [];
    var data = sheet.getDataRange().getValues();
    var list = [];
    
    for (var i = 1; i < data.length; i++) {
      if (role === 'admin') {
        var name = data[i][5] ? data[i][5].toString().trim() : '';
        if (name !== '' && list.indexOf(name) === -1) list.push(name);
      } else if (role === 'verifikator') {
        var verifName = data[i][9] ? data[i][9].toString().trim() : '';
        if (verifName !== '' && list.indexOf(verifName) === -1) list.push(verifName);
      } else {
        var tipeStasiun = data[i][0] ? data[i][0].toString().trim() : '';
        var namaStasiun = data[i][1] ? data[i][1].toString().trim() : '';
        var displayPetugas = (tipeStasiun ? tipeStasiun + ' - ' : '') + namaStasiun;
        if (displayPetugas.trim() !== '' && list.indexOf(displayPetugas) === -1) list.push(displayPetugas);
      }
    }
    return list;
  } catch (error) {
    return [];
  }
}

function verifyLogin(role, username, password, islandName) {
  try {
    var targetIsland = islandName || "Pulau Timor";
    var ss = getSpreadsheetByIsland(targetIsland);
    var sheet = ss.getSheetByName('USERNAME LOGIN');
    if (!sheet) return { success: false, message: 'Sheet "USERNAME LOGIN" tidak ditemukan pada database ' + targetIsland + '!' };
    
    var data = sheet.getDataRange().getValues();
    var inputUser = username.trim().toLowerCase();
    var inputPass = password.trim();

    if (role === 'admin') {
      for (var i = 1; i < data.length; i++) {
        var dbNamaAdmin = data[i][5] ? data[i][5].toString().trim() : '';
        var dbJabatan   = data[i][6] ? data[i][6].toString().trim() : '';
        var dbPassAdmin = data[i][7] ? data[i][7].toString().trim() : '';
        if (dbNamaAdmin !== '' && dbNamaAdmin.toLowerCase() === inputUser && dbPassAdmin === inputPass) {
          recordAdminLoginHistory(ss, dbNamaAdmin, dbJabatan);
          return { success: true, role: 'admin', name: dbNamaAdmin + (dbJabatan ? ' (' + dbJabatan + ')' : ''), island: targetIsland };
        }
      }
      return { success: false, message: 'Nama atau Password Admin di ' + targetIsland + ' salah!' };

    } else if (role === 'verifikator') {
      var lastRowJ = getLastRowInColumn(sheet, 10);
      for (var k = 1; k < lastRowJ; k++) {
        var dbNamaVerif = data[k][9] ? data[k][9].toString().trim() : '';
        var dbKabVerif  = data[k][10] ? data[k][10].toString().trim() : '';
        var dbPassVerif = data[k][11] ? data[k][11].toString().trim() : '';
        if (dbNamaVerif !== '' && dbNamaVerif.toLowerCase() === inputUser && dbPassVerif === inputPass) {
          recordVerifikatorLoginHistory(ss, dbNamaVerif, dbKabVerif);
          return {
            success: true,
            role: 'verifikator',
            name: dbNamaVerif,
            kabupaten: normalizeKabName(dbKabVerif),
            island: targetIsland
          };
        }
      }
      return { success: false, message: 'Nama Verifikator atau Password salah!' };

    } else {
      for (var j = 1; j < data.length; j++) {
        var dbTipeStasiun = data[j][0] ? data[j][0].toString().trim() : '';
        var dbNamaStasiun = data[j][1] ? data[j][1].toString().trim() : '';
        var dbKabupaten   = data[j][2] ? data[j][2].toString().trim() : '';
        var dbPassPetugas = data[j][3] ? data[j][3].toString().trim() : '';
        var fullStasiunName = (dbTipeStasiun ? dbTipeStasiun + ' - ' : '') + dbNamaStasiun;
        if (fullStasiunName.trim() !== '' && fullStasiunName.toLowerCase() === inputUser && dbPassPetugas === inputPass) {
          return { success: true, role: 'petugas', name: fullStasiunName + (dbKabupaten ? ' (' + dbKabupaten + ')' : ''), island: targetIsland };
        }
      }
      return { success: false, message: 'Nama Stasiun atau Password Petugas di ' + targetIsland + ' salah!' };
    }
  } catch (error) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + error.toString() };
  }
}

function recordAdminLoginHistory(ss, namaAdmin, jabatanAdmin) {
  try {
    var historySheet = ss.getSheetByName('Riwayat Login Admin') || ss.insertSheet('Riwayat Login Admin');
    if (historySheet.getLastRow() === 0) {
      historySheet.appendRow(['Timestamp (WITA)', 'Nama Admin', 'Jabatan', 'Status Login', 'Keterangan Sistem']);
      historySheet.getRange("A1:E1").setBackground("#312e81").setFontColor("#ffffff").setFontWeight("bold");
    }
    historySheet.appendRow([Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss"), namaAdmin, jabatanAdmin || 'Administrator', 'BERHASIL MASUK', 'Login sukses']);
  } catch (err) {}
}

function recordVerifikatorLoginHistory(ss, namaVerif, kabVerif) {
  try {
    var historySheet = ss.getSheetByName('Log Login Verifikator') || ss.insertSheet('Log Login Verifikator');
    if (historySheet.getLastRow() === 0) {
      historySheet.appendRow(['Timestamp (WITA)', 'Nama Verifikator', 'Kabupaten', 'Status Login', 'Keterangan Sistem']);
      historySheet.getRange("A1:E1").setBackground("#047857").setFontColor("#ffffff").setFontWeight("bold");
    }
    historySheet.appendRow([Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss"), namaVerif, kabVerif || 'Kabupaten', 'BERHASIL MASUK', 'Login sukses']);
  } catch (err) {}
}

function getLastRowInColumn(sheet, columnNumber) {
  var values = sheet.getRange(1, columnNumber, sheet.getMaxRows(), 1).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    if (values[i][0] !== "") return i + 1;
  }
  return 0;
}

function getAllSheetNames(islandName) {
  try {
    var sheets = getSpreadsheetByIsland(islandName).getSheets();
    var sheetNames = [];
    for (var i = 0; i < sheets.length; i++) sheetNames.push(sheets[i].getName());
    return sheetNames;
  } catch (error) { return []; }
}

function getDataBySheetName(sheetName, islandName) {
  try {
    var ss = getSpreadsheetByIsland(islandName);
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { headers: [], rows: [], error: "Sheet " + sheetName + " tidak ditemukan." };
    var data = sheet.getDataRange().getValues();
    if (data.length === 0) return { headers: [], rows: [] };
    return { headers: data[0], rows: data.slice(1) };
  } catch (error) { return { headers: [], rows: [], error: error.toString() }; }
}

function getStationSheetUrl(stationName, islandName) {
  try {
    var ss = getSpreadsheetByIsland(islandName);
    var rawStationName = stationName || "";
    var matchName = rawStationName.match(/^(.*?)\s*\(/);
    var stationClean = matchName ? matchName[1].trim() : rawStationName.trim();
    var sheetName = 'Pos_' + sanitizeName(stationClean) + '_Rekap';
    var sheet = ss.getSheetByName(sheetName);
    if (sheet) return ss.getUrl() + '#gid=' + sheet.getSheetId();
    var masterSheet = ss.getSheetByName('RekapLaporan');
    if (masterSheet) return ss.getUrl() + '#gid=' + masterSheet.getSheetId();
    return ss.getUrl();
  } catch (e) {
    return "";
  }
}

function getPetugasMonthlyRainfall(stationName, islandName) {
  try {
    var ss = getSpreadsheetByIsland(islandName);
    var rawStationName = stationName || "";
    var matchName = rawStationName.match(/^(.*?)\s*\(/);
    var stationClean = matchName ? matchName[1].trim() : rawStationName.trim();
    var sheet = ss.getSheetByName('Pos_' + sanitizeName(stationClean) + '_Harian');
    var monthsIndo = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
    var now = new Date();
    var currentYear = now.getFullYear().toString();
    var currentMonthName = monthsIndo[now.getMonth()];
    var lastDay = new Date(now.getFullYear(), now.getMonth() + 1, 0).getDate();
    var dailyMap = {};
    for (var d = 1; d <= lastDay; d++) dailyMap["Tgl " + d] = 0;
    
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = 1; i < data.length; i++) {
        var rowBulan = data[i][4] ? data[i][4].toString().trim() : '';
        var rowTahun = data[i][5] ? data[i][5].toString().trim() : '';
        var rowTanggal = data[i][6] ? data[i][6].toString().trim() : '';
        var rowHujan = parseFloat(data[i][7]) || 0;
        if (rowBulan.toLowerCase() === currentMonthName.toLowerCase() && rowTahun === currentYear) {
          var key = rowTanggal.toLowerCase().includes('tgl') ? rowTanggal : ('Tgl ' + rowTanggal);
          if (dailyMap.hasOwnProperty(key)) dailyMap[key] += rowHujan;
        }
      }
    }
    var labels = [], values = [];
    for (var day = 1; day <= lastDay; day++) {
      labels.push('Tgl ' + day);
      values.push(dailyMap['Tgl ' + day] || 0);
    }
    return { success: true, labels: labels, data: values };
  } catch (error) { return { success: false, labels: [], data: [] }; }
}

function getAdminSummaryStats(islandName) {
  try {
    var ss = getSpreadsheetByIsland(islandName);
    var sheetUser = ss.getSheetByName('USERNAME LOGIN');
    var totalStasiun = 0;
    
    if (sheetUser) {
      var userData = sheetUser.getDataRange().getValues();
      for (var i = 1; i < userData.length; i++) {
        if (userData[i][1] && userData[i][1].toString().trim() !== '') totalStasiun++;
      }
    }
    
    var sheetLaporan = ss.getSheetByName('RekapLaporan');
    var latestStationStatus = {};

    if (sheetLaporan) {
      var data = sheetLaporan.getDataRange().getValues();
      for (var j = 1; j < data.length; j++) {
        var rawStationName = data[j][2] ? data[j][2].toString().trim() : '';
        var gelasKondisi = data[j][5] ? data[j][5].toString().trim() : '';
        var mistarKondisi = data[j][6] ? data[j][6].toString().trim() : '';
        var kranKondisi  = data[j][7] ? data[j][7].toString().trim() : '';
        var kondisiAset  = data[j][8] ? data[j][8].toString().trim() : '';
        
        var cleanKey = rawStationName.replace(/\s*\([^)]*\)/g, '').trim().toLowerCase();
        if (cleanKey !== '') {
          latestStationStatus[cleanKey] = { gelas: gelasKondisi, mistar: mistarKondisi, kran: kranKondisi, pos: kondisiAset };
        }
      }
    }

    var alatBaik = 0, alatBuruk = 0, posBaik = 0, posBuruk = 0;

    function isKondisiBaik(condStr) {
      if (!condStr) return false;
      var str = condStr.toString().toLowerCase();
      return str.includes('baik') && !str.includes('rusak') && !str.includes('hilang') && !str.includes('rata');
    }

    for (var st in latestStationStatus) {
      var status = latestStationStatus[st];
      if (isKondisiBaik(status.gelas) && isKondisiBaik(status.mistar) && isKondisiBaik(status.kran)) alatBaik++;
      else alatBuruk++;

      if (isKondisiBaik(status.pos)) posBaik++;
      else posBuruk++;
    }

    return { success: true, totalStasiun: totalStasiun, alatBaik: alatBaik, alatBuruk: alatBuruk, posBaik: posBaik, posBuruk: posBuruk };
  } catch (error) { 
    return { success: false, totalStasiun: 0, alatBaik: 0, alatBuruk: 0, posBaik: 0, posBuruk: 0 }; 
  }
}

function getAdminChartDataByDistrict(islandName) {
  try {
    var ss = getSpreadsheetByIsland(islandName);
    var sheetUser = ss.getSheetByName('USERNAME LOGIN');
    var posToKabMap = {};
    
    if (sheetUser) {
      var userData = sheetUser.getDataRange().getValues();
      for (var i = 1; i < userData.length; i++) {
        var tipe = userData[i][0] ? userData[i][0].toString().trim() : '';
        var nama = userData[i][1] ? userData[i][1].toString().trim() : '';
        var kab  = userData[i][2] ? userData[i][2].toString().trim() : '';
        var fullName = (tipe ? tipe + ' - ' : '') + nama;
        var normKab = normalizeKabName(kab);
        if (fullName.trim() !== '') {
          posToKabMap[fullName.toLowerCase()] = normKab;
          posToKabMap[nama.toLowerCase()] = normKab;
        }
      }
    }

    var sheetLaporan = ss.getSheetByName('RekapLaporan');
    var finalKabData = {};
    var stationLatestStatus = {};
    
    if (sheetLaporan) {
      var data = sheetLaporan.getDataRange().getValues();
      var latestTimestamp = data.length > 1 ? (data[data.length - 1][0] ? data[data.length - 1][0].toString() : "") : "";
      var latestYearMonth = latestTimestamp.substring(0, 7); 
      var rainfallByKabAndStation = {};
      
      for (var j = 1; j < data.length; j++) {
        var rawStationName = data[j][2] ? data[j][2].toString().trim() : 'Pos Unassigned';
        var totalRain = parseFloat(data[j][4]) || 0;
        var kGelas = data[j][5] ? data[j][5].toString().trim() : '';
        var kMistar = data[j][6] ? data[j][6].toString().trim() : '';
        var kKran = data[j][7] ? data[j][7].toString().trim() : '';
        var rowTimestamp = data[j][0] ? data[j][0].toString() : '';
        var rowYearMonth = rowTimestamp.substring(0, 7);
        
        var detectedKab = '';
        var matchInParen = rawStationName.match(/\(([^)]+)\)/);
        if (matchInParen && matchInParen[1]) detectedKab = normalizeKabName(matchInParen[1]);
        
        var cleanStationName = rawStationName.replace(/\s*\([^)]*\)/g, '').trim();
        var cleanKey = cleanStationName.toLowerCase();
        var kabName = normalizeKabName(detectedKab || posToKabMap[cleanKey] || 'Kab. Utama');

        stationLatestStatus[cleanKey] = { kabupaten: kabName, gelas: kGelas, mistar: kMistar, kran: kKran };
        
        if (rowYearMonth === latestYearMonth) {
          if (!rainfallByKabAndStation[kabName]) rainfallByKabAndStation[kabName] = {};
          if (!rainfallByKabAndStation[kabName][cleanStationName]) rainfallByKabAndStation[kabName][cleanStationName] = 0;
          rainfallByKabAndStation[kabName][cleanStationName] += totalRain;
        }
      }

      for (var kName in rainfallByKabAndStation) {
        if (!finalKabData[kName]) {
          finalKabData[kName] = { gelasCounts: [0, 0, 0, 0], mistarCounts: [0, 0, 0, 0], kranCounts: [0, 0, 0, 0], rainfallMap: {} };
        }
        finalKabData[kName].rainfallMap = rainfallByKabAndStation[kName];
      }
    }

    function evalConditionIndex(condStr) {
      if (!condStr) return 0;
      var str = condStr.toString();
      if (str.includes('Baik')) return 0;
      if (str.includes('Ringan') || str.includes('RR')) return 1;
      if (str.includes('Berat') || str.includes('RB')) return 2;
      if (str.includes('Hilang') || str.includes('H/RT') || str.includes('Rata')) return 3;
      return 0;
    }

    for (var stKey in stationLatestStatus) {
      var kTarget = stationLatestStatus[stKey].kabupaten;
      var stInfo = stationLatestStatus[stKey];
      if (!finalKabData[kTarget]) {
        finalKabData[kTarget] = { gelasCounts: [0, 0, 0, 0], mistarCounts: [0, 0, 0, 0], kranCounts: [0, 0, 0, 0], rainfallMap: {} };
      }
      finalKabData[kTarget].gelasCounts[evalConditionIndex(stInfo.gelas)]++;
      finalKabData[kTarget].mistarCounts[evalConditionIndex(stInfo.mistar)]++;
      finalKabData[kTarget].kranCounts[evalConditionIndex(stInfo.kran)]++;
    }
    
    return { success: true, data: finalKabData };
  } catch (error) { return { success: false, error: error.toString() }; }
}

function submitReport(formData) {
  try {
    var targetIsland = formData.island || 'Pulau Timor';
    var ss = getSpreadsheetByIsland(targetIsland);
    var waktuWitaStr = Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss");
    var rawStationName = formData.stationName || "Stasiun";
    var matchName = rawStationName.match(/^(.*?)\s*\(/);
    var stationClean = matchName ? matchName[1].trim() : rawStationName.trim();
    var cleanStationName = sanitizeName(stationClean);
    
    var standardHeaders = [
      'Timestamp (WITA)', 'ID Pos', 'Nama Pos', 'Periode', 
      'Total Hujan (mm)', 'Gelas Penakar', 'Mistar Ukur', 'Kran Penakar',
      'Kondisi Aset', 'Kondisi Formulir', 'Latitude', 'Longitude', 'Tautan Google Drive', 
      'Foto 1', 'Foto 2', 'Foto 3', 'Foto 4', 'Foto 5', 'Foto 6', 'Foto 7', 'Foto 8', 'Foto 9', 'Foto 10',
      'Status Verifikasi', 'Catatan Verifikator', 'Oleh Verifikator', 'Riwayat Log Verifikasi (JSON)'
    ];

    var sheetHarian = ss.getSheetByName('Pos_' + cleanStationName + '_Harian') || ss.insertSheet('Pos_' + cleanStationName + '_Harian');
    if (sheetHarian.getLastRow() === 0) {
      sheetHarian.appendRow(['Timestamp (WITA)', 'ID Pos', 'Nama Pos', 'Periode', 'Bulan', 'Tahun', 'Tanggal', 'Curah Hujan (mm)']);
      sheetHarian.getRange("A1:H1").setBackground("#1e3a8a").setFontColor("#ffffff").setFontWeight("bold");
    }
    
    var sheetRekap = ss.getSheetByName('Pos_' + cleanStationName + '_Rekap') || ss.insertSheet('Pos_' + cleanStationName + '_Rekap');
    if (sheetRekap.getLastRow() === 0) {
      sheetRekap.appendRow(standardHeaders);
      sheetRekap.getRange("A1:AA1").setBackground("#4f46e5").setFontColor("#ffffff").setFontWeight("bold");
    }
    
    var masterLaporan = ss.getSheetByName('RekapLaporan') || ss.insertSheet('RekapLaporan');
    if (masterLaporan.getLastRow() === 0) {
      masterLaporan.appendRow(standardHeaders);
      masterLaporan.getRange("A1:AA1").setBackground("#4f46e5").setFontColor("#ffffff").setFontWeight("bold");
    }
    
    var rootFolder = DriveApp.getFoldersByName('Si_Pos_Hidro_Storage').hasNext() ? DriveApp.getFoldersByName('Si_Pos_Hidro_Storage').next() : DriveApp.createFolder('Si_Pos_Hidro_Storage');
    var stationFolder = rootFolder.getFoldersByName(cleanStationName).hasNext() ? rootFolder.getFoldersByName(cleanStationName).next() : rootFolder.createFolder(cleanStationName);
    var timestampStr = Utilities.formatDate(new Date(), "GMT+8", "yyyyMMdd_HHmmss");
    var reportFolder = stationFolder.createFolder(sanitizeName(formData.period) + '_' + timestampStr);
    
    reportFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    var folderUrl = reportFolder.getUrl();
    
    var photoFormulas = [];
    var photoArray = formData.photos || [];
    for (var p = 0; p < 10; p++) {
      if (p < photoArray.length && photoArray[p].data) {
        try {
          var photo = photoArray[p];
          var splitBase64 = photo.data.split(',');
          var blob = Utilities.newBlob(Utilities.base64Decode(splitBase64[1]), splitBase64[0].match(/:(.*?);/)[1], photo.name);
          var savedFile = reportFolder.createFile(blob);
          savedFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
          photoFormulas.push('=IMAGE("https://lh3.googleusercontent.com/d/' + savedFile.getId() + '")');
        } catch (e) { photoFormulas.push(''); }
      } else photoFormulas.push('');
    }
    
    var totalRainfall = 0;
    var monthYearParts = (formData.monthYear || "").split(" ");
    var bulanVal = monthYearParts[0] || "", tahunVal = monthYearParts[1] || "2026";

    formData.dailyData.forEach(function(item) {
      var val = parseFloat(item.rainfall) || 0;
      totalRainfall += val;
      sheetHarian.appendRow([waktuWitaStr, formData.stationId, formData.stationName, formData.period, bulanVal, tahunVal, item.date, val]);
    });

    var rowData = [
      waktuWitaStr, formData.stationId, formData.stationName, formData.period, totalRainfall,
      formData.kondisiGelasPenakar + (formData.ketGelas ? ' (' + formData.ketGelas + ')' : ''),
      formData.kondisiMistar + (formData.ketMistar ? ' (' + formData.ketMistar + ')' : ''),
      formData.kondisiKran + (formData.ketKran ? ' (' + formData.ketKran + ')' : ''),
      formData.kondisiAset + (formData.ketPos ? ' (' + formData.ketPos + ')' : ''),
      formData.kondisiFormulir, formData.latitude || '-', formData.longitude || '-', folderUrl,
      photoFormulas[0], photoFormulas[1], photoFormulas[2], photoFormulas[3], photoFormulas[4],
      photoFormulas[5], photoFormulas[6], photoFormulas[7], photoFormulas[8], photoFormulas[9],
      'PENDING', '', '', '{}'
    ];
    
    sheetRekap.appendRow(rowData);
    sheetRekap.setRowHeight(sheetRekap.getLastRow(), 110); 

    masterLaporan.appendRow(rowData);
    masterLaporan.setRowHeight(masterLaporan.getLastRow(), 110);
    
    return { success: true, message: 'Laporan berhasil terkirim ke database ' + targetIsland + '!' };
  } catch (error) { return { success: false, message: 'Gagal mengirim laporan: ' + error.toString() }; }
}

function getAdminReports(islandName) {
  var ss = getSpreadsheetByIsland(islandName);
  var sheetLaporan = ss.getSheetByName('RekapLaporan');
  var sheetUser = ss.getSheetByName('USERNAME LOGIN');
  
  var verifikatorKabMap = {};
  if (sheetUser) {
    var lastRowJ = getLastRowInColumn(sheetUser, 10);
    var vData = sheetUser.getRange(2, 10, Math.max(lastRowJ - 1, 1), 3).getValues();
    for (var v = 0; v < vData.length; v++) {
      var vNama = vData[v][0] ? vData[v][0].toString().trim() : '';
      var vKab  = vData[v][1] ? normalizeKabName(vData[v][1]) : '';
      if (vNama && vKab) {
        if (!verifikatorKabMap[vKab]) verifikatorKabMap[vKab] = [];
        verifikatorKabMap[vKab].push(vNama);
      }
    }
  }

  if (!sheetLaporan) return { reports: [], verifikatorMap: verifikatorKabMap };
  
  var data = sheetLaporan.getDataRange().getValues();
  var rows = [];

  for (var i = 1; i < data.length; i++) {
    var rawStationName = data[i][2] ? data[i][2].toString().trim() : '';
    var gelas = data[i][5] || 'Baik (B)';
    var mistar = data[i][6] || 'Baik (B)';
    var kran = data[i][7] || 'Baik (B)';
    var combinedAlat = "Gelas: " + gelas + " | Mistar: " + mistar + " | Kran: " + kran;

    var detectedKab = 'Kab. Utama';
    var matchInParen = rawStationName.match(/\(([^)]+)\)/);
    if (matchInParen && matchInParen[1]) detectedKab = normalizeKabName(matchInParen[1]);

    var statusVerif = (data[i].length > 23 && data[i][23]) ? data[i][23].toString().trim() : 'PENDING';
    var catatanVerif = (data[i].length > 24 && data[i][24]) ? data[i][24].toString().trim() : '';
    var olehVerif = (data[i].length > 25 && data[i][25]) ? data[i][25].toString().trim() : '';
    var logJsonStr = (data[i].length > 26 && data[i][26]) ? data[i][26].toString().trim() : '{}';

    var currentLogs = {};
    try { currentLogs = JSON.parse(logJsonStr || '{}'); } catch(e) {}

    var allKabVerifs = verifikatorKabMap[detectedKab] || [];
    var pendingVerifs = [];
    allKabVerifs.forEach(function(vName) {
      if (!currentLogs[vName]) pendingVerifs.push(vName);
    });

    rows.push({
      rowIndex: i + 1,
      timestamp: data[i][0] ? data[i][0].toString() : '',
      stationId: data[i][1],
      stationName: rawStationName,
      kabupaten: detectedKab,
      period: data[i][3],
      totalRainfall: data[i][4],
      kondisiAlat: combinedAlat,
      folderUrl: data[i][12] || '',
      statusVerifikasi: statusVerif,
      catatanVerifikator: catatanVerif,
      olehVerifikator: olehVerif,
      pendingVerifikators: pendingVerifs
    });
  }
  return { reports: rows.reverse(), verifikatorMap: verifikatorKabMap };
}

function getVerifikatorReportList(islandName, verifikatorKabupaten) {
  var ss = getSpreadsheetByIsland(islandName);
  var sheetLaporan = ss.getSheetByName('RekapLaporan');
  if (!sheetLaporan) return [];
  
  var data = sheetLaporan.getDataRange().getValues();
  var rows = [];
  var targetKab = normalizeKabName(verifikatorKabupaten);

  for (var i = 1; i < data.length; i++) {
    var rawStationName = data[i][2] ? data[i][2].toString().trim() : '';
    var gelas = data[i][5] || 'Baik (B)';
    var mistar = data[i][6] || 'Baik (B)';
    var kran = data[i][7] || 'Baik (B)';
    var combinedAlat = "Gelas: " + gelas + " | Mistar: " + mistar + " | Kran: " + kran;
    var kondisiPosAset = data[i][8] || 'Baik (B)';

    var detectedKab = 'Kab. Utama';
    var matchInParen = rawStationName.match(/\(([^)]+)\)/);
    if (matchInParen && matchInParen[1]) detectedKab = normalizeKabName(matchInParen[1]);

    var logJsonStr = (data[i].length > 26 && data[i][26]) ? data[i][26].toString().trim() : '{}';

    if (!verifikatorKabupaten || detectedKab.toLowerCase() === targetKab.toLowerCase()) {
      rows.push({
        rowIndex: i + 1,
        timestamp: data[i][0] ? data[i][0].toString() : '',
        stationId: data[i][1],
        stationName: rawStationName,
        kabupaten: detectedKab,
        period: data[i][3],
        totalRainfall: data[i][4],
        kondisiAlat: combinedAlat,
        kondisiPos: kondisiPosAset,
        folderUrl: data[i][12] || '',
        sheetRekapUrl: getStationSheetUrl(rawStationName, islandName),
        statusVerifikasi: data[i][23] || 'PENDING',
        catatanVerifikator: data[i][24] || '',
        olehVerifikator: data[i][25] || '',
        logVerifikasi: logJsonStr
      });
    }
  }
  return rows.reverse();
}

function updateVerifikatorDecision(rowIndex, statusVerifikasi, catatan, namaVerifikator, verifikatorKab, islandName) {
  try {
    var ss = getSpreadsheetByIsland(islandName);
    var sheetLaporan = ss.getSheetByName('RekapLaporan');
    var sheetUser = ss.getSheetByName('USERNAME LOGIN');
    if (!sheetLaporan) return { success: false, message: 'Sheet RekapLaporan tidak ditemukan.' };
    
    var rawStationName = sheetLaporan.getRange(rowIndex, 3).getValue().toString();
    var detectedKab = 'Kab. Utama';
    var matchInParen = rawStationName.match(/\(([^)]+)\)/);
    if (matchInParen && matchInParen[1]) detectedKab = normalizeKabName(matchInParen[1]);

    if (verifikatorKab && detectedKab.toLowerCase() !== normalizeKabName(verifikatorKab).toLowerCase()) {
      return { success: false, message: '❌ Akses Ditolak: Anda tidak memiliki hak memverifikasi data dari kabupaten ' + detectedKab + '!' };
    }

    var allKabVerifs = [];
    if (sheetUser) {
      var lastRowJ = getLastRowInColumn(sheetUser, 10);
      var vData = sheetUser.getRange(2, 10, Math.max(lastRowJ - 1, 1), 3).getValues();
      for (var v = 0; v < vData.length; v++) {
        var vNama = vData[v][0] ? vData[v][0].toString().trim() : '';
        var vKab  = vData[v][1] ? normalizeKabName(vData[v][1]) : '';
        if (vNama && vKab && vKab.toLowerCase() === detectedKab.toLowerCase()) {
          allKabVerifs.push(vNama);
        }
      }
    }

    var logCell = sheetLaporan.getRange(rowIndex, 27);
    var currentLogStr = logCell.getValue() ? logCell.getValue().toString() : '{}';
    var currentLogObj = {};
    try { currentLogObj = JSON.parse(currentLogStr || '{}'); } catch(e) {}

    currentLogObj[namaVerifikator] = {
      status: statusVerifikasi,
      catatan: catatan,
      timestamp: Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM-dd HH:mm:ss")
    };

    var finalStatus = 'PENDING';
    var reviewedVerifs = Object.keys(currentLogObj);
    var hasRejected = false;
    var approvedCount = 0;

    reviewedVerifs.forEach(function(vKey) {
      if (currentLogObj[vKey].status === 'REJECTED') hasRejected = true;
      if (currentLogObj[vKey].status === 'APPROVED') approvedCount++;
    });

    if (hasRejected) finalStatus = 'REJECTED';
    else if (allKabVerifs.length > 0 && approvedCount >= allKabVerifs.length) finalStatus = 'APPROVED';
    else finalStatus = 'PENDING';

    var listApprovedNames = reviewedVerifs.join(', ');
    var combinedNotes = reviewedVerifs.map(function(k) { return k + ': ' + currentLogObj[k].catatan; }).join(' | ');

    sheetLaporan.getRange(rowIndex, 24).setValue(finalStatus);
    sheetLaporan.getRange(rowIndex, 25).setValue(combinedNotes);
    sheetLaporan.getRange(rowIndex, 26).setValue(listApprovedNames);
    logCell.setValue(JSON.stringify(currentLogObj));
    
    return { success: true, message: 'Keputusan verifikasi berhasil disimpan!' };
  } catch (error) {
    return { success: false, message: 'Gagal menyimpan keputusan: ' + error.toString() };
  }
}

function deleteMultipleReports(rowIndices, islandName) {
  try {
    var ss = getSpreadsheetByIsland(islandName);
    var sheetLaporan = ss.getSheetByName('RekapLaporan');
    if (!sheetLaporan) return { success: false, message: 'Sheet RekapLaporan tidak ditemukan.' };
    
    rowIndices.sort(function(a, b) { return b - a; });
    var data = sheetLaporan.getDataRange().getValues();
    
    for (var i = 0; i < rowIndices.length; i++) {
      var rIndex = rowIndices[i];
      if (rIndex > 1 && rIndex <= data.length) {
        var rowData = data[rIndex - 1];
        var folderUrl = rowData[12];
        
        if (folderUrl) {
          try {
            var folderIdMatch = folderUrl.match(/folders\/([a-zA-Z0-9-_]+)/);
            if (folderIdMatch && folderIdMatch[1]) {
              DriveApp.getFolderById(folderIdMatch[1]).setTrashed(true);
            }
          } catch (drvErr) {}
        }
        sheetLaporan.deleteRow(rIndex);
      }
    }
    return { success: true, message: 'Data laporan berhasil dihapus.' };
  } catch (error) { return { success: false, message: 'Gagal menghapus data: ' + error.toString() }; }
}
