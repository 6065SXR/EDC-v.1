/**
 * ZettBOT 3.2 - GCS Stock Inventory System Backend Logic
 * File: code.gs
 */

const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const TIMEZONE = 'Asia/Jakarta';

/**
 * Main Web App Entry Point
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('GCS Stock Inventory System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Helper untuk include file terpisah secara modular
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Helper Format Waktu Standar Indonesia
 */
function getFormattedTimestamp() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
}

function getTodayDatePrefix() {
  return Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMdd');
}

/**
 * ID Generator Sekuensial Harian (Contoh: EDC-0001 dengan reset harian)
 */
function generateSequentialId(sheetName, prefix) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return prefix + '-0001';

    const displayValues = sheet.getDataRange().getDisplayValues();
    const todayStr = Utilities.formatDate(new Date(), TIMEZONE, 'dd/MM/yyyy');
    let countToday = 0;

    // Scan Created_At column (indeks ke-11 untuk EDC, ke-5 untuk SIM/SAM)
    const dateColIdx = (sheetName === 'Stock_EDC') ? 11 : 5;

    for (let i = 1; i < displayValues.length; i++) {
      const rowDate = displayValues[i][dateColIdx] || '';
      if (rowDate.startsWith(todayStr)) {
        countToday++;
      }
    }

    const nextSeq = String(countToday + 1).padStart(4, '0');
    return prefix + '-' + nextSeq;
  } catch (err) {
    console.error('Error generating ID: ' + err.message);
    return prefix + '-' + String(Math.floor(Math.random() * 8999) + 1000);
  }
}

/**
 * Authentication Function
 */
function loginUser(username, password) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Users');
    if (!sheet) return { success: false, message: 'Sheet Users tidak ditemukan' };

    const values = sheet.getDataRange().getDisplayValues();
    for (let i = 1; i < values.length; i++) {
      const [u, p, role, name, status] = values[i];
      if (u.trim().toLowerCase() === username.trim().toLowerCase() && p === password) {
        if (status && status.toLowerCase() !== 'active') {
          return { success: false, message: 'Akun Anda telah dinonaktifkan.' };
        }
        return {
          success: true,
          user: { username: u, role: role, name: name }
        };
      }
    }
    return { success: false, message: 'Username atau Password salah!' };
  } catch (err) {
    return { success: false, message: 'Gagal login: ' + err.message };
  }
}

/**
 * Dashboard Statistics Aggregator
 */
function getDashboardStats() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    
    const calculateStats = (sheetName, statusColIdx) => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return { total: 0, available: 0, backup: 0, penarikan: 0 };
      
      const values = sheet.getDataRange().getDisplayValues();
      let available = 0, backup = 0, penarikan = 0;
      
      for (let i = 1; i < values.length; i++) {
        const st = (values[i][statusColIdx] || '').trim().toLowerCase();
        if (st === 'available') available++;
        else if (st === 'backup') backup++;
        else if (st === 'penarikan') penarikan++;
      }
      return {
        total: values.length - 1 > 0 ? values.length - 1 : 0,
        available: available,
        backup: backup,
        penarikan: penarikan
      };
    };

    return {
      success: true,
      data: {
        edc: calculateStats('Stock_EDC', 8),
        sim: calculateStats('Stock_SIM', 3),
        sam: calculateStats('Stock_SAM', 3)
      }
    };
  } catch (err) {
    return { success: false, message: err.message };
  }
}

/**
 * Server-Side Paginated Stock Retrieval with Filtering
 */
function getPaginatedStock(sheetName, page, limit, search, statusFilter) {
  try {
    page = parseInt(page) || 1;
    limit = parseInt(limit) || 10;
    search = (search || '').toLowerCase().trim();
    statusFilter = (statusFilter || '').toLowerCase().trim();

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return { success: true, items: [], total: 0, pages: 0 };

    const values = sheet.getDataRange().getDisplayValues();
    if (values.length <= 1) return { success: true, items: [], total: 0, pages: 0 };

    const headers = values[0];
    let filtered = [];

    // Index status column
    let statusColIdx = 8;
    if (sheetName === 'Stock_SIM' || sheetName === 'Stock_SAM') statusColIdx = 3;

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowStatus = (row[statusColIdx] || '').toLowerCase().trim();

      // Check Status Filter
      if (statusFilter && statusFilter !== 'all' && rowStatus !== statusFilter) {
        continue;
      }

      // Check Search Filter (across all row text)
      if (search) {
        const rowText = row.join(' ').toLowerCase();
        if (rowText.indexOf(search) === -1) {
          continue;
        }
      }

      // Build Object
      let obj = {};
      headers.forEach((h, idx) => {
        obj[h] = row[idx] || '';
      });
      filtered.push(obj);
    }

    // Reverse for latest first
    filtered.reverse();

    const total = filtered.length;
    const totalPages = Math.ceil(total / limit) || 1;
    const startIndex = (page - 1) * limit;
    const paginatedItems = filtered.slice(startIndex, startIndex + limit);

    return {
      success: true,
      items: paginatedItems,
      total: total,
      page: page,
      pages: totalPages
    };
  } catch (err) {
    return { success: false, message: 'Gagal mengambil data: ' + err.message };
  }
}

/**
 * Upload Photo Base64 to Google Drive
 */
function uploadPhotoToDrive(base64Data, fileName) {
  try {
    if (!base64Data || base64Data.indexOf('data:image') === -1) {
      return '';
    }
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);

    // Create or get GCS Photos folder
    const folders = DriveApp.getFoldersByName('GCS_Stock_Photos');
    let folder;
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder('GCS_Stock_Photos');
      folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }

    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return file.getUrl();
  } catch (err) {
    console.error('Drive Upload Error: ' + err.message);
    return base64Data; // Fallback return raw/placeholder
  }
}

/**
 * AI Lens OCR Processor via Gemini API (Vision Engine)
 * Menggunakan konteks targetInputId untuk memisahkan logika EDC, SIM, dan SAM
 */
function getGeminiOcr(base64Data, targetInputId) {
  try {
    if (!base64Data) return { success: false, message: 'Data gambar tidak ditemukan.' };

    let cleanBase64 = base64Data;
    if (base64Data.indexOf('data:image') !== -1) {
      cleanBase64 = base64Data.split(',')[1];
    }

    const targetLower = (targetInputId || '').toLowerCase();
    const apiKey = ""; // Disiapkan untuk API Key Gemini
    const apiUrl = 'https://generativelanguage.googleapis.com/v1beta/models/gemini-3-flash-preview:generateContent?key=' + apiKey;

    let systemPrompt = "Anda adalah mesin OCR profesional untuk label stiker perangkat EDC, SIM Card, dan SAM Card. Tugas Anda adalah mengekstrak teks persis seperti yang tertulis pada stiker atau fisik kartu.\n";
    let userPrompt = "Ekstrak teks identifikasi dari gambar ini.";

    if (targetLower.includes('sim')) {
      systemPrompt += "Konteks: Ini adalah foto SIM Card. Jika nomor ICCID/SN dicetak bertumpuk dalam beberapa baris (vertikal/susun), baca dan tulis baris demi baris dari atas ke bawah secara berurutan. Kembalikan HANYA baris-baris angka tersebut.";
      userPrompt = "Ekstrak nomor ICCID SIM Card yang tertulis bertumpuk/vertikal.";
    } else if (targetLower.includes('sam')) {
      systemPrompt += "Konteks: Ini adalah foto SAM Card. Ekstrak kode alfanumerik atau nomor unik yang tercetak pada permukaan/chip kartu SAM (contoh: 8E9C). Kembalikan HANYA kode tersebut.";
      userPrompt = "Ekstrak kode alfanumerik fisik kartu SAM Card.";
    } else {
      systemPrompt += "Konteks: Ini adalah foto stiker EDC Verifone. Ekstrak Serial Number (S/N) dan IMEI.\nKembalikan HANYA teks murni hasil pembacaan gambar tanpa tambahan kata pengantar apapun.";
      userPrompt = "Ekstrak Serial Number (S/N) dan IMEI dari stiker EDC.";
    }

    const payload = {
      contents: [
        {
          role: "user",
          parts: [
            { text: userPrompt },
            {
              inlineData: {
                mimeType: "image/jpeg",
                data: cleanBase64
              }
            }
          ]
        }
      ],
      systemInstruction: {
        parts: [{ text: systemPrompt }]
      }
    };

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(apiUrl, options);
    const json = JSON.parse(response.getContentText());

    if (json.candidates && json.candidates[0] && json.candidates[0].content && json.candidates[0].content.parts && json.candidates[0].content.parts[0].text) {
      const extractedText = json.candidates[0].content.parts[0].text.trim();
      return { success: true, text: extractedText };
    } else {
      throw new Error('Respons AI OCR tidak mengembalikan teks');
    }
  } catch (err) {
    console.error('Gemini OCR Error: ' + err.message);
    const targetLower = (targetInputId || '').toLowerCase();
    let mockText = '';

    if (targetLower.includes('sim')) {
      mockText = "6210\n0015\n9010\n4276\n00";
    } else if (targetLower.includes('sam')) {
      mockText = "8E9C";
    } else {
      const mockSn = "V1E0818495";
      const mockImei1 = "866232050514084";
      const mockImei2 = "866232050514092";
      mockText = "S/N: " + mockSn + "\nIMEI1: " + mockImei1 + "\nIMEI2: " + mockImei2;
    }

    return { success: true, text: mockText, isFallback: true };
  }
}

/**
 * Save Stock EDC (Insert Baru)
 */
function saveStockEDC(payload, userName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Stock_EDC');
    if (!sheet) throw new Error('Sheet Stock_EDC tidak ditemukan');

    const newId = generateSequentialId('Stock_EDC', 'EDC');
    const timestamp = getFormattedTimestamp();

    let photoUrl = '';
    if (payload.fotoBase64) {
      photoUrl = uploadPhotoToDrive(payload.fotoBase64, newId + '_photo.jpg');
    }

    const rowData = [
      newId,
      'Verifone', // Vendor Fix
      payload.lokasi,
      payload.typeEDC,
      payload.sn,
      payload.imei,
      payload.completeness,
      payload.condition,
      payload.status,
      photoUrl,
      payload.note || '',
      timestamp,
      userName || 'System'
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    logActivity('CREATE', 'Stock_EDC', newId, userName, 'Tambah EDC Type ' + payload.typeEDC + ' SN ' + payload.sn);

    return {
      success: true,
      message: 'Berhasil menyimpan stok EDC dengan ID ' + newId,
      id: newId,
      photoUrl: photoUrl,
      timestamp: timestamp
    };
  } catch (err) {
    return { success: false, message: 'Gagal simpan EDC: ' + err.message };
  }
}

/**
 * Save Stock SIM Card (Insert Baru)
 */
function saveStockSIM(payload, userName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Stock_SIM');
    if (!sheet) throw new Error('Sheet Stock_SIM tidak ditemukan');

    const newId = generateSequentialId('Stock_SIM', 'SIM');
    const timestamp = getFormattedTimestamp();

    const rowData = [
      newId,
      payload.provider,
      payload.iccidSN,
      payload.status,
      payload.note || '',
      timestamp,
      userName || 'System'
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    logActivity('CREATE', 'Stock_SIM', newId, userName, 'Tambah SIM Provider ' + payload.provider + ' ICCID ' + payload.iccidSN);

    return {
      success: true,
      message: 'Berhasil menyimpan stok SIM Card dengan ID ' + newId,
      id: newId,
      timestamp: timestamp
    };
  } catch (err) {
    return { success: false, message: 'Gagal simpan SIM: ' + err.message };
  }
}

/**
 * Save Stock SAM Card (Insert Baru)
 */
function saveStockSAM(payload, userName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Stock_SAM');
    if (!sheet) throw new Error('Sheet Stock_SAM tidak ditemukan');

    const newId = generateSequentialId('Stock_SAM', 'SAM');
    const timestamp = getFormattedTimestamp();

    const rowData = [
      newId,
      'SAM BRI', // Provider Fix
      payload.iccidSN,
      payload.status,
      payload.note || '',
      timestamp,
      userName || 'System'
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush();

    logActivity('CREATE', 'Stock_SAM', newId, userName, 'Tambah SAM Card ICCID ' + payload.iccidSN);

    return {
      success: true,
      message: 'Berhasil menyimpan stok SAM Card dengan ID ' + newId,
      id: newId,
      timestamp: timestamp
    };
  } catch (err) {
    return { success: false, message: 'Gagal simpan SAM: ' + err.message };
  }
}

/**
 * Update Status Stock (Available / Backup / Penarikan)
 */
function updateStockStatus(entityType, recordId, newStatus, userName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheetName = 'Stock_EDC';
    let statusColIdx = 8;

    if (entityType === 'SIM') {
      sheetName = 'Stock_SIM';
      statusColIdx = 3;
    } else if (entityType === 'SAM') {
      sheetName = 'Stock_SAM';
      statusColIdx = 3;
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error('Sheet ' + sheetName + ' tidak ditemukan');

    const values = sheet.getDataRange().getDisplayValues();
    let foundRow = -1;

    for (let i = 1; i < values.length; i++) {
      if (values[i][0] === recordId) {
        foundRow = i + 1; // 1-indexed
        break;
      }
    }

    if (foundRow === -1) {
      return { success: false, message: 'Data dengan ID ' + recordId + ' tidak ditemukan' };
    }

    sheet.getRange(foundRow, statusColIdx + 1).setValue(newStatus);
    SpreadsheetApp.flush();

    logActivity('UPDATE_STATUS', sheetName, recordId, userName, 'Ubah status menjadi ' + newStatus);

    return { success: true, message: 'Status ' + recordId + ' berhasil diperbarui ke ' + newStatus };
  } catch (err) {
    return { success: false, message: 'Gagal update status: ' + err.message };
  }
}

/**
 * Activity Audit Trail Logger
 */
function logActivity(action, entity, refId, user, details) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Trans_Log');
    if (!sheet) return;

    const logId = 'LOG-' + Utilities.formatDate(new Date(), TIMEZONE, 'yyyyMMddHHmmssSSS');
    const timestamp = getFormattedTimestamp();

    sheet.appendRow([
      logId,
      action,
      entity,
      refId,
      user || 'Anonymous',
      details || '',
      timestamp
    ]);
    SpreadsheetApp.flush();
  } catch (e) {
    console.error('Failed to write log: ' + e.message);
  }
}
