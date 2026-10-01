/**
 * ZettBOT 3.2 - GCS Stock Inventory System Database Setup & Safe Migrate
 * File: setup.gs
 */

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Sheet Schemas & Headers
  const schema = {
    'Users': ['Username', 'Password', 'Role', 'Nama', 'Status'],
    'Stock_EDC': ['ID_EDC', 'Vendor', 'Lokasi', 'Type_EDC', 'SN', 'IMEI', 'Completeness', 'Condition', 'Status', 'Foto_Drive_URL', 'Note', 'Created_At', 'Created_By'],
    'Stock_SIM': ['ID_SIM', 'Provider', 'ICCID_SN', 'Status', 'Note', 'Created_At', 'Created_By'],
    'Stock_SAM': ['ID_SAM', 'Provider', 'ICCID_SN', 'Status', 'Note', 'Created_At', 'Created_By'],
    'Trans_Log': ['Log_ID', 'Action', 'Entity', 'Reference_ID', 'User', 'Details', 'Timestamp']
  };

  Object.keys(schema).forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    const targetHeaders = schema[sheetName];

    if (!sheet) {
      // Create new sheet
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
      sheet.getRange(1, 1, 1, targetHeaders.length)
           .setFontWeight('bold')
           .setBackground('#0088ff')
           .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
    } else {
      // Safe Migrate Header check: DO NOT clear rows!
      const lastCol = sheet.getLastColumn();
      if (lastCol === 0) {
        sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
        sheet.getRange(1, 1, 1, targetHeaders.length)
             .setFontWeight('bold')
             .setBackground('#0088ff')
             .setFontColor('#ffffff');
        sheet.setFrozenRows(1);
      } else {
        const currentHeaders = sheet.getRange(1, 1, 1, Math.max(lastCol, targetHeaders.length)).getDisplayValues()[0];
        // Ensure missing headers are appended without altering data
        targetHeaders.forEach((hdr, idx) => {
          if (!currentHeaders[idx] || currentHeaders[idx] !== hdr) {
            sheet.getRange(1, idx + 1).setValue(hdr);
          }
        });
      }
    }
  });

  const userSheet = ss.getSheetByName('Users');
  if (userSheet && userSheet.getLastRow() <= 1) {
    const defaultUsers = [
      ['superuser', 'admin123', 'Super User', 'Super Administrator', 'Active'],
      ['admin', 'admin123', 'Admin', 'Admin Gudang GCS', 'Active'],
      ['engineer', 'eng123', 'Engineer', 'Teknisi Field GCS', 'Active'],
      ['customer', 'cust123', 'Customer', 'Client Partner', 'Active']
    ];
    userSheet.getRange(2, 1, defaultUsers.length, 5).setValues(defaultUsers);
  }

  seedDummyStockData(ss);

  SpreadsheetApp.flush();
  Logger.log('✅ Setup & Safe Migration database GCS Stock Inventory berhasil diselesaikan!');
}

/**
 * Helper Seeding Data Dummy untuk kebutuhan preview awal
 */
function seedDummyStockData(ss) {
  const today = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');

  // EDC Dummy
  const edcSheet = ss.getSheetByName('Stock_EDC');
  if (edcSheet && edcSheet.getLastRow() <= 1) {
    const dummyEDC = [
      ['EDC-0001', 'Verifone', 'Jakarta', 'V2 Colour', 'SN-VER-882191', '356781092837101', 'Complete', 'New', 'Available', '', 'Stock awal gudang utama', today, 'superuser'],
      ['EDC-0002', 'Verifone', 'Bandung', 'V4 Colour', 'SN-VER-882192', '356781092837102', 'Complete', 'Good', 'Backup', '', 'Persiapan tim support', today, 'admin'],
      ['EDC-0003', 'Verifone', 'Surabaya', 'V2 Black White', 'SN-VER-882193', '356781092837103', 'Incomplete', 'Damage', 'Penarikan', '', 'Kabel power hilang & layar rusak', today, 'engineer']
    ];
    edcSheet.getRange(2, 1, dummyEDC.length, dummyEDC[0].length).setValues(dummyEDC);
  }

  // SIM Dummy
  const simSheet = ss.getSheetByName('Stock_SIM');
  if (simSheet && simSheet.getLastRow() <= 1) {
    const dummySIM = [
      ['SIM-0001', 'Telkomsel', '8962010198273641', 'Available', 'Ready pasang EDC', today, 'admin'],
      ['SIM-0002', 'Indosat', '8962020918273642', 'Backup', 'Cadangan event', today, 'admin'],
      ['SIM-0003', 'Telkomsel', '8962010827364109', 'Penarikan', 'Kartu rusak/deaktif', today, 'engineer']
    ];
    simSheet.getRange(2, 1, dummySIM.length, dummySIM[0].length).setValues(dummySIM);
  }

  // SAM Dummy
  const samSheet = ss.getSheetByName('Stock_SAM');
  if (samSheet && samSheet.getLastRow() <= 1) {
    const dummySAM = [
      ['SAM-0001', 'SAM BRI', '890100291827361', 'Available', 'Master Key Aktif', today, 'superuser'],
      ['SAM-0002', 'SAM BRI', '890100291827362', 'Backup', 'Stock backup cabang', today, 'admin']
    ];
    samSheet.getRange(2, 1, dummySAM.length, dummySAM[0].length).setValues(dummySAM);
  }
}
