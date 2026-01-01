/**
 * Serves the HTML file for the web app.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Sales Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Fetches data from the defined Spreadsheet and Sheet.
 * returns {string} JSON string of the processed data
 */
function getSalesData() {
  const spreadsheetId = '1ntLJJR_V7SEXDLuh2FMLp-dNgpanx6-yF7oKHbN-s08';
  const sheetName = 'Raw Faktur';
  
  try {
    const ss = SpreadsheetApp.openById(spreadsheetId);
    const sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`Sheet "${sheetName}" not found.`);
    }

    // Get all data
    const data = sheet.getDataRange().getDisplayValues();
    const headers = data.shift();
    
    // Column Mapping based on user description:
    // 0:INVOICENO, 1:INVOICEDATE, 2:NAME, 3:SALESMAN, 4:ADDRESS1, 5:ADDRESS2, 
    // 6:CITY, 7:Tipe Pelanggan, 8:Tipe Produk, 9:Kategori Customer, 10:Sales Amount
    
    const processedData = data.map(row => {
      // NUMERIC HANDLING:
      // Indonesian format often uses "." as thousands separator (e.g., 1.500.000)
      // We remove the dots to make it a valid JS number string (1500000)
      let amountStr = row[10] || "0";
      let cleanAmount = amountStr.replace(/\./g, '').replace(/,/g, '.');
      
      // Parse Date (Assuming format like YYYY-MM-DD or DD/MM/YYYY)
      // We will pass the raw string and let the frontend parse it to ensure timezone consistency
      
      return {
        invoiceNo: row[0],
        date: row[1],
        customer: row[2], // NAME
        salesman: row[3],
        city: row[6],
        customerType: row[7],
        productType: row[8],
        category: row[9],
        amount: parseFloat(cleanAmount) || 0
      };
    }).filter(item => item.invoiceNo !== ""); 

    return JSON.stringify({
      status: 'success',
      data: processedData
    });

  } catch (e) {
    return JSON.stringify({
      status: 'error',
      message: e.toString()
    });
  }
}


function parseIndoDate(str) {
  if (!str) return null;

  const bulan = {
    Jan: 0, Feb: 1, Mar: 2, Apr: 3, Mei: 4, Jun: 5,
    Jul: 6, Agu: 7, Sep: 8, Okt: 9, Nov: 10, Des: 11
  };

  const parts = str.trim().split(' ');
  if (parts.length !== 3) return null;

  const day = parseInt(parts[0], 10);
  const month = bulan[parts[1]];
  let year = parseInt(parts[2], 10);

  // 24 → 2024
  year = year < 100 ? 2000 + year : year;

  const d = new Date(year, month, day);
  return isNaN(d) ? null : Utilities.formatDate(d, 'GMT+7', 'yyyy-MM-dd');
}




function getOmzetVsCostData() {
  const spreadsheetId = '1ntLJJR_V7SEXDLuh2FMLp-dNgpanx6-yF7oKHbN-s08';
  const sheetName = 'DATA: Transaksi';
  const rangeA1 = 'I6:Q'; // ✅ FINAL

  try {
    const sheet = SpreadsheetApp
      .openById(spreadsheetId)
      .getSheetByName(sheetName);

    if (!sheet) throw new Error('Sheet DATA: Transaksi tidak ditemukan');

    const data = sheet.getRange(rangeA1).getDisplayValues();

    const result = data
      .slice(1)
      .filter(r => r[0]) // kolom I = Tanggal
      .map(r => ({
        tanggal: parseIndoDate(r[0]),                    // I
        customer: r[1],                   // J
        nominal: Number(
          (r[2] || '0').replace(/[^\d]/g, '')
        ) || 0,
        sales: r[3],                      // L
        divisi: r[4],                     // M
        jenisAjuan: r[5],                 // N
        metodePembayaran: r[6],           // O
        deskripsi: r[7],                  // P
        progress: r[8]                    // Q
      }));

    return {
      status: 'success',
      data: result
    };

  } catch (err) {
    return {
      status: 'error',
      message: err.message
    };
  }
}




