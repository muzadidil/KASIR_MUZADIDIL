function doGet() {
  // Menggunakan createTemplateFromFile dan evaluate() agar JS.html bisa terbaca
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('AEMITRA INVENTORY')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// Fungsi wajib untuk memanggil file JS.html ke dalam Index.html
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// 1. Fungsi Mengambil Daftar Nama Barang & Harga (Melibatkan MASTER & CUST)
function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMaster = ss.getSheetByName("MASTER"); 
  let dataBarangLengkap = {}; 
  
  if (sheetMaster) {
    const dataMaster = sheetMaster.getDataRange().getValues();
    dataMaster.shift(); 
    
    dataMaster.forEach(row => {
      let namaBarang = row[1]; // Kolom B
      let kategori = row[2] || "-"; // Kolom C
      let unit = row[3] || "-"; // Kolom D
      let hargaJual = row[5] || 0;  // Kolom F
      let hargaBeli = row[6] || 0;  // Kolom G
      
      if (namaBarang && namaBarang.toString().trim() !== "") {
        dataBarangLengkap[namaBarang.toString().trim()] = {
          kategori: kategori,
          unit: unit,
          jual: hargaJual,
          beli: hargaBeli
        };
      }
    });
  }
  
  const sheetCust = ss.getSheetByName("CUST");
  let listCustomer = [];
  if (sheetCust) {
    const dataCust = sheetCust.getDataRange().getValues();
    dataCust.shift(); 
    dataCust.forEach(row => {
      let namaCust = row[0]; 
      if (namaCust && namaCust.toString().trim() !== "") {
        listCustomer.push(namaCust.toString().trim());
      }
    });
  }

  return {
    barangDict: dataBarangLengkap, 
    customer: [...new Set(listCustomer)]
  };
}