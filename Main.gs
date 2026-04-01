function doGet() {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('AEMITRA INVENTORY')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMaster = ss.getSheetByName("MASTER"); 
  let dataBarangLengkap = {}; 
  let listKategori = [];
  let listUnit = [];
  
  if (sheetMaster) {
    const dataMaster = sheetMaster.getDataRange().getValues();
    dataMaster.shift(); 
    
    dataMaster.forEach(row => {
      let namaBarang = row[1]; // Kolom B
      let kategori = row[2] ? row[2].toString().trim() : ""; // Kolom C
      let unit = row[3] ? row[3].toString().trim() : ""; // Kolom D
      let hargaJual = row[5] || 0;  // Kolom F
      let hargaBeli = row[6] || 0;  // Kolom G
      
      if (namaBarang && namaBarang.toString().trim() !== "") {
        dataBarangLengkap[namaBarang.toString().trim()] = {
          kategori: kategori || "-",
          unit: unit || "-",
          jual: hargaJual,
          beli: hargaBeli
        };
      }
      
      // Kumpulkan kategori & unit unik
      if (kategori !== "" && kategori !== "-") listKategori.push(kategori);
      if (unit !== "" && unit !== "-") listUnit.push(unit);
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
    customer: [...new Set(listCustomer)],
    kategori: [...new Set(listKategori)], // Kirim ke dropdown UI
    unit: [...new Set(listUnit)]          // Kirim ke dropdown UI
  };
}