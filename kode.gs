function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('AEMITRA INVENTORY')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

// 1. Fungsi Mengambil Daftar Nama Barang & Harga
function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetMaster = ss.getSheetByName("MASTER"); 
  let dataBarangLengkap = {}; 
  
  if (sheetMaster) {
    const dataMaster = sheetMaster.getDataRange().getValues();
    dataMaster.shift(); 
    
    dataMaster.forEach(row => {
      let namaBarang = row[1]; // Kolom B
      let hargaJual = row[5] || 0;  // Kolom F
      let hargaBeli = row[6] || 0;  // Kolom G
      
      if (namaBarang && namaBarang.toString().trim() !== "") {
        dataBarangLengkap[namaBarang.toString().trim()] = {
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

// 2. Fungsi Membuat Nomor Nota Otomatis
function generateNotaNumber(sheet, barisBaru) {
  const dataC = sheet.getRange("C:C").getValues(); 
  let lastNota = "";
  for (let i = barisBaru - 2; i >= 0; i--) {
    if (dataC[i][0] && dataC[i][0].toString().startsWith("SH")) {
      lastNota = dataC[i][0].toString();
      break;
    }
  }
  if (lastNota === "") return "SH00001"; 
  const numStr = lastNota.replace("SH", "");
  let nextNum = 1;
  if (!isNaN(numStr) && numStr !== "") {
    nextNum = parseInt(numStr, 10) + 1;
  }
  let paddedNum = nextNum.toString();
  while (paddedNum.length < 5) paddedNum = "0" + paddedNum;
  return "SH" + paddedNum;
}

// 3. Fungsi Menyimpan Data Mutasi/Keranjang
function simpanDataEntry(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEntry = ss.getSheetByName("MUTASI"); 
  if (!sheetEntry) throw new Error("Sheet MUTASI tidak ditemukan!");

  let rawDate = new Date();
  let formattedTime = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "HH:mm:ss");
  let formattedDate = Utilities.formatDate(rawDate, Session.getScriptTimeZone(), "dd-MMM-yyyy");

  const kolomC = sheetEntry.getRange("C:C").getValues();
  let barisTerakhir = 1; 
  for (let i = 0; i < kolomC.length; i++) {
    if (kolomC[i][0] !== "") barisTerakhir = i + 1;
  }
  let barisBaru = barisTerakhir + 1; 

  const newNota = generateNotaNumber(sheetEntry, barisBaru);
  const keranjang = payload.dataKeranjang;
  
  for(let i = 0; i < keranjang.length; i++) {
    let item = keranjang[i];
    sheetEntry.getRange(barisBaru, 1).setValue(formattedTime);       
    sheetEntry.getRange(barisBaru, 2).setValue(formattedDate);       
    sheetEntry.getRange(barisBaru, 3).setValue(newNota);             
    sheetEntry.getRange(barisBaru, 4).setValue(payload.customer);    
    // Kolom E dilewati
    sheetEntry.getRange(barisBaru, 6).setValue(item.namaBarang);     
    sheetEntry.getRange(barisBaru, 9).setValue(item.harga);          
    sheetEntry.getRange(barisBaru, 10).setValue(item.jumlah);        
    sheetEntry.getRange(barisBaru, 11).setValue(payload.kodeMutasi); 
    barisBaru++; 
  }
  return "Berhasil! " + keranjang.length + " barang tersimpan dengan Nota: " + newNota;
}

// 4. Simpan Customer Baru
function simpanCustomerBaru(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CUST");
  if (!sheet) throw new Error("Sheet CUST tidak ditemukan!");
  sheet.appendRow([form.nama, form.hp, form.alamat]);
  return "Berhasil menambahkan pelanggan: " + form.nama;
}

// 5. Simpan Barang Baru
function simpanBarangBaru(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MASTER");
  if (!sheet) throw new Error("Sheet MASTER tidak ditemukan!");
  let barisBaru = [
    form.kode,        
    form.nama,        
    form.kategori,    
    form.unit,        
    "",               
    form.hargaJual,   
    form.hargaBeli    
  ];
  sheet.appendRow(barisBaru);
  return "Berhasil menambahkan barang: " + form.nama;
}

// 6. Ambil Data STOCK
function getStokData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("STOCK");
  if (!sheet) return []; 
  
  const data = sheet.getDataRange().getValues();
  let result = [];
  
  data.forEach(row => {
    let namaBarang = row[1] ? row[1].toString().trim() : "";
    if (namaBarang !== "" && namaBarang.toUpperCase() !== "NAMA BARANG") {
      result.push({
        kolomB: namaBarang,
        kolomF: row[5] || 0,
        kolomG: row[6] || 0,
        kolomH: row[7] || 0,
        kolomI: row[8] || 0
      });
    }
  });
  return result; 
}

// 7. Ambil Data MUTASI / RIWAYAT (BARU)
function getMutasiData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MUTASI");
  if (!sheet) return []; 
  
  // Gunakan getDisplayValues agar format tanggal & jam persis seperti di layar
  const data = sheet.getDataRange().getDisplayValues();
  let result = [];
  
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let nota = row[2] ? row[2].toString().trim() : "";
    
    if (nota !== "" && nota.toUpperCase() !== "NOTA") {
      result.push({
        waktu: row[0],     // Kolom A
        tanggal: row[1],   // Kolom B
        nota: nota,        // Kolom C
        customer: row[3],  // Kolom D
        barang: row[5],    // Kolom F
        harga: row[8],     // Kolom I
        qty: row[9],       // Kolom J
        jenis: row[10]     // Kolom K (IN/OUT)
      });
    }
  }
  
  // Dibalik agar transaksi terbaru muncul paling atas
  return result.reverse(); 
}
