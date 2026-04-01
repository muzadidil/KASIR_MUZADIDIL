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

// 3. Fungsi Menyimpan Data Mutasi/Keranjang & UPDATE HARGA MASTER
function simpanDataEntry(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetEntry = ss.getSheetByName("MUTASI"); 
  const sheetMaster = ss.getSheetByName("MASTER"); 
  
  if (!sheetEntry) throw new Error("Sheet MUTASI tidak ditemukan!");
  if (!sheetMaster) throw new Error("Sheet MASTER tidak ditemukan!");

  // --- 1. Ambil data MASTER untuk mengecek baris dan harga lama ---
  const dataMaster = sheetMaster.getDataRange().getValues();
  let mapMaster = {}; 
  for(let r = 1; r < dataMaster.length; r++) { 
    let nama = dataMaster[r][1]; // Kolom B (Nama Barang)
    if (nama) {
      mapMaster[nama.toString().trim()] = {
        row: r + 1, 
        jual: dataMaster[r][5] || 0, // Kolom F
        beli: dataMaster[r][6] || 0  // Kolom G
      };
    }
  }

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
    
    // --- 2. Simpan Riwayat ke sheet MUTASI ---
    sheetEntry.getRange(barisBaru, 1).setValue(formattedTime);       
    sheetEntry.getRange(barisBaru, 2).setValue(formattedDate);       
    sheetEntry.getRange(barisBaru, 3).setValue(newNota);             
    sheetEntry.getRange(barisBaru, 4).setValue(payload.customer);    
    // Kolom E dilewati
    sheetEntry.getRange(barisBaru, 6).setValue(item.namaBarang);     
    sheetEntry.getRange(barisBaru, 9).setValue(item.hargaNota); // Simpan Harga Transaksi Aktif        
    sheetEntry.getRange(barisBaru, 10).setValue(item.jumlah);        
    sheetEntry.getRange(barisBaru, 11).setValue(payload.kodeMutasi); 
    
    // --- 3. LOGIKA AUTO-UPDATE KEDUA HARGA DI SHEET MASTER ---
    let masterInfo = mapMaster[item.namaBarang];
    if (masterInfo) {
      // Cek dan Update Harga Jual (Kolom F)
      if (masterInfo.jual != item.hargaJual) {
        sheetMaster.getRange(masterInfo.row, 6).setValue(item.hargaJual);
      }
      // Cek dan Update Harga Beli (Kolom G)
      if (masterInfo.beli != item.hargaBeli) {
        sheetMaster.getRange(masterInfo.row, 7).setValue(item.hargaBeli);
      }
    }

    barisBaru++; 
  }
  return "Berhasil! " + keranjang.length + " barang tersimpan dengan Nota: " + newNota;
}

// 7. Ambil Data MUTASI / RIWAYAT
function getMutasiData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MUTASI");
  if (!sheet) return []; 
  
  const data = sheet.getDataRange().getDisplayValues();
  let result = [];
  
  for (let i = 1; i < data.length; i++) {
    let row = data[i];
    let nota = row[2] ? row[2].toString().trim() : "";
    
    if (nota !== "" && nota.toUpperCase() !== "NOTA") {
      result.push({
        waktu: row[0],     
        tanggal: row[1],   
        nota: nota,        
        customer: row[3],  
        barang: row[5],    
        harga: row[8],     
        qty: row[9],       
        jenis: row[10]     
      });
    }
  }
  return result.reverse(); 
}