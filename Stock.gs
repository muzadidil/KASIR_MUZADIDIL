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