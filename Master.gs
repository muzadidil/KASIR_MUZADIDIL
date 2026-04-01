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