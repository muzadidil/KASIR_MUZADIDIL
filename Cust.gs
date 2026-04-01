// 4. Simpan Customer Baru
function simpanCustomerBaru(form) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CUST");
  if (!sheet) throw new Error("Sheet CUST tidak ditemukan!");
  sheet.appendRow([form.nama, form.hp, form.alamat]);
  return "Berhasil menambahkan pelanggan: " + form.nama;
}