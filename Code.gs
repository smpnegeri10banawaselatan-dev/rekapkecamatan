function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('CSV Proxy + Tabel');
}

// Ambil CSV dari URL publik
function fetchCsvData() {
  const url = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vR_BmCgfv-HOMRKzNMZ1aHL05CZsR7Ie0VZGRJTD29tX1C-ANiKASW1hB_PM69WuQ/pub?gid=1519352447&single=true&output=csv';
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  if (response.getResponseCode() !== 200) {
    throw new Error('Gagal fetch CSV, status: ' + response.getResponseCode());
  }
  return response.getContentText();
}

// Fungsi bikin file Excel sesuai tampilan HTML dan kembalikan URL download-nya
function showDownloadDialog() {
  const file = createExcelFile();
  // Buat link download XLSX
  const url = file.getUrl().replace(/edit$/, 'export?format=xlsx');
  return url;
}

function createExcelFile() {
  const csvUrl = 'https://docs.google.com/spreadsheets/d/e/2PACX-1vR_BmCgfv-HOMRKzNMZ1aHL05CZsR7Ie0VZGRJTD29tX1C-ANiKASW1hB_PM69WuQ/pub?gid=1519352447&single=true&output=csv';
  const csvContent = UrlFetchApp.fetch(csvUrl).getContentText();
  const csvData = Utilities.parseCsv(csvContent);

  // Buat spreadsheet baru
  const ss = SpreadsheetApp.create('Export Rekap Kecamatan Dampelas');
  const sheet = ss.getActiveSheet();
  sheet.clear();

  // Setup header gabungan sesuai HTML
  // Baris 1
  sheet.getRange('A1').setValue('PENGGUNA BARANG');
  sheet.getRange('B1').setValue('DITEMUKAN');
  sheet.getRange('E1').setValue('JUMLAH DITEMUKAN');
  sheet.getRange('F1').setValue('JUMLAH TOTAL');

  // Merge cells header sesuai
  sheet.getRange('A1:A2').merge();
  sheet.getRange('B1:D1').merge();
  sheet.getRange('E1:E2').merge();
  sheet.getRange('F1:G1').merge();

  // Baris 2
  sheet.getRange('B2').setValue('BAIK');
  sheet.getRange('C2').setValue('RUSAK BERAT');
  sheet.getRange('D2').setValue('RUSAK RINGAN');
  sheet.getRange('F2').setValue('TIDAK DITEMUKAN');
  sheet.getRange('G2').setValue('JUMLAH TOTAL');

  // Mulai tulis data mulai baris 3 (karena 2 baris header)
  for (let i = 4; i < csvData.length; i++) {  // index 4 = baris 5 CSV (sesuai html)
    const row = csvData[i];
    const dataRow = [
      row[0] || '',        // A
      parseNumber(row[3]), // D (BAIK)
      parseNumber(row[4]), // E (RUSAK BERAT)
      parseNumber(row[5]), // F (RUSAK RINGAN)
      parseNumber(row[6]), // G (JUMLAH DITEMUKAN)
      parseNumber(row[8]), // I (TIDAK DITEMUKAN)
      parseNumber(row[9])  // J (JUMLAH TOTAL)
    ];
    sheet.getRange(i - 1, 1, 1, dataRow.length).setValues([dataRow]);
  }

  // Format header bold + center align
  sheet.getRange('A1:G2').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');

  // Format border untuk header dan data
  const lastRow = csvData.length - 1;
  sheet.getRange(1, 1, lastRow + 1, 7).setBorder(true, true, true, true, true, true);

  // Format number to Rupiah for data cells (rows 3 to lastRow+1, columns B to G)
  sheet.getRange(3, 2, lastRow - 2, 6).setNumberFormat('"Rp." #,##0');

  // Autofit columns
  sheet.autoResizeColumns(1, 7);

  // Return file Drive
  return DriveApp.getFileById(ss.getId());
}

// Helper untuk parsing string angka menjadi Number, kosong jika tidak valid
function parseNumber(value) {
  if (!value) return null;
  // Hapus titik ribuan dan ganti koma desimal (jika ada)
  const normalized = value.replace(/\./g, '').replace(/,/g, '.');
  const n = Number(normalized);
  return isNaN(n) ? null : n;
}
