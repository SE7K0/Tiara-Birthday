/**
 * setup.gs — Database Initialization & Safe Migration
 * Birthday Special Website for My Future Partner, Tiara
 *
 * Safe Migrate Rules:
 * - Checks getSheetByName() first.
 * - If sheet exists and has data, NEVER uses clear() or deletes existing rows.
 * - Only sets / formats the header row 1.
 */

var SHEET_SCHEMAS = {
  'MEMORIES': ['ID', 'Foto_URL', 'Judul_Momen', 'Tanggal_Momen', 'Deskripsi', 'Kategori_Milestone'],
  'VOICE_NOTES': ['ID', 'Judul', 'Audio_URL', 'Durasi', 'Pengirim', 'Pesan_Pengantar'],
  'CANDLE_WISHES': ['ID', 'Timestamp', 'Harapan_Rahasia', 'Status_Tertiup']
};

/**
 * Mendapatkan spreadsheet yang tersimpan di ScriptProperties atau membuat baru secara otomatis.
 * @return {Spreadsheet}
 */
function getOrCreateSpreadsheet() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');
  var ss = null;

  if (ssId) {
    try {
      ss = SpreadsheetApp.openById(ssId);
    } catch (e) {
      Logger.log('Spreadsheet ID lama tidak ditemukan. Membuat spreadsheet baru...');
      ss = null;
    }
  }

  if (!ss) {
    ss = SpreadsheetApp.create('Database Birthday Special — Tiara');
    props.setProperty('SPREADSHEET_ID', ss.getId());
    Logger.log('Spreadsheet baru berhasil dibuat: ' + ss.getUrl());
    Logger.log('ID: ' + ss.getId());
  }

  return ss;
}

/**
 * Setup seluruh sheet dengan safe migration.
 * Tidak akan pernah menghapus data lama pengguna.
 */
function setupAllSheets() {
  var ss = getOrCreateSpreadsheet();
  var logs = [];

  var sheetNames = Object.keys(SHEET_SCHEMAS);
  for (var i = 0; i < sheetNames.length; i++) {
    var name = sheetNames[i];
    var headers = SHEET_SCHEMAS[name];
    var sheet = ss.getSheetByName(name);

    if (!sheet) {
      sheet = ss.insertSheet(name);
      logs.push(name + ': Sheet baru dibuat.');
    } else {
      logs.push(name + ': Sheet sudah ada.');
    }

    // Hanya perbarui baris header (Row 1)
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

    // Format visual header dengan tema Maroon & Emas
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4A0000');
    headerRange.setFontColor('#FAEBD7');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    for (var col = 1; col <= headers.length; col++) {
      sheet.autoResizeColumn(col);
    }

    var lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      logs.push('  -> ' + (lastRow - 1) + ' baris data aman dipertahankan.');
    }
  }

  // Hapus sheet default jika ada
  var defaultSheet = ss.getSheetByName('Sheet1') || ss.getSheetByName('Sheet 1');
  if (defaultSheet && ss.getSheets().length > 1) {
    try {
      ss.deleteSheet(defaultSheet);
      logs.push('Sheet default dihapus.');
    } catch (e) {}
  }

  SpreadsheetApp.flush();
  Logger.log('=== HASIL SAFE SETUP ===\n' + logs.join('\n'));
  return logs;
}

/**
 * Mengisi data awal/sample jika sheet masih kosong (hanya ada header)
 */
function seedSampleData() {
  var ss = getOrCreateSpreadsheet();
  var seedLogs = [];

  // 1. MEMORIES
  var memSheet = ss.getSheetByName('MEMORIES');
  if (memSheet && memSheet.getLastRow() <= 1) {
    var sampleMemories = [
      [
        'MEM-0001',
        'https://images.unsplash.com/photo-1517841905240-472988babdf9?w=600&auto=format&fit=crop',
        'Pertama Kali Chat di WA',
        '2025-09-23',
        'Sangat canggung tidak ada topik yang menarik.',
        'First Chat'
      ],
      [
        'MEM-0002',
        'https://images.unsplash.com/photo-1529333166437-7750a6dd5a70?w=600&auto=format&fit=crop',
        'PDD Terbaik',
        '2025-10-07',
        'Kamu ngasi aku foto yang keren banget emang kamu pdd terbaik.',
        'PDD Terbaik'
      ],
      [
        'MEM-0003',
        'https://images.unsplash.com/photo-1492562080023-ab3db95bfbce?w=600&auto=format&fit=crop',
        'First Date',
        '2025-10-18',
        'Sebenernya bukan first date sih tapi kan ya..... gitu lah pokoknya.',
        'First Date'
      ],
      [
        'MEM-0004',
        'https://images.unsplash.com/photo-1522673607200-164d1b6ce486?w=600&auto=format&fit=crop',
        'Moment Unforgettable (Valentine Pertama)',
        '2026-02-14',
        'Valentine pertama ku yang bener bener sama pasangan (in my opinion).',
        'Unforgettable'
      ],
      [
        'MEM-0005',
        'https://images.unsplash.com/photo-1464349095431-e9a21285b5f3?w=600&auto=format&fit=crop',
        'Moment Ulang Tahunku',
        '2026-09-06',
        'Kita rayain bareng, baru kali ini ultahku dirayain sama someone special.',
        'My Birthday'
      ],
      [
        'MEM-0006',
        'https://images.unsplash.com/photo-1513151233558-d860c5398176?w=600&auto=format&fit=crop',
        'Moment Ulang Tahunmu',
        '2026-10-05',
        'Mau minta maap karena kemarin nggak put priority on u, se tiba tiba itu di gencet banyak tugas dan permohonan tolong.',
        'Your Birthday'
      ]
    ];
    for (var m = 0; m < sampleMemories.length; m++) {
      memSheet.appendRow(sampleMemories[m]);
    }
    seedLogs.push('MEMORIES diisi ' + sampleMemories.length + ' data sample.');
  }

  // 2. VOICE_NOTES
  var vnSheet = ss.getSheetByName('VOICE_NOTES');
  if (vnSheet && vnSheet.getLastRow() <= 1) {
    var sampleVN = [
      [
        'VN-0001',
        'Pesan Suara Rahasia #1',
        'assets/voicenotes/secret.mp3',
        '01:32',
        'Your Future Partner',
        'Dengerin ini saat kamu lagi kangen ya sayang...'
      ],
      [
        'VN-0002',
        'Doa & Harapan di Hari Ulang Tahunmu',
        'assets/voicenotes/doa.mp3',
        '02:00',
        'Your Future Partner',
        'Pesan rahasia yang direkam khusus menyambut hari bahagiamu.'
      ]
    ];
    for (var v = 0; v < sampleVN.length; v++) {
      vnSheet.appendRow(sampleVN[v]);
    }
    seedLogs.push('VOICE_NOTES diisi ' + sampleVN.length + ' data sample.');
  }

  SpreadsheetApp.flush();
  Logger.log('=== HASIL SEED DATA ===\n' + (seedLogs.length ? seedLogs.join('\n') : 'Semua sheet sudah terisi data.'));
}

/**
 * Jalankan fungsi ini pertama kali di Apps Script Editor
 */
function initializeDatabase() {
  setupAllSheets();
  seedSampleData();
  var ss = getOrCreateSpreadsheet();
  return {
    success: true,
    spreadsheetUrl: ss.getUrl(),
    spreadsheetId: ss.getId()
  };
}
