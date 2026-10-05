/**
 * code.gs — Backend API & Web App Controller
 * Birthday Special Website for My Future Partner, Tiara
 *
 * Rules:
 * - Always uses sheet.getDisplayValues() (NEVER getValues).
 * - Always calls SpreadsheetApp.flush() after write mutations.
 * - Extracts and accepts only primitive arguments (String, Number, Boolean).
 * - Implements server-side pagination & filtering to prevent lag on mobile devices.
 * - Exports native PDF via DriveApp & HtmlService without external paid libraries.
 */

/**
 * Entry point Google Apps Script Web App
 */
function doGet(e) {
  try {
    setupAllSheets();
  } catch (err) {
    Logger.log('Setup notice: ' + err.toString());
  }

  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Happy Birthday, My Future Partner, Tiara! \uD83C\uDF82\uD83D\uDC96')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

/**
 * Mendapatkan spreadsheet aktif
 * @private
 */
function getSpreadsheet_() {
  var props = PropertiesService.getScriptProperties();
  var ssId = props.getProperty('SPREADSHEET_ID');

  if (!ssId) {
    setupAllSheets();
    ssId = props.getProperty('SPREADSHEET_ID');
  }

  return SpreadsheetApp.openById(ssId);
}

/**
 * Mendapatkan sheet berdasarkan nama
 * @private
 */
function getSheet_(sheetName) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    setupAllSheets();
    sheet = ss.getSheetByName(sheetName);
  }
  return sheet;
}

/**
 * Generator ID sekuensial aman (format: PREFIX-0001)
 * @private
 */
function generateId_(prefix, sheet) {
  var lastRow = sheet.getLastRow();
  var count = (lastRow > 1) ? (lastRow - 1) : 0;
  var nextNumber = count + 1;
  var padded = ('0000' + String(nextNumber)).slice(-4);
  return prefix + '-' + padded;
}

/* ==========================================================================
   MUTATION APIS (WRITE)
   ========================================================================== */

/**
 * Menyimpan permohonan rahasia tiup lilin ke sheet CANDLE_WISHES
 * @param {string} wish - Teks harapan rahasia
 * @param {string} blownStatus - Status tiup ('true'/'false')
 * @return {Object}
 */
function submitCandleWish(wish, blownStatus) {
  var sheet = getSheet_('CANDLE_WISHES');
  var id = generateId_('WISH-CANDLE', sheet);
  var timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'Asia/Jakarta',
    'yyyy-MM-dd HH:mm:ss'
  );

  var cleanWish = String(wish || '').trim();
  var cleanStatus = String(blownStatus || 'true');

  sheet.appendRow([
    String(id),
    String(timestamp),
    cleanWish,
    cleanStatus
  ]);

  SpreadsheetApp.flush();

  return {
    success: true,
    id: String(id),
    timestamp: String(timestamp)
  };
}

/* ==========================================================================
   QUERY APIS (READ WITH SERVER-SIDE PAGINATION & FILTERING)
   ========================================================================== */

/**
 * Helper generic pembaca data sheet dengan getDisplayValues() dan pagination
 * @private
 */
function getPaginatedData_(sheetName, page, pageSize, filterCol, filterVal) {
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(sheetName);

  if (!sheet || sheet.getLastRow() <= 1) {
    return {
      data: [],
      page: 1,
      totalPages: 0,
      totalItems: 0
    };
  }

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  // WAJIB: getDisplayValues() agar seluruh tanggal, jam, dan angka terbaca sebagai String murni
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0];
  var rows = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  // Server-side filtering jika ada parameter filter
  if (filterCol && filterVal && String(filterVal) !== 'all') {
    var colIdx = -1;
    for (var h = 0; h < headers.length; h++) {
      if (headers[h] === filterCol) {
        colIdx = h;
        break;
      }
    }
    if (colIdx >= 0) {
      var filteredRows = [];
      for (var r = 0; r < rows.length; r++) {
        if (rows[r][colIdx] === String(filterVal)) {
          filteredRows.push(rows[r]);
        }
      }
      rows = filteredRows;
    }
  }

  var totalItems = rows.length;
  var safePageSize = Number(pageSize) || 6;
  var totalPages = Math.ceil(totalItems / safePageSize) || 1;
  var safePage = Math.max(1, Math.min(Number(page) || 1, totalPages));

  var startIndex = (safePage - 1) * safePageSize;
  var sliceRows = rows.slice(startIndex, startIndex + safePageSize);

  var items = [];
  for (var i = 0; i < sliceRows.length; i++) {
    var rowObj = {};
    for (var c = 0; c < headers.length; c++) {
      rowObj[headers[c]] = String(sliceRows[i][c] || '');
    }
    items.push(rowObj);
  }

  return {
    data: items,
    page: Number(safePage),
    totalPages: Number(totalPages),
    totalItems: Number(totalItems)
  };
}

/**
 * Mengambil data Galeri Kenangan
 */
function getMemories(page, pageSize, category) {
  var p = Number(page) || 1;
  var s = Number(pageSize) || 6;
  var c = category ? String(category) : '';
  return getPaginatedData_('MEMORIES', p, s, 'Kategori_Milestone', c);
}

/**
 * Mengambil data Linimasa Milestone Road
 */
function getTimeline(page, pageSize) {
  var p = Number(page) || 1;
  var s = Number(pageSize) || 10;
  return getPaginatedData_('MEMORIES', p, s, null, null);
}

/**
 * Mengambil data Voice Notes
 */
function getVoiceNotes(page, pageSize) {
  var p = Number(page) || 1;
  var s = Number(pageSize) || 6;
  return getPaginatedData_('VOICE_NOTES', p, s, null, null);
}

/* ==========================================================================
   DIGITAL KEEPSAKE PDF EXPORTER
   ========================================================================== */

/**
 * Menyusun kompilasi kenangan dan harapan rahasia menjadi dokumen PDF elegan
 * Menggunakan DriveApp & Utilities native tanpa library berbayar
 * @return {string} Download URL file PDF di Google Drive
 */
function generateKeepsakePDF() {
  var ss = getSpreadsheet_();

  // 1. Baca data MEMORIES
  var memSheet = ss.getSheetByName('MEMORIES');
  var memories = [];
  if (memSheet && memSheet.getLastRow() > 1) {
    var memHeaders = memSheet.getRange(1, 1, 1, memSheet.getLastColumn()).getDisplayValues()[0];
    var memRows = memSheet.getRange(2, 1, memSheet.getLastRow() - 1, memSheet.getLastColumn()).getDisplayValues();
    for (var m = 0; m < memRows.length; m++) {
      var mObj = {};
      for (var mh = 0; mh < memHeaders.length; mh++) {
        mObj[memHeaders[mh]] = memRows[m][mh] || '';
      }
      memories.push(mObj);
    }
  }

  // 2. Baca data CANDLE_WISHES
  var wishSheet = ss.getSheetByName('CANDLE_WISHES');
  var candleWishes = [];
  if (wishSheet && wishSheet.getLastRow() > 1) {
    var wishHeaders = wishSheet.getRange(1, 1, 1, wishSheet.getLastColumn()).getDisplayValues()[0];
    var wishRows = wishSheet.getRange(2, 1, wishSheet.getLastRow() - 1, wishSheet.getLastColumn()).getDisplayValues();
    for (var w = 0; w < wishRows.length; w++) {
      var wObj = {};
      for (var wh = 0; wh < wishHeaders.length; wh++) {
        wObj[wishHeaders[wh]] = wishRows[w][wh] || '';
      }
      candleWishes.push(wObj);
    }
  }

  var timestampStr = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone() || 'Asia/Jakarta',
    'dd MMMM yyyy, HH:mm'
  );

  // 3. Susun HTML Dokumen dengan String Concatenation murni (Zero template literals)
  var html = '<!DOCTYPE html>' +
    '<html>' +
    '<head>' +
    '<meta charset="UTF-8">' +
    '<style>' +
    '  * { box-sizing: border-box; margin: 0; padding: 0; }' +
    '  body { font-family: "Georgia", serif; background-color: #2D0000; color: #FAEBD7; padding: 40px; }' +
    '  .header { text-align: center; border-bottom: 2px solid #A30000; padding-bottom: 25px; margin-bottom: 30px; }' +
    '  .title { font-size: 28px; color: #FFD700; letter-spacing: 2px; margin-bottom: 8px; }' +
    '  .subtitle { font-size: 16px; color: #FFB6C1; font-style: italic; }' +
    '  .meta { font-size: 11px; color: #FAEBD7; opacity: 0.7; margin-top: 10px; }' +
    '  .section-title { font-size: 20px; color: #FFD700; margin: 30px 0 15px 0; border-left: 4px solid #A30000; padding-left: 10px; }' +
    '  .card { background-color: #4A0000; border: 1px solid rgba(250,235,215,0.2); border-radius: 8px; padding: 18px; margin-bottom: 16px; }' +
    '  .card-title { font-size: 16px; font-weight: bold; color: #FAEBD7; margin-bottom: 5px; }' +
    '  .card-date { font-size: 12px; color: #FFD700; margin-bottom: 8px; }' +
    '  .card-desc { font-size: 13px; line-height: 1.6; color: #FFF; opacity: 0.9; }' +
    '  .tag { display: inline-block; background-color: #A30000; color: #FFF; font-size: 10px; padding: 3px 8px; border-radius: 10px; margin-top: 8px; }' +
    '  .footer { text-align: center; margin-top: 40px; padding-top: 20px; border-top: 1px dashed rgba(250,235,215,0.3); font-size: 12px; color: #FFB6C1; }' +
    '</style>' +
    '</head>' +
    '<body>' +
    '  <div class="header">' +
    '    <h1 class="title">DIGITAL KEEPSAKE CAPSULE</h1>' +
    '    <p class="subtitle">A Special Birthday Keepsake for My Future Partner, Tiara</p>' +
    '    <p class="meta">Terdokumentasi pada: ' + escapeHtml_(timestampStr) + '</p>' +
    '  </div>';

  // Section Memories
  html += '<div class="section-title">Dokumentasi Momen &amp; Cerita Berharga</div>';
  if (memories.length === 0) {
    html += '<p style="font-style: italic; opacity: 0.7;">Belum ada momen yang tersimpan.</p>';
  } else {
    for (var i = 0; i < memories.length; i++) {
      var item = memories[i];
      html += '<div class="card">' +
        '  <div class="card-title">' + escapeHtml_(item.Judul_Momen || 'Momen Spesial') + '</div>' +
        '  <div class="card-date">\uD83D\uDCC5 ' + escapeHtml_(item.Tanggal_Momen || '-') + '</div>' +
        '  <div class="card-desc">' + escapeHtml_(item.Deskripsi || '') + '</div>';
      if (item.Kategori_Milestone) {
        html += '  <span class="tag">' + escapeHtml_(item.Kategori_Milestone) + '</span>';
      }
      html += '</div>';
    }
  }

  // Section Candle Wishes
  html += '<div class="section-title">Koleksi Harapan Rahasia (Make-A-Wish)</div>';
  if (candleWishes.length === 0) {
    html += '<p style="font-style: italic; opacity: 0.7;">Lilin ulang tahun siap ditiup untuk membuat permohonan pertama.</p>';
  } else {
    for (var w = 0; w < candleWishes.length; w++) {
      var cWish = candleWishes[w];
      html += '<div class="card">' +
        '  <div class="card-title">\u2728 Permohonan #' + escapeHtml_(cWish.ID || (w + 1)) + '</div>' +
        '  <div class="card-date">\u23F0 ' + escapeHtml_(cWish.Timestamp || '-') + '</div>' +
        '  <div class="card-desc">&ldquo;' + escapeHtml_(cWish.Harapan_Rahasia || '') + '&rdquo;</div>' +
        '</div>';
    }
  }

  // Footer
  html += '  <div class="footer">' +
    '    <p>Dibuat dengan sepenuh cinta untuk merayakan hari ulang tahun Tiara \uD83D\uDC96</p>' +
    '    <p>Semoga setiap doa dan impian indahmu terkabul.</p>' +
    '  </div>' +
    '</body>' +
    '</html>';

  // 4. Generate Blob dan Simpan ke Google Drive
  var htmlBlob = Utilities.newBlob(html, 'text/html', 'Keepsake_Birthday_Tiara.html');
  var pdfBlob = htmlBlob.getAs('application/pdf');
  pdfBlob.setName('Digital_Keepsake_Tiara_Birthday.pdf');

  var file = DriveApp.createFile(pdfBlob);
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (errSharing) {
    Logger.log('Sharing notice: ' + errSharing.toString());
  }

  var downloadUrl = file.getUrl();
  Logger.log('PDF berhasil digenerate: ' + downloadUrl);
  return downloadUrl;
}

/**
 * Utility escape HTML backend
 * @private
 */
function escapeHtml_(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}
