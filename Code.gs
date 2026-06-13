const DISPATCH_SHEET      = 'Dispatches';
const DESTINATIONS_SHEET  = 'Destinations';

// ── GET ──
function doGet(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (e.parameter.action === 'getDestinations') {
    var sheet = ss.getSheetByName(DESTINATIONS_SHEET);
    if (!sheet || sheet.getLastRow() === 0) return json({ destinations: [] });
    var dests = sheet.getRange(1, 1, sheet.getLastRow(), 1)
      .getValues().flat().filter(function(d) { return d; });
    return json({ destinations: dests });
  }

  if (e.parameter.action === 'getDispatches') {
    var sheet = ss.getSheetByName(DISPATCH_SHEET);
    if (!sheet || sheet.getLastRow() <= 1) return json({ dispatches: [] });
    var tz = Session.getScriptTimeZone();

    function fmtDate(v) {
      if (!v) return '';
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
      return String(v);
    }
    function fmtTime(v) {
      if (!v) return '';
      if (v instanceof Date) return Utilities.formatDate(v, tz, 'HH:mm');
      return String(v);
    }

    // Row 1 is the header row — start from row 2
    var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();
    var records = rows.filter(function(r) { return r[0]; }).map(function(r) {
      return {
        date:    fmtDate(r[0]),
        time:    fmtTime(r[1]),
        dest:    String(r[2]),
        staff:   String(r[3]),
        product: String(r[4]),
        batch:   String(r[5]),
        expiry:  fmtDate(r[6]),
        qty:     String(r[7]),
        unit:    String(r[8]),
        savedAt: String(r[9] || '')
      };
    });
    return json({ dispatches: records });
  }

  return json({ status: 'ok' });
}

// ── POST ──
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // Save dispatch records
  if (data.action === 'saveDispatch') {
    var sheet = ss.getSheetByName(DISPATCH_SHEET);
    if (!sheet) {
      sheet = ss.insertSheet(DISPATCH_SHEET);
      var headers = ['Date','Time','Destination','Dispatched By','Product','Batch','Expiry','Qty','Unit','Saved At'];
      sheet.appendRow(headers);
      var headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setFontWeight('bold')
                 .setBackground('#1a1916')
                 .setFontColor('#ffffff');
      sheet.setFrozenRows(1);
      sheet.setColumnWidth(1, 100);
      sheet.setColumnWidth(2, 70);
      sheet.setColumnWidth(3, 130);
      sheet.setColumnWidth(4, 130);
      sheet.setColumnWidth(5, 160);
      sheet.setColumnWidth(6, 120);
      sheet.setColumnWidth(7, 100);
      sheet.setColumnWidth(8, 60);
      sheet.setColumnWidth(9, 70);
      sheet.setColumnWidth(10, 160);
    }
    data.rows.forEach(function(r) {
      sheet.appendRow([
        r.date, r.time, r.dest, r.staff,
        r.product, r.batch, r.expiry,
        r.qty, r.unit, r.savedAt
      ]);
    });
    return json({ result: 'success' });
  }

  // Save destinations list
  if (data.action === 'saveDestinations') {
    var sheet = ss.getSheetByName(DESTINATIONS_SHEET);
    if (!sheet) sheet = ss.insertSheet(DESTINATIONS_SHEET);
    sheet.clearContents();
    (data.destinations || []).forEach(function(d) { sheet.appendRow([d]); });
    return json({ result: 'success' });
  }

  return json({ result: 'ok' });
}

// ── Helper ──
function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
