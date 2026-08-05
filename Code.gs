function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var data = JSON.parse(e.postData.contents);
    var timestamp = new Date();

    if (!data.members || !data.members.length) {
      throw new Error('No attendees submitted');
    }

    // --- Registrations tab: one row per attendee ---
    var regSheet = ss.getSheetByName('Registrations');
    if (!regSheet) {
      regSheet = ss.insertSheet('Registrations');
    }
    if (regSheet.getLastRow() === 0) {
      regSheet.appendRow([
        'Timestamp', 'Contact Name', 'Phone', 'Email',
        'Attendee Name', 'Category', 'Attendance', 'Fee (INR)',
        'Group Total (INR)'
      ]);
      regSheet.setFrozenRows(1);
    }

    data.members.forEach(function (m) {
      regSheet.appendRow([
        timestamp,
        data.contactName || '',
        data.phone || '',
        data.email || '',
        m.name || '',
        m.category || '',
        m.attendance || '',
        m.fee || 0,
        data.total || 0
      ]);
    });

    // --- Payments tab: one row per registration ---
    var paySheet = ss.getSheetByName('Payments');
    var isNewPaySheet = false;
    if (!paySheet) {
      paySheet = ss.insertSheet('Payments');
      isNewPaySheet = true;
    }
    if (paySheet.getLastRow() === 0) {
      paySheet.appendRow([
        'Timestamp', 'Contact Name', 'Phone', 'Email',
        'Amount Due (INR)', 'Paid or Not Paid'
      ]);
      paySheet.setFrozenRows(1);
      isNewPaySheet = true;
    }

    paySheet.appendRow([
      timestamp,
      data.contactName || '',
      data.phone || '',
      data.email || '',
      data.total || 0,
      'Not Paid'
    ]);

    if (isNewPaySheet) {
      setupPaymentsFormatting(paySheet);
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', total: data.total }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

/**
 * Adds a Paid / Not Paid dropdown + color coding to the "Paid or Not Paid"
 * column (F) of the Payments sheet, for a generous range of future rows.
 * Runs automatically the first time the Payments tab is created.
 *
 * Apps Script has no API to set the dropdown's display style
 * (Chip / Arrow / Plain text) — that has to be set manually in Sheets:
 * select the column > Data > Data validation > click the rule > Display style > Plain text.
 */
function setupPaymentsFormatting(sheet) {
  var statusCol = 6; // column F
  var numRows = 500;

  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Paid', 'Not Paid'], true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, statusCol, numRows, 1).setDataValidation(rule);

  var range = sheet.getRange(2, statusCol, numRows, 1);
  var rules = sheet.getConditionalFormatRules();
  rules.push(
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Paid')
      .setBackground('#d9ead3')
      .setFontColor('#274e13')
      .setRanges([range])
      .build(),
    SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo('Not Paid')
      .setBackground('#f4cccc')
      .setFontColor('#990000')
      .setRanges([range])
      .build()
  );
  sheet.setConditionalFormatRules(rules);

  sheet.autoResizeColumns(1, 6);
}

function doGet(e) {
  return ContentService.createTextOutput(
    'This is a POST-only endpoint for the ELIM Church Retreat registration form.'
  );
}
