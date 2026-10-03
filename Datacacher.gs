var dataFileName = "calendar_data";
var dataHeaders = ["start", "end", "summary", "location", "description"];

/**
 * Caches upcoming events using the calendar snapshots maintained by the sync.
 *
 * @param {Object<string, Array<Calendar.Event>>} eventsByCalendar Events keyed by calendar ID.
 */
function updateDataCache(eventsByCalendar) {
  var calendarIds = Object.keys(eventsByCalendar || {});
  if (!calendarIds.length)
    return;

  var eventsSheet = getMetadataSheet(dataFileName, "events", dataHeaders.concat("calid"));
  var spreadsheet = eventsSheet.getParent();
  var sheetsByName = Object.create(null);
  spreadsheet.getSheets().forEach(function(sheet) {
    sheetsByName[sheet.getName()] = sheet;
  });

  // Ensure the calendars list sheet exists
  var calendarsSheet = sheetsByName.calendars;
  if (!calendarsSheet) {
    calendarsSheet = spreadsheet.insertSheet("calendars", 1);
    calendarsSheet.getRange(1, 1).setValue("calid");
    calendarsSheet.setFrozenRows(1);
  } else if (calendarsSheet.getLastRow() == 0) {
    calendarsSheet.getRange(1, 1).setValue("calid");
    calendarsSheet.setFrozenRows(1);
  }

  // Get the list of already stored calendar IDs
  var lastRow = calendarsSheet.getLastRow();
  var registered = Object.create(null);
  if (lastRow) {
    calendarsSheet.getRange(1, 1, lastRow, 1).getValues().forEach(function(row) {
      if (row[0] && row[0] != "calid")
        registered[row[0]] = true;
    });
  }

  var now = Date.now();
  var newIds = [];
  calendarIds.forEach(function(calendarId) {
    var sheet = sheetsByName[calendarId];
    // Ensure the sheet for this calendar exists
    if (!sheet) {
      sheet = spreadsheet.insertSheet(calendarId);
      sheet.setFrozenRows(1);
    }

    // Write events to the sheet for this calendar
    writeDataCacheRows(sheet, [dataHeaders].concat(
      getDataCacheRows(eventsByCalendar[calendarId], now)));

    // Add the calendar ID to the list of new IDs if it hasn't been registered yet
    if (!registered[calendarId]) {
      newIds.push([calendarId]);
      registered[calendarId] = true;
    }
  });

  // Store the new calendar IDs in the list of registered IDs
  if (newIds.length) {
    resizeDataCacheSheet(calendarsSheet, lastRow + newIds.length, 1);
    calendarsSheet.getRange(lastRow + 1, 1, newIds.length, 1).setValues(newIds);
  }

  // Reserve spill space for up to 500 events per registered calendar.
  resizeDataCacheSheet(eventsSheet, Object.keys(registered).length * 500 + 1, 5);

  // Initialize the formula for the events sheet.
  // It dynamically merges data from all calendar event data sheets.
  var formula = `=IFERROR(
  LET(
    ids, FILTER(calendars!A:A, calendars!A:A<>"", calendars!A:A<>"calid"),
    rows, REDUCE(
      {"","","","","",""},
      ids,
      LAMBDA(
        merged, calendar,
        VSTACK(
          merged,
          IFERROR(
            LET(
              data, FILTER(
                INDIRECT("'"&calendar&"'!A2:E"),
                INDIRECT("'"&calendar&"'!A2:A")<>""
              ),
              HSTACK(data, BYROW(data, LAMBDA(r, calendar)))
            ),
            {"","","","","",""}
          )
        )
      )
    ),
    FILTER(rows, INDEX(rows,0,1)<>"")
  ),
  ""
)`;
  var formulaCell = eventsSheet.getRange(2, 1);
  if (formulaCell.getFormula() != formula)
    formulaCell.setFormula(formula);

  // Reorder the sheets; events first, calendars second, rest last.
  if (eventsSheet.getIndex() != 1 || calendarsSheet.getIndex() != 2) {
    var activeSheet = spreadsheet.getActiveSheet();
    spreadsheet.setActiveSheet(eventsSheet);
    spreadsheet.moveActiveSheet(1);
    spreadsheet.setActiveSheet(calendarsSheet);
    spreadsheet.moveActiveSheet(2);
    spreadsheet.setActiveSheet(activeSheet);
  }
  Logger.log("Updated data cache for %s calendars", calendarIds.length);
}

/**
 * Builds export rows with ISO UTC dates and unchanged event text, including HTML.
 *
 * @param {Array<Calendar.Event>} events Calendar events already loaded by the sync.
 * @param {number} now Current UTC epoch milliseconds.
 * @return {Array<Array<string>>} Upcoming event rows.
 */
function getDataCacheRows(events, now) {
  var rows = [];
  (events || []).forEach(function(event) {
    if (!event || event.status == "cancelled")
      return;
    var end = getEventEnd(event);
    if (end == null || end <= now)
      return;
    var start = event.start || {};
    var startTime = new Date(start.dateTime || start.date).getTime();
    if (isNaN(startTime))
      return;
    rows.push([new Date(startTime).toISOString(), new Date(end).toISOString(),
      event.summary == null ? "" : String(event.summary),
      event.location == null ? "" : String(event.location),
      event.description == null ? "" : String(event.description)]);
  });
  return rows;
}

/**
 * Replaces a calendar's cached rows and clears any stale trailing events.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Destination calendar sheet.
 * @param {Array<Array<string>>} rows Header and event rows.
 */
function writeDataCacheRows(sheet, rows) {
  var lastRow = sheet.getLastRow();
  resizeDataCacheSheet(sheet, Math.max(2, rows.length), dataHeaders.length);
  // Rich text writes literal strings, avoiding formula injection and date/number coercion.
  var values = rows.map(function(row) {
    return row.map(function(value) {
      return SpreadsheetApp.newRichTextValue().setText(value).build();
    });
  });
  sheet.getRange(1, 1, rows.length, dataHeaders.length).setRichTextValues(values);
  if (lastRow > rows.length)
    sheet.getRange(rows.length + 1, 1, lastRow - rows.length, 5).clearContent();
}

/**
 * Extends a sheet only when its existing grid cannot hold the required range.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet Sheet to resize.
 * @param {number} rows Required row count.
 * @param {number} columns Required column count.
 */
function resizeDataCacheSheet(sheet, rows, columns) {
  var currentRows = sheet.getMaxRows();
  var currentColumns = sheet.getMaxColumns();
  if (rows > currentRows)
    sheet.insertRowsAfter(currentRows, rows - currentRows);
  if (columns > currentColumns)
    sheet.insertColumnsAfter(currentColumns, columns - currentColumns);
}
