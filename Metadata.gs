var metadataFileName = "calendar_meta";
var metadataHeaders = ["calid", "count"];

/**
 * Updates future event counts for calendars processed by the current sync.
 *
 * @param {Object<string, Array<Calendar.Event>>} eventsByCalendar Events keyed by calendar ID.
 */
function updateMetadata(eventsByCalendar) {
  var calendarIds = Object.keys(eventsByCalendar || {});
  if (!calendarIds.length)
    return;

  // Only count events that are in the future
  var now = Date.now();
  var countsByCalendar = {};
  calendarIds.forEach(function(calendarId) {
    countsByCalendar[calendarId] = eventsByCalendar[calendarId].filter(function(event) {
      var end = getEventEnd(event);
      return end != null && end >= now;
    }).length;
  });

  var sheet = getMetadataSheet(metadataFileName, "metadata", metadataHeaders);
  var lastRow = sheet.getLastRow();

  // Load existing sheet data
  var rows = lastRow > 1
    ? sheet.getRange(2, 1, lastRow - 1, metadataHeaders.length).getValues()
    : [];

  // Create a calendar index for existing rows
  var rowByCalendar = {};
  rows.forEach(function(row, index) {
    if (row[0] && rowByCalendar[row[0]] == null)
      rowByCalendar[row[0]] = index + 2;
  });

  // Add or update metadata for each calendar processed
  calendarIds.forEach(function(calendarId) {
    var rowNumber = rowByCalendar[calendarId];
    if (rowNumber == null) {
      sheet.appendRow([calendarId, countsByCalendar[calendarId]]);
    } else {
      sheet.getRange(rowNumber, 2).setValue(countsByCalendar[calendarId]);
    }
  });

  Logger.log("Updated metadata for %s calendars", calendarIds.length);
}

/**
 * Records an event added to or updated in a calendar during synchronization.
 *
 * @param {string} calendarId Google Calendar ID.
 * @param {Calendar.Event} event Added or updated event.
 * @param {string=} previousEventId ID of the event replaced by an update.
 */
function recordMetadataEvent(calendarId, event, previousEventId) {
  if (!event || !metadataEventsByCalendar[calendarId])
    return;

  var events = metadataEventsByCalendar[calendarId];
  var index = previousEventId == null ? -1 : events.findIndex(function(existingEvent) {
    return existingEvent.id == previousEventId;
  });
  if (index < 0)
    events.push(event);
  else
    events[index] = event;
}

/**
 * Records an event removed from a calendar during synchronization.
 *
 * @param {string} calendarId Google Calendar ID.
 * @param {string} eventId Removed event ID.
 */
function removeMetadataEvent(calendarId, eventId) {
  if (!metadataEventsByCalendar[calendarId])
    return;

  metadataEventsByCalendar[calendarId] = metadataEventsByCalendar[calendarId].filter(function(event) {
    return event.id != eventId;
  });
}

/**
 * Opens or creates a spreadsheet and sheet.
 *
 * @param {string} fileName Spreadsheet file name.
 * @param {string} sheetName Sheet name.
 * @param {Array<string>} headers Sheet column names.
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The initialized sheet.
 */
function getMetadataSheet(fileName, sheetName, headers) {
  var properties = PropertiesService.getScriptProperties();
  var spreadsheet;
  var propertyName = fileName.toUpperCase().replace(/[^A-Z0-9]+/g, "_") + "_SPREADSHEET_ID";
  var spreadsheetId = properties.getProperty(propertyName);
  try {
    if (spreadsheetId)
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    properties.deleteProperty(propertyName);
  }

  if (!spreadsheet) {
    var files = DriveApp.getFilesByName(fileName);
    while (files.hasNext() && !spreadsheet) {
      var file = files.next();
      if (file.getMimeType() == MimeType.GOOGLE_SHEETS)
        spreadsheet = SpreadsheetApp.openById(file.getId());
    }
    spreadsheet = spreadsheet || SpreadsheetApp.create(fileName);
    properties.setProperty(propertyName, spreadsheet.getId());
  }

  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet)
    sheet = spreadsheet.getSheets().length == 1 && spreadsheet.getSheets()[0].getLastRow() == 0
      ? spreadsheet.getSheets()[0].setName(sheetName) : spreadsheet.insertSheet(sheetName);
  if (sheet.getLastRow() == 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Converts a Calendar event's end value to UTC epoch milliseconds.
 *
 * @param {Calendar.Event} event A Google Calendar event.
 * @return {?number} UTC milliseconds, or null for a missing/invalid end.
 */
function getEventEnd(event) {
  var end = event.end || {};
  var value = end.dateTime || end.date;
  if (value) {
    var milliseconds = new Date(value).getTime();
    return isNaN(milliseconds) ? null : milliseconds;
  }

  var start = event.start || {};
  if (start.dateTime) {
    var startMilliseconds = new Date(start.dateTime).getTime();
    return isNaN(startMilliseconds) ? null : startMilliseconds + 3 * 60 * 60 * 1000;
  }
  if (start.date) {
    var startDateMilliseconds = new Date(start.date).getTime();
    return isNaN(startDateMilliseconds) ? null : startDateMilliseconds + 24 * 60 * 60 * 1000;
  }
  return null;
}