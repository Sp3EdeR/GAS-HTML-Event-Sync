var sheetName = "geocoding";
var geocodingHeaders = ["locname", "lat", "lng", "dtlast", "dtrefresh"];
var geocodingBounds = {
  low: {latitude: 45.256432120351754, longitude: 15.182080358666552},
  high: {latitude: 48.56561148385152, longitude: 23.340697176398876}
};

/**
 * Updates cached coordinates and latest event end times for changed events.
 *
 * @param {Array<Calendar.Event>} events New or modified Google Calendar events.
 */
function updateGeocoding(events) {
  Logger.log("Updating the geocoding data: events=%s", events.length);

  // Get location and event end time for future events
  events = events || [];
  var latestByLocation = {};
  var now = Date.now();
  events.forEach(function(event) {
    // Calculate the last use of the location; it can be deleted afterwards
    var end = getGeocodingEventEnd_(event);
    if (end == null || end < now)
      return;
    var location = (event.location || "").trim();
    if (location)
      latestByLocation[location] = Math.max(latestByLocation[location] || 0, end);
  });

  var locations = Object.keys(latestByLocation);
  var sheet = getGeocodingSheet_();
  var oldRowCount = sheet.getLastRow();

  // Load existing sheet data
  var rows = sheet.getRange(1, 1, oldRowCount, geocodingHeaders.length).getValues();

  // Create a location index for existing rows
  var rowByLocation = {};
  rows.slice(1).forEach(function(row, index) {
    if (row[0] && rowByLocation[row[0]] == null)
      rowByLocation[row[0]] = index + 1;
  });

  // Get new locations that aren't in the sheet yet, and create empty rows for them
  var newLocations = locations.filter(function(location) {
    return rowByLocation[location] == null;
  });
  newLocations.forEach(function(location) {
    rowByLocation[location] = rows.length;
    rows.push([location, "", "", 0, 0]);
  });

  // Update the last use of each location
  locations.forEach(function(location) {
    var row = rows[rowByLocation[location]];
    row[3] = Math.max(Number(row[3]) || 0, latestByLocation[location]);
  });

  // Delete rows that are no longer used
  var deleted = 0;
  rows = rows.filter(function(row, index) {
    if (index && row[3] !== "" && isFinite(Number(row[3])) && Number(row[3]) < now) {
      deleted++;
      return false;
    }
    return true;
  });
  // Rebuild the location index after deletions
  rowByLocation = {};
  rows.slice(1).forEach(function(row, index) {
    rowByLocation[row[0]] = index + 1;
  });

  // Select locations to refresh
  // Google permits caching Places API data for 30 days
  var refreshBefore = now - 29 * 24 * 60 * 60 * 1000;
  var refreshLocations = rows.slice(1).filter(function(row, index) {
    return row[0] && row[4] !== "" && Number(row[4]) < refreshBefore;
  }).map(function(row) { return row[0]; });

  // Geocode the selected locations
  var coordinates = geocodeLocations_(refreshLocations);
  coordinates.forEach(function(coordinate, index) {
    var rowIndex = rowByLocation[refreshLocations[index]];
    rows[rowIndex][1] = coordinate ? coordinate.latitude : "";
    rows[rowIndex][2] = coordinate ? coordinate.longitude : "";
    rows[rowIndex][4] = coordinate ? now : "";
  });

  if (locations.length || coordinates.length || deleted) {
    sheet.getRange(1, 1, rows.length, geocodingHeaders.length).setValues(rows);
    if (oldRowCount > rows.length)
      sheet.deleteRows(rows.length + 1, oldRowCount - rows.length);
  }
  var resolved = coordinates.filter(Boolean).length;
  Logger.log("Geocoding: events=%s, locations=%s, existing=%s, new=%s, deleted=%s, refreshed=%s, resolved=%s, unresolved=%s",
    events.length, locations.length, locations.length - newLocations.length, newLocations.length,
    deleted, refreshLocations.length, resolved, coordinates.filter(function(value) { return value === null; }).length);
}

/**
 * Opens or creates the spreadsheet and geocoding sheet.
 *
 * @return {GoogleAppsScript.Spreadsheet.Sheet} The initialized geocoding sheet.
 */
function getGeocodingSheet_() {
  var properties = PropertiesService.getScriptProperties();
  var spreadsheet;
  var spreadsheetId = properties.getProperty("GEOCODING_SPREADSHEET_ID");
  try {
    if (spreadsheetId)
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (error) {
    properties.deleteProperty("GEOCODING_SPREADSHEET_ID");
  }

  if (!spreadsheet) {
    var files = DriveApp.getFilesByName(sheetName);
    while (files.hasNext() && !spreadsheet) {
      var file = files.next();
      if (file.getMimeType() == MimeType.GOOGLE_SHEETS)
        spreadsheet = SpreadsheetApp.openById(file.getId());
    }
    spreadsheet = spreadsheet || SpreadsheetApp.create(sheetName);
    properties.setProperty("GEOCODING_SPREADSHEET_ID", spreadsheet.getId());
  }

  var sheet = spreadsheet.getSheetByName("geocoding");
  if (!sheet)
    sheet = spreadsheet.getSheets().length == 1 && spreadsheet.getSheets()[0].getLastRow() == 0
      ? spreadsheet.getSheets()[0].setName("geocoding") : spreadsheet.insertSheet("geocoding");
  if (sheet.getLastRow() == 0) {
    sheet.getRange(1, 1, 1, geocodingHeaders.length).setValues([geocodingHeaders]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

/**
 * Resolves locations with the built-in Apps Script Maps service.
 *
 * @param {Array<string>} locations Location strings to resolve.
 * @return {Array<?Object>} Coordinates aligned with the input locations.
 */
function geocodeLocations_(locations) {
  var geocoder = Maps.newGeocoder()
    .setBounds(geocodingBounds.low.latitude, geocodingBounds.low.longitude,
      geocodingBounds.high.latitude, geocodingBounds.high.longitude)
    .setLanguage("hu")
    .setRegion("hu");
  return locations.map(function(location) {
    try {
      var response = geocoder.geocode(location);
      if (response.status == "ZERO_RESULTS")
        return null;
      if (response.status != "OK" || !response.results.length) {
        Logger.log("Geocoding failed for '%s': %s", location, response.status);
        return undefined;
      }
      var coordinate = response.results[0].geometry.location;
      return {latitude: coordinate.lat, longitude: coordinate.lng};
    } catch (error) {
      Logger.log("Geocoding failed for '%s': %s", location, error.message || error);
      return undefined;
    }
  });
}

/**
 * Converts a Calendar event's end value to UTC epoch milliseconds.
 *
 * @param {Calendar.Event} event A Google Calendar event.
 * @return {?number} UTC milliseconds, or null for a missing/invalid end.
 */
function getGeocodingEventEnd_(event) {
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