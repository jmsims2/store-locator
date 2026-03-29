/**
 * Store Locator - Auto Geocode Script
 *
 * HOW TO INSTALL:
 * 1. Open your Google Sheet
 * 2. Go to Extensions → Apps Script
 * 3. Delete any code in the editor
 * 4. Paste this entire file
 * 5. Click Save (disk icon)
 * 6. Close the Apps Script tab
 * 7. Reload your Google Sheet
 * 8. A new "Geocode" menu will appear in the menu bar
 *
 * HOW TO USE:
 * - Add a new store row with at least a city and state
 * - Click Geocode → Fill Missing Coordinates
 * - The lat and lng columns will be filled in automatically
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Geocode')
    .addItem('Fill Missing Coordinates', 'geocodeMissing')
    .addToUi();
}

function geocodeMissing() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  var headers = data[0].map(function(h) { return h.toString().toLowerCase().trim(); });

  var nameCol = headers.indexOf('name');
  var addrCol = headers.indexOf('address');
  var cityCol = headers.indexOf('city');
  var stateCol = headers.indexOf('state');
  var zipCol = headers.indexOf('zip');
  var latCol = headers.indexOf('lat');
  var lngCol = headers.indexOf('lng');

  if (latCol === -1 || lngCol === -1) {
    SpreadsheetApp.getUi().alert('Could not find "lat" and "lng" columns. Make sure your header row has these columns.');
    return;
  }

  var filled = 0;
  var failed = 0;
  var skipped = 0;

  for (var i = 1; i < data.length; i++) {
    var row = data[i];

    // Skip rows that already have coordinates
    if (row[latCol] && row[lngCol]) {
      continue;
    }

    // Skip rows with no city/state
    var city = cityCol >= 0 ? row[cityCol].toString().trim() : '';
    var state = stateCol >= 0 ? row[stateCol].toString().trim() : '';
    if (!city && !state) {
      skipped++;
      continue;
    }

    // Build address string from available fields
    var parts = [];
    if (addrCol >= 0 && row[addrCol]) parts.push(row[addrCol].toString().trim());
    if (city) parts.push(city);
    if (state) parts.push(state);
    if (zipCol >= 0 && row[zipCol]) parts.push(row[zipCol].toString().trim());
    var address = parts.join(', ');

    try {
      var geo = Maps.newGeocoder().geocode(address);
      if (geo.status === 'OK' && geo.results.length > 0) {
        var location = geo.results[0].geometry.location;
        sheet.getRange(i + 1, latCol + 1).setValue(location.lat);
        sheet.getRange(i + 1, lngCol + 1).setValue(location.lng);
        filled++;
      } else {
        failed++;
        Logger.log('Geocode failed for row ' + (i + 1) + ': ' + address + ' (' + geo.status + ')');
      }
    } catch (e) {
      failed++;
      Logger.log('Error geocoding row ' + (i + 1) + ': ' + e.message);
    }

    // Pause briefly to avoid rate limits
    Utilities.sleep(200);
  }

  var msg = filled + ' store(s) geocoded successfully.';
  if (failed > 0) msg += '\n' + failed + ' store(s) could not be found — check the address.';
  if (skipped > 0) msg += '\n' + skipped + ' row(s) skipped (no city/state).';
  SpreadsheetApp.getUi().alert(msg);
}
