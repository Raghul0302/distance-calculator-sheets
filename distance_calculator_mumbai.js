/**
 * Distance Calculator - Mumbai Sheet
 * Uses Google Maps built-in service (no API key, no billing needed)
 */

const CONFIG = {
  COL_STORE_LAT_LONG : 17,  // Column Q
  COL_STARTING_POINT : 18,  // Column R
  COL_DISTANCE       : 19,  // Column S
  DATA_START_ROW     : 2,   // Row 1 is header, data starts at row 2
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("🚗 Distance Tools")
    .addItem("Fill Missing Distances",    "fillMissingDistances")
    .addItem("Recalculate All Distances", "recalculateAllDistances")
    .addToUi();
}

function fillMissingDistances()    { calculateDistances(false); }
function recalculateAllDistances() {
  const ui = SpreadsheetApp.getUi();
  if (ui.alert("Recalculate All?", "This will overwrite ALL existing distances. Continue?", ui.ButtonSet.YES_NO) === ui.Button.YES)
    calculateDistances(true);
}

function calculateDistances(recalculateAll) {
  const sheet   = SpreadsheetApp.getActiveSheet();
  const lastRow = sheet.getLastRow();
  const ui      = SpreadsheetApp.getUi();

  if (lastRow < CONFIG.DATA_START_ROW) { ui.alert("No data rows found."); return; }

  const data = sheet.getRange(
    CONFIG.DATA_START_ROW, 1,
    lastRow - CONFIG.DATA_START_ROW + 1,
    Math.max(CONFIG.COL_DISTANCE, CONFIG.COL_STORE_LAT_LONG, CONFIG.COL_STARTING_POINT)
  ).getValues();

  let processed = 0, skipped = 0, errors = 0;

  for (let i = 0; i < data.length; i++) {
    const row       = data[i];
    const sheetRow  = i + CONFIG.DATA_START_ROW;
    const storeCell = String(row[CONFIG.COL_STORE_LAT_LONG - 1] || "").trim();
    const startCell = String(row[CONFIG.COL_STARTING_POINT - 1] || "").trim();
    const distCell  = String(row[CONFIG.COL_DISTANCE - 1]       || "").trim();

    if (!storeCell) continue;
    if (distCell && !recalculateAll) { skipped++; continue; }

    const store = parseLatLng(storeCell);
    const start = parseLatLng(startCell);

    if (!store || !start) {
      Logger.log(`Row ${sheetRow}: invalid coords — store="${storeCell}" start="${startCell}"`);
      errors++;
      continue;
    }

    const distance = getGoogleMapsDistance(start.lat, start.lng, store.lat, store.lng);

    if (distance !== null) {
      sheet.getRange(sheetRow, CONFIG.COL_DISTANCE).setValue(distance);
      processed++;
    } else {
      errors++;
    }

    Utilities.sleep(500);
  }

  ui.alert("✅ Done!", `Filled: ${processed}\nSkipped: ${skipped}\nErrors: ${errors}`, ui.ButtonSet.OK);
}

function getGoogleMapsDistance(lat1, lng1, lat2, lng2) {
  try {
    const directions = Maps.newDirectionFinder()
      .setOrigin(lat1, lng1)
      .setDestination(lat2, lng2)
      .setMode(Maps.DirectionFinder.Mode.DRIVING)
      .getDirections();

    if (directions.status === "OK") {
      const meters = directions.routes[0].legs[0].distance.value;
      return Math.round(meters / 1000 * 1000) / 1000;
    }
    Logger.log(`Maps error: status=${directions.status}`);
  } catch (e) {
    Logger.log(`Maps exception: ${e.message}`);
  }
  return null;
}

function parseLatLng(str) {
  const nums = str.match(/[-+]?\d+\.?\d*/g);
  if (nums && nums.length >= 2) {
    const lat = parseFloat(nums[0]);
    const lng = parseFloat(nums[1]);
    if (!isNaN(lat) && !isNaN(lng)) return { lat, lng };
  }
  return null;
}
