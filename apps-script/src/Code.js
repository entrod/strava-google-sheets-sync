const MAX_HR = 000; // Update to your max HR

// HR Zones as % of MAX_HR - these are examples and should be changed
const HR_ZONES = {
  Z1: { min: 0.50, max: 0.60 },
  Z2: { min: 0.60, max: 0.70 },
  Z3: { min: 0.70, max: 0.80 },
  Z4: { min: 0.80, max: 0.90 },
  Z5: { min: 0.90, max: 1.00 }
};

// Privacy: keep false if sharing/exporting your sheet publicly
const INCLUDE_LOCATION_DATA = false;

// Default number of activities to fetch (Strava API max per page is 200) - The more the longer the update takes. Can be changed over time. First have 200 to get many activites, but then change to ex 5.
const PER_PAGE = 50;

// Menu below
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏃 Strava')
    .addItem('📥 Importera aktiviteter', 'importStravaActivities')
    .addSeparator()
    .addItem('⚙️ Setup (första gången)', 'showSetupDialog')
    .addItem('🔄 Testa token', 'testToken')
    .addToUi();
}

// SETUP DIALOG (HTML in separate file)

function showSetupDialog() {
  const html = HtmlService.createTemplateFromFile('SetupDialog')
    .evaluate()
    .setWidth(600)
    .setHeight(650);

  SpreadsheetApp.getUi().showModalDialog(html, 'Strava Setup');
}

/**
 * Called from SetupDialog.html via google.script.run
 * Saves credentials and exchanges auth code for tokens.
 */
function setupStravaFromDialog(clientId, clientSecret, authCode) {
  logToSheet_('Starting Strava setup...', 'INFO');

  try {
    const props = PropertiesService.getScriptProperties();

    props.setProperty('STRAVA_CLIENT_ID', String(clientId).trim());
    props.setProperty('STRAVA_CLIENT_SECRET', String(clientSecret).trim());

    const resp = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
      method: 'post',
      payload: {
        client_id: String(clientId).trim(),
        client_secret: String(clientSecret).trim(),
        code: String(authCode).trim(),
        grant_type: 'authorization_code'
      },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() !== 200) {
      logToSheet_('Setup failed: ' + resp.getContentText(), 'ERROR');
      throw new Error('Kunde inte hämta tokens: ' + resp.getContentText());
    }

    const data = JSON.parse(resp.getContentText());
    props.setProperty('STRAVA_ACCESS_TOKEN', data.access_token);
    props.setProperty('STRAVA_REFRESH_TOKEN', data.refresh_token);
    props.setProperty('STRAVA_EXPIRES_AT', String(data.expires_at));

    logToSheet_('Strava setup completed successfully', 'SUCCESS');
    return 'Setup klar! Du kan nu importera aktiviteter.';

  } catch (error) {
    logToSheet_('Setup failed: ' + error.message, 'ERROR');
    throw error;
  }
}

// TOKEN MANAGEMENT below

function getValidAccessToken_() {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID');
  const clientSecret = props.getProperty('STRAVA_CLIENT_SECRET');
  const refreshToken = props.getProperty('STRAVA_REFRESH_TOKEN');
  const expiresAt = props.getProperty('STRAVA_EXPIRES_AT');
  const accessToken = props.getProperty('STRAVA_ACCESS_TOKEN');

  if (!clientId || !clientSecret || !refreshToken) {
    const msg = 'Missing Strava credentials. Run Setup first (🏃 Strava → Setup).';
    logToSheet_(msg, 'ERROR');
    throw new Error(msg);
  }

  const now = Math.floor(Date.now() / 1000);
  if (accessToken && expiresAt && parseInt(expiresAt, 10) > (now + 300)) {
    return accessToken;
  }

  logToSheet_('Refreshing access token...', 'INFO');

  const resp = UrlFetchApp.fetch('https://www.strava.com/oauth/token', {
    method: 'post',
    payload: {
      client_id: clientId,
      client_secret: clientSecret,
      refresh_token: refreshToken,
      grant_type: 'refresh_token'
    },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    const err = resp.getContentText();
    logToSheet_('Token refresh failed: ' + err, 'ERROR');
    throw new Error('Could not refresh access token: ' + err);
  }

  const data = JSON.parse(resp.getContentText());
  props.setProperty('STRAVA_ACCESS_TOKEN', data.access_token);
  props.setProperty('STRAVA_REFRESH_TOKEN', data.refresh_token);
  props.setProperty('STRAVA_EXPIRES_AT', String(data.expires_at));

  logToSheet_('Token refreshed successfully', 'SUCCESS');
  return data.access_token;
}

// EXECUTION LOGGING

function logToSheet_(message, status) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let logSheet = ss.getSheetByName('Execution Log');

    if (!logSheet) {
      logSheet = ss.insertSheet('Execution Log');
      logSheet.appendRow(['Timestamp', 'Status', 'Message']);
      logSheet.getRange('A1:C1').setFontWeight('bold').setBackground('#f3f3f3');
    }

    const timestamp = new Date();
    const statusEmoji =
      status === 'SUCCESS' ? '✓' :
      status === 'ERROR' ? '✗' :
      status === 'INFO' ? 'ℹ️' : '▶';

    logSheet.appendRow([timestamp, statusEmoji + ' ' + status, message]);

    const lastRow = logSheet.getLastRow();
    const statusCell = logSheet.getRange(lastRow, 2);

    if (status === 'SUCCESS') {
      statusCell.setBackground('#d4edda').setFontColor('#155724');
    } else if (status === 'ERROR') {
      statusCell.setBackground('#f8d7da').setFontColor('#721c24');
    } else if (status === 'INFO') {
      statusCell.setBackground('#d1ecf1').setFontColor('#0c5460');
    } else {
      statusCell.setBackground('#fff3cd').setFontColor('#856404');
    }

    logSheet.autoResizeColumns(1, 3);

    // Keep latest 100 rows (+ header)
    if (logSheet.getLastRow() > 101) {
      logSheet.deleteRows(2, logSheet.getLastRow() - 101);
    }
  } catch (e) {
    Logger.log('Failed to log to sheet: ' + e.message);
  }
}

// HR ZONES CALCULATION

function getActivityStreams_(activityId, accessToken) {
  const url =
    'https://www.strava.com/api/v3/activities/' +
    activityId +
    '/streams?keys=heartrate,time&key_by_type=true';

  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + accessToken },
    muteHttpExceptions: true
  };

  try {
    const resp = UrlFetchApp.fetch(url, options);
    if (resp.getResponseCode() !== 200) return null;
    return JSON.parse(resp.getContentText());
  } catch (e) {
    Logger.log('Error fetching streams for ' + activityId + ': ' + e.message);
    return null;
  }
}

function computeTimeInZones_(streams, maxHR) {
  if (!streams || !streams.heartrate || !streams.time) {
    return { Z1_min: '', Z2_min: '', Z3_min: '', Z4_min: '', Z5_min: '' };
  }

  const hrData = streams.heartrate.data;
  const timeData = streams.time.data;

  if (!hrData || !timeData || hrData.length === 0) {
    return { Z1_min: '', Z2_min: '', Z3_min: '', Z4_min: '', Z5_min: '' };
  }

  const zoneTimes = { Z1: 0, Z2: 0, Z3: 0, Z4: 0, Z5: 0 };

  for (let i = 0; i < hrData.length; i++) {
    const hr = hrData[i];
    if (!hr || hr <= 0) continue;

    let intervalSec = 1;
    if (i > 0) intervalSec = timeData[i] - timeData[i - 1];

    const hrPercent = hr / maxHR;

    if (hrPercent >= HR_ZONES.Z5.min) zoneTimes.Z5 += intervalSec;
    else if (hrPercent >= HR_ZONES.Z4.min) zoneTimes.Z4 += intervalSec;
    else if (hrPercent >= HR_ZONES.Z3.min) zoneTimes.Z3 += intervalSec;
    else if (hrPercent >= HR_ZONES.Z2.min) zoneTimes.Z2 += intervalSec;
    else if (hrPercent >= HR_ZONES.Z1.min) zoneTimes.Z1 += intervalSec;
  }

  return {
    Z1_min: Math.round((zoneTimes.Z1 / 60) * 10) / 10,
    Z2_min: Math.round((zoneTimes.Z2 / 60) * 10) / 10,
    Z3_min: Math.round((zoneTimes.Z3 / 60) * 10) / 10,
    Z4_min: Math.round((zoneTimes.Z4 / 60) * 10) / 10,
    Z5_min: Math.round((zoneTimes.Z5 / 60) * 10) / 10
  };
}

// UTILITY FUNCTIONS

function formatPaceFromSeconds_(secondsPerKm) {
  if (!secondsPerKm || secondsPerKm <= 0) return '';
  const totalSeconds = Math.round(secondsPerKm);
  const mm = Math.floor(totalSeconds / 60);
  const ss = totalSeconds % 60;
  return (mm < 10 ? '0' + mm : mm) + ':' + (ss < 10 ? '0' + ss : ss);
}

function getGearInfo_(gearIds, accessToken) {
  const gearInfo = {};
  if (!gearIds || gearIds.length === 0) return gearInfo;

  const baseUrl = 'https://www.strava.com/api/v3/gear/';
  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + accessToken },
    muteHttpExceptions: true
  };

  gearIds.forEach(id => {
    try {
      const resp = UrlFetchApp.fetch(baseUrl + id, options);
      if (resp.getResponseCode() === 200) {
        const g = JSON.parse(resp.getContentText());
        gearInfo[id] = {
          name: g.name || id,
          distanceKm: (g.distance || 0) / 1000.0
        };
      } else {
        gearInfo[id] = { name: id, distanceKm: '' };
      }
    } catch (e) {
      gearInfo[id] = { name: id, distanceKm: '' };
    }
  });

  return gearInfo;
}

// SPLITS FUNCTIONS

function importStravaSplits_(activityId, splitsSheet, accessToken) {
  const url = 'https://www.strava.com/api/v3/activities/' + activityId;

  const resp = UrlFetchApp.fetch(url, {
    method: 'get',
    headers: { Authorization: 'Bearer ' + accessToken },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) return;

  const detail = JSON.parse(resp.getContentText());
  const splits = detail.splits_metric;

  if (!splits || !splits.length) return;

  splits.forEach((s, index) => {
    const paceSec = s.pace;
    const paceDisplay = formatPaceFromSeconds_(paceSec);

    splitsSheet.appendRow([
      activityId,
      index + 1,
      paceSec,
      paceDisplay,
      s.average_heartrate || '',
      s.max_heartrate || '',
      s.elevation_difference || 0,
      s.moving_time || '',
      s.elapsed_time || ''
    ]);
  });
}

function syncSplitsToStored_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName('Splits');
  if (!srcSheet || srcSheet.getLastRow() < 2) return;

  let destSheet = ss.getSheetByName('SplitsDataStored');
  if (!destSheet) destSheet = ss.insertSheet('SplitsDataStored');

  const srcLastRow = srcSheet.getLastRow();
  const srcLastCol = srcSheet.getLastColumn();
  const srcValues = srcSheet.getRange(2, 1, srcLastRow - 1, srcLastCol).getValues();

  if (destSheet.getLastRow() === 0) {
    const header = srcSheet.getRange(1, 1, 1, srcLastCol).getValues()[0];
    destSheet.getRange(1, 1, 1, srcLastCol).setValues([header]);
  }

  const destLastRow = destSheet.getLastRow();
  const existingKeys = new Set();

  if (destLastRow >= 2) {
    const existingRange = destSheet.getRange(2, 1, destLastRow - 1, 2).getValues();
    existingRange.forEach(row => {
      if (row[0] && row[1]) existingKeys.add(String(row[0]) + '|' + String(row[1]));
    });
  }

  const rowsToAppend = srcValues.filter(row => {
    if (!row[0] || !row[1]) return false;
    const key = String(row[0]) + '|' + String(row[1]);
    return !existingKeys.has(key);
  });

  if (rowsToAppend.length > 0) {
    const destStartRow = destSheet.getLastRow() + 1;
    destSheet.getRange(destStartRow, 1, rowsToAppend.length, rowsToAppend[0].length)
      .setValues(rowsToAppend);
  }
}

// DATA SYNC (StravaData -> Data append only)

function syncStravaToData_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName('StravaData');
  const destSheet = ss.getSheetByName('Data');

  if (!srcSheet) throw new Error('StravaData sheet not found');
  if (!destSheet) throw new Error('Data sheet not found');

  const srcLastRow = srcSheet.getLastRow();
  if (srcLastRow < 2) return;

  const srcValues = srcSheet.getRange(2, 1, srcLastRow - 1, srcSheet.getLastColumn()).getValues();

  const destLastRow = destSheet.getLastRow();
  const existingIds = new Set();

  if (destLastRow >= 2) {
    const idRange = destSheet.getRange(2, 1, destLastRow - 1, 1).getValues();
    idRange.forEach(r => { if (r[0]) existingIds.add(String(r[0])); });
  }

  const newRows = srcValues.filter(row => row[0] && !existingIds.has(String(row[0])));

  if (newRows.length > 0) {
    destSheet.getRange(destLastRow + 1, 1, newRows.length, newRows[0].length)
      .setValues(newRows);
  }
}

// ============================================
// MAIN IMPORT
// ============================================

function importStravaActivities() {
  const startTime = new Date();
  logToSheet_('Starting Strava import...', 'START');

  try {
    const accessToken = getValidAccessToken_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Splits sheet
    let splitsSheet = ss.getSheetByName('Splits');
    if (!splitsSheet) splitsSheet = ss.insertSheet('Splits');

    if (splitsSheet.getLastRow() === 0) {
      splitsSheet.appendRow([
        'activity_id', 'km', 'pace_sec', 'pace_mmss',
        'avg_hr', 'max_hr', 'elev_diff',
        'moving_time_sec', 'elapsed_time_sec'
      ]);
    }

    // Fetch activities
    logToSheet_('Fetching activities from Strava...', 'INFO');

    const url = 'https://www.strava.com/api/v3/athlete/activities?per_page=' + PER_PAGE;
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });

    if (response.getResponseCode() !== 200) {
      throw new Error('Failed to fetch activities: ' + response.getContentText());
    }

    const activities = JSON.parse(response.getContentText());
    logToSheet_('Found ' + activities.length + ' activities', 'INFO');

    // StravaData sheet
    let sheet = ss.getSheetByName('StravaData');
    if (!sheet) sheet = ss.insertSheet('StravaData');

    sheet.clearContents();
    sheet.appendRow([
      'id', 'date', 'type', 'sport_type', 'name', 'city',
      'distance_km', 'moving_time_min', 'moving_time_sec', 'elapsed_time_sec',
      'elevation_gain_m', 'elev_low_m', 'elev_high_m',
      'paceSecPerKm', 'paceDisplay', 'avg_speed_kmh', 'max_speed_kmh',
      'avg_heartrate', 'max_heartrate',
      'start_latlng', 'end_latlng', 'polyline',
      'gear_name', 'gear_id', 'gear_km_after_run',
      'Z1_min', 'Z2_min', 'Z3_min', 'Z4_min', 'Z5_min'
    ]);

    // Gear info
    logToSheet_('Fetching gear information...', 'INFO');
    const gearIds = Array.from(new Set(activities.map(a => a.gear_id).filter(Boolean)));
    const gearInfo = getGearInfo_(gearIds, accessToken);
    const gearConsumedKm = {};

    // Process activities
    logToSheet_('Processing activities (HR zones + splits)...', 'INFO');

    activities.forEach(a => {
      const distanceKm = (a.distance || 0) / 1000.0;
      const movingMinutes = (a.moving_time || 0) / 60.0;
      const paceSecPerKm = distanceKm > 0 ? (a.moving_time / distanceKm) : 0;
      const paceDisplay = formatPaceFromSeconds_(paceSecPerKm);

      const avgSpeedKmh = a.average_speed ? a.average_speed * 3.6 : '';
      const maxSpeedKmh = a.max_speed ? a.max_speed * 3.6 : '';

      const startLatLng = INCLUDE_LOCATION_DATA && a.start_latlng && a.start_latlng.length === 2
        ? a.start_latlng[0] + ',' + a.start_latlng[1]
        : '';

      const endLatLng = INCLUDE_LOCATION_DATA && a.end_latlng && a.end_latlng.length === 2
        ? a.end_latlng[0] + ',' + a.end_latlng[1]
        : '';

      const polyline = INCLUDE_LOCATION_DATA && a.map && a.map.summary_polyline
        ? a.map.summary_polyline
        : '';

      const city = a.location_city || '';

      // Gear fields
      const gearId = a.gear_id || '';
      let gearName = '';
      let gearKmAfterRun = '';

      if (gearId && gearInfo[gearId]) {
        gearName = gearInfo[gearId].name || gearId;
        const totalGearKmNow = gearInfo[gearId].distanceKm || 0;
        const consumed = gearConsumedKm[gearId] || 0;
        gearKmAfterRun = totalGearKmNow - consumed;
        gearConsumedKm[gearId] = consumed + distanceKm;
      }

      // HR zones
      const streams = getActivityStreams_(a.id, accessToken);
      const zones = computeTimeInZones_(streams, MAX_HR);

      sheet.appendRow([
        a.id,
        new Date(a.start_date_local),
        a.type,
        a.sport_type || '',
        a.name,
        city,
        distanceKm,
        movingMinutes,
        a.moving_time,
        a.elapsed_time,
        a.total_elevation_gain,
        a.elev_low || '',
        a.elev_high || '',
        paceSecPerKm,
        paceDisplay,
        avgSpeedKmh,
        maxSpeedKmh,
        a.has_heartrate ? a.average_heartrate : '',
        a.has_heartrate ? a.max_heartrate : '',
        startLatLng,
        endLatLng,
        polyline,
        gearName,
        gearId,
        gearKmAfterRun,
        zones.Z1_min,
        zones.Z2_min,
        zones.Z3_min,
        zones.Z4_min,
        zones.Z5_min
      ]);

      // Splits
      importStravaSplits_(a.id, splitsSheet, accessToken);
    });

    // Append-only sync
    logToSheet_('Syncing StravaData → Data (append-only)...', 'INFO');
    syncStravaToData_();

    // Splits append-only sync
    logToSheet_('Syncing Splits → SplitsDataStored (append-only)...', 'INFO');
    syncSplitsToStored_();

    const durationSec = ((new Date() - startTime) / 1000).toFixed(1);
    logToSheet_('Import completed in ' + durationSec + 's', 'SUCCESS');

    SpreadsheetApp.getUi().alert(
      '✓ Import klar!',
      'Importerade ' + activities.length + ' aktiviteter på ' + durationSec + ' sekunder.\n' +
      'Kolla "Execution Log" för detaljer.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );

  } catch (error) {
    logToSheet_('Import failed: ' + error.message, 'ERROR');
    SpreadsheetApp.getUi().alert(
      '✗ Import misslyckades',
      error.message + '\n\nKolla "Execution Log" för detaljer.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw error;
  }
}

// ============================================
// TOKEN TEST
// ============================================

function testToken() {
  const ui = SpreadsheetApp.getUi();

  try {
    logToSheet_('Testing Strava token...', 'INFO');

    const accessToken = getValidAccessToken_();
    const resp = UrlFetchApp.fetch('https://www.strava.com/api/v3/athlete', {
      method: 'get',
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    });

    if (resp.getResponseCode() === 200) {
      const athlete = JSON.parse(resp.getContentText());
      logToSheet_('Token test successful', 'SUCCESS');
      ui.alert('✓ Token fungerar!', 'Inloggad som: ' + athlete.firstname + ' ' + athlete.lastname, ui.ButtonSet.OK);
    } else {
      logToSheet_('Token test failed: ' + resp.getResponseCode(), 'ERROR');
      ui.alert('✗ Token fungerar inte', 'Felkod: ' + resp.getResponseCode() + '\nKör Setup igen.', ui.ButtonSet.OK);
    }
  } catch (e) {
    logToSheet_('Token test failed: ' + e.message, 'ERROR');
    ui.alert('✗ Fel', e.message + '\nKör Setup igen.', ui.ButtonSet.OK);
  }
}
