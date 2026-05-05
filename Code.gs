/**
 * Serves the web app UI.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Organization Chart')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Returns employee records from the active sheet as JSON-safe objects.
 */
function getOrgData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  var headers = values[0].map(function (h) {
    return String(h || '').trim();
  });

  var idx = {
    name: headers.indexOf('Name'),
    role: headers.indexOf('Role'),
    manager: headers.indexOf('Manager'),
    dep: headers.indexOf('Dep'),
    unit: headers.indexOf('Unit'),
    team: headers.indexOf('Team'),
    photo: headers.indexOf('Photo'),
    level: headers.indexOf('Level')
  };

  Object.keys(idx).forEach(function (key) {
    if (idx[key] === -1) {
      throw new Error('Missing required column: ' + key.toUpperCase());
    }
  });

  return values.slice(1)
    .filter(function (row) {
      return String(row[idx.name] || '').trim() !== '';
    })
    .map(function (row) {
      return {
        name: String(row[idx.name] || '').trim(),
        role: String(row[idx.role] || '').trim(),
        manager: String(row[idx.manager] || '').trim(),
        dep: String(row[idx.dep] || '').trim(),
        unit: String(row[idx.unit] || '').trim(),
        team: String(row[idx.team] || '').trim(),
        photo: normalizeDrivePhotoUrl_(String(row[idx.photo] || '').trim()),
        level: normalizeLevel_(String(row[idx.level] || '').trim())
      };
    });
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function normalizeLevel_(value) {
  var raw = (value || '').toLowerCase();
  if (raw === 'chief') return 'Chief';
  if (raw === 'head') return 'Head';
  if (raw === 'lead') return 'Lead';
  if (raw === 'ic') return 'IC';
  return 'IC';
}

function normalizeDrivePhotoUrl_(url) {
  if (!url) return '';

  var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (match && match[1]) {
    return 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w256-h256';
  }

  var altMatch = url.match(/[?&]id=([a-zA-Z0-9_-]+)/);
  if (altMatch && altMatch[1]) {
    return 'https://drive.google.com/thumbnail?id=' + altMatch[1] + '&sz=w256-h256';
  }

  return url;
}
