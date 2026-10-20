/**
 * DEAL MEMO LOADER  v1.0  (2026-10-20)
 *
 * This is the ONLY code that lives inside a deal Sheet (Extensions > Apps Script). It is pasted once into the
 * master template and once into each live deal Sheet, and never again: every click on the Deal Memo menu
 * fetches the current _publisher/Code.gs from GitHub (btmolitor/deal-memos, branch main) and runs it, so a
 * change to Code.gs in the repo is live in every Sheet immediately. New deals copied from the master inherit it.
 *
 * The menu is built here (a simple onOpen trigger may not call the network), so the menu items are the one
 * thing that still needs a paste when they change. Everything else — publishing, the comp map, PDF export,
 * setup checks — lives in Code.gs.
 */

var LOADER = {
  RAW_URL: 'https://raw.githubusercontent.com/btmolitor/deal-memos/main/_publisher/Code.gs',
  API_URL: 'https://api.github.com/repos/btmolitor/deal-memos/contents/_publisher/Code.gs?ref=main',
  KEY_FILE_NAME: 'deal-memos-publish-key (private - do not share)'
};

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Deal Memo')
    .addItem('Refresh numbers & publish', 'refreshAndPublish')
    .addItem('Preview only (save to Drive, do not publish)', 'previewOnly')
    .addSeparator()
    .addItem('Export Summary PDF to deal folder', 'exportSummaryPdf')
    .addSeparator()
    .addItem('Open live memo', 'showLiveLink')
    .addItem('Check setup', 'checkSetup')
    .addItem('Set GitHub key (one time)', 'setGithubKey')
    .addToUi();
}

function refreshAndPublish() { publisher_().refreshAndPublish(); }
function previewOnly()       { publisher_().previewOnly(); }
function exportSummaryPdf()  { publisher_().exportSummaryPdf(); }
function showLiveLink()      { publisher_().showLiveLink(); }
function checkSetup()        { publisher_().checkSetup(); }
function setGithubKey()      { publisher_().setGithubKey(); }

/** Fetches Code.gs from the repo and returns its top-level functions as an object. */
function publisher_() {
  var src = fetchSource_();
  var names = [], re = /^(?:async\s+)?function\s+([A-Za-z_]\w*)\s*\(/gm, m;
  while ((m = re.exec(src)) !== null) names.push(m[1]);
  if (names.indexOf('refreshAndPublish') === -1) throw new Error('The file fetched from GitHub does not look like the Deal Memo Publisher.');
  var exports = 'return {' + names.map(function (n) { return n + ':' + n; }).join(',') + '};';
  try {
    return new Function(src + '\n' + exports)();
  } catch (e) {
    throw new Error('The publisher code in GitHub has an error: ' + (e && e.message ? e.message : e));
  }
}

function fetchSource_() {
  // The repo is public, so the raw file needs no key; the timestamp defeats the CDN cache.
  var r = UrlFetchApp.fetch(LOADER.RAW_URL + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (r.getResponseCode() === 200) return r.getContentText('UTF-8');
  // Fallback: the GitHub API, with the stored key if there is one (also works if the repo is ever made private).
  var headers = { Accept: 'application/vnd.github.raw', 'X-GitHub-Api-Version': '2022-11-28' };
  var key = storedKey_(); if (key) headers.Authorization = 'Bearer ' + key;
  var r2 = UrlFetchApp.fetch(LOADER.API_URL, { headers: headers, muteHttpExceptions: true });
  if (r2.getResponseCode() === 200) return r2.getContentText('UTF-8');
  throw new Error('Could not load the Deal Memo Publisher from GitHub (' + r.getResponseCode() + ' / ' + r2.getResponseCode() + '). Check the internet connection and that btmolitor/deal-memos still has _publisher/Code.gs, then try again.');
}

function storedKey_() {
  var it = DriveApp.getFilesByName(LOADER.KEY_FILE_NAME);
  while (it.hasNext()) { var f = it.next(); if (f.isTrashed()) continue; var t = f.getBlob().getDataAsString('UTF-8').replace(/\s+/g, ''); if (t) return t; }
  return '';
}
