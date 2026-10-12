/**
 * DEAL MEMO PUBLISHER  v2.1  (2026-10-12)
 * Bound to the deal model Sheet (lives in the master template; every copy inherits it).
 *
 * What it does: "Deal Memo > Refresh numbers & publish" re-reads the model, refills the
 * numeric tokens of the memo shell Claude stored at render time (deal-memos repository, folder
 * _shells, one file per deal Sheet named by the Sheet's ID), saves the finished HTML to the deal
 * folder, publishes it to the deal-memos site and updates the deal's line in the Deal Analyzer List.
 *
 * What it does NOT do: rewrite the written sections (story, risks, caution note). Those come
 * from Claude's render ("render this deal"). The script warns when the numbers have moved away
 * from the ones the narrative was written on.
 *
 * Nothing in this file is deal-specific. Deal settings travel inside the shell file.
 * Token math mirrors render spec v1.7 (dual lane, template 10c). Do not edit one without the other.
 */

var PUB = {
  GITHUB_OWNER: 'btmolitor',
  GITHUB_REPO: 'deal-memos',
  GITHUB_BRANCH: 'main',
  SITE_BASE: 'https://btmolitor.github.io/deal-memos/',
  PORTFOLIO_DOC_ID: '1EUK54eNCNqoCpcdfveIswl92cbxJRimZO1xnK7Vmlk8',
  KEY_FILE_NAME: 'deal-memos-publish-key (private - do not share)',
  SHELL_DIR: '_shells',
  FIN: 'Financials',
  SUM: 'Summary',
  CLOSING: 0.015, EXIT_CAP: 0.05, SELL: 0.02, TAX: 0.35
};

/* ───────────────────────── menu ───────────────────────── */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Deal Memo')
    .addItem('Refresh numbers & publish', 'refreshAndPublish')
    .addItem('Preview only (save to Drive, do not publish)', 'previewOnly')
    .addSeparator()
    .addItem('Open live memo', 'showLiveLink')
    .addItem('Check setup', 'checkSetup')
    .addItem('Set GitHub key (one time)', 'setGithubKey')
    .addToUi();
}

function refreshAndPublish() { run_(true); }
function previewOnly() { run_(false); }

function run_(publish) {
  var ui = SpreadsheetApp.getUi();
  try {
    SpreadsheetApp.flush();
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var folder = dealFolder_(ss);
    var token = getGithubKey_();
    var parsed = parseShell_(fetchShell_(ss.getId(), token));
    var cfg = parsed.cfg;
    var m = readModel_(ss);
    var ui_ = sliderSettings_(m, cfg);
    var tokens = buildTokens_(m, ui_);
    for (var tk in tokens) if (/NaN|undefined|Infinity/.test(tokens[tk])) throw new Error('Computed value ' + tk + ' came out blank or invalid. Nothing was saved or published.');
    var html = fillShell_(parsed.html, tokens);
    validateOutput_(html);

    var outFile = saveToFolder_(folder, cfg.output_name, html);
    var lines = [];
    lines.push('Numbers read from the Sheet: price ' + money_(m.price) + ', rate ' + pct_(m.rate, 2) +
               ', down ' + pct_(m.down_pct) + ', Year 1 NOI ' + money_(m.noi[0]) + '.');
    lines.push('Saved to the deal folder: ' + outFile.getName());

    if (publish) {
      if (!token) throw new Error('No GitHub key is stored yet. Run Deal Memo > Set GitHub key (one time), then try again.');
      pushToGithub_(cfg.slug + '.html', html, token, 'Regenerate ' + cfg.slug + '.html from Sheet');
      lines.push('Published: ' + PUB.SITE_BASE + cfg.slug + '.html  (allow 1-2 minutes for the site to rebuild)');
      var p = updatePortfolioLine_(cfg, m);
      lines.push(p);
    } else {
      lines.push('Preview only. Nothing was published.');
    }

    var drift = narrativeDrift_(cfg, m);
    if (drift.length) {
      lines.push('');
      lines.push('HEADS UP: the written sections (story, risks, caution note) were written by Claude on different numbers:');
      for (var i = 0; i < drift.length; i++) lines.push('   - ' + drift[i]);
      lines.push('The tables, charts and sliders are current. Ask Claude to "render this deal" to rewrite the text.');
    }
    if (publish) linkDialog_('Deal memo published', lines, PUB.SITE_BASE + cfg.slug + '.html', 'Open the live memo', false);
    else linkDialog_('Deal memo preview saved', lines, outFile.getUrl(), 'Open the preview file in Drive', false);
  } catch (e) {
    ui.alert('Deal memo NOT ' + (publish ? 'published' : 'saved'), String(e && e.message ? e.message : e), ui.ButtonSet.OK);
  }
}

function showLiveLink() {
  var ui = SpreadsheetApp.getUi();
  try {
    var cfg = parseShell_(fetchShell_(SpreadsheetApp.getActiveSpreadsheet().getId(), getGithubKey_())).cfg;
    linkDialog_('Live memo: ' + cfg.name, [], PUB.SITE_BASE + cfg.slug + '.html', 'Open the live memo', true);
  } catch (e) { ui.alert('Live memo', String(e.message || e), ui.ButtonSet.OK); }
}

/** Pop-up with a clickable link that opens in a new browser tab. autoOpen tries to open it straight away. */
function linkDialog_(title, lines, url, label, autoOpen) {
  var esc = function (t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;'); };
  var body = [];
  for (var i = 0; i < lines.length; i++) body.push(lines[i] === '' ? '<div style="height:8px"></div>' : '<div>' + esc(lines[i]) + '</div>');
  var html = '<div style="font:13px/1.5 Arial,sans-serif;color:#202124">' + body.join('') +
    '<p style="margin:14px 0 6px"><a href="' + esc(url) + '" target="_blank" rel="noopener" ' +
    'style="display:inline-block;background:#1a73e8;color:#fff;text-decoration:none;padding:8px 16px;border-radius:4px;font-weight:bold">' + esc(label) + '</a></p>' +
    '<div style="color:#5f6368;font-size:11px;word-break:break-all">' + esc(url) + '</div></div>' +
    (autoOpen ? '<script>try{window.open(' + JSON.stringify(url) + ',"_blank");}catch(e){}</script>' : '');
  var h = Math.min(420, 150 + 22 * lines.length);
  SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(520).setHeight(h), title);
}

function checkSetup() {
  var ui = SpreadsheetApp.getUi(), out = [];
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  try { var folder = dealFolder_(ss); out.push('OK  Deal folder: ' + folder.getName()); } catch (e) { out.push('X   ' + e.message); }
  try { var cfg = parseShell_(fetchShell_(ss.getId(), getGithubKey_())).cfg;
        out.push('OK  Memo shell found for ' + cfg.name + '  (publishes as ' + cfg.slug + '.html, rendered ' + (cfg.rendered || '?') + ')'); } catch (e2) { out.push('X   ' + e2.message); }
  try { var m = readModel_(ss); out.push('OK  Model reads cleanly: ' + money_(m.price) + ', ' + m.units + ' units, Year 1 NOI ' + money_(m.noi[0])); }
  catch (e3) { out.push('X   ' + e3.message); }
  out.push(getGithubKey_() ? 'OK  GitHub key is stored' : 'X   No GitHub key stored. Run Deal Memo > Set GitHub key (one time).');
  ui.alert('Deal Memo setup check', out.join('\n'), ui.ButtonSet.OK);
}

function setGithubKey() {
  var ui = SpreadsheetApp.getUi();
  var r = ui.prompt('GitHub key', 'Paste the GitHub token that can write to ' + PUB.GITHUB_OWNER + '/' + PUB.GITHUB_REPO +
                    '.\nIt is stored once in a private file in your Drive and reused by every deal Sheet.', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() !== ui.Button.OK) return;
  var t = String(r.getResponseText() || '').replace(/\s+/g, '');
  if (t.length < 20) { ui.alert('That does not look like a token. Nothing was saved.'); return; }
  var it = DriveApp.getFilesByName(PUB.KEY_FILE_NAME);
  if (it.hasNext()) it.next().setContent(t); else DriveApp.createFile(PUB.KEY_FILE_NAME, t, 'text/plain');
  ui.alert('Saved. Every deal Sheet you own can now publish.');
}

/* ───────────────────────── Drive / GitHub / Doc ───────────────────────── */

function dealFolder_(ss) {
  var parents = DriveApp.getFileById(ss.getId()).getParents();
  if (!parents.hasNext()) throw new Error('This Sheet is not inside a deal folder in Drive.');
  return parents.next();
}

/** The shell is stored by Claude at render time: deal-memos/_shells/<this Sheet's ID>.html */
function fetchShell_(sheetId, token) {
  var url = 'https://api.github.com/repos/' + PUB.GITHUB_OWNER + '/' + PUB.GITHUB_REPO + '/contents/' + PUB.SHELL_DIR + '/' + sheetId + '.html?ref=' + PUB.GITHUB_BRANCH;
  var headers = { Accept: 'application/vnd.github.raw', 'X-GitHub-Api-Version': '2022-11-28' };
  if (token) headers.Authorization = 'Bearer ' + token;
  var r = UrlFetchApp.fetch(url, { method: 'get', headers: headers, muteHttpExceptions: true });
  var c = r.getResponseCode();
  if (c === 200) return r.getContentText('UTF-8');
  if (c === 404) throw new Error('No memo has been rendered for this Sheet yet (or this is a copy of a deal Sheet). Ask Claude to "render this deal" once; after that this button works.');
  if (c === 401) throw new Error('GitHub rejected the stored key (401). It may have expired. Run Deal Memo > Set GitHub key with a new token.');
  throw new Error('Could not read the memo shell from GitHub (' + c + '): ' + r.getContentText().slice(0, 200));
}

function saveToFolder_(folder, name, html) {
  var it = folder.getFilesByName(name);
  if (it.hasNext()) { var f = it.next(); f.setContent(html); return f; }
  return folder.createFile(name, html, 'text/html');
}

function getGithubKey_() {
  var it = DriveApp.getFilesByName(PUB.KEY_FILE_NAME);
  while (it.hasNext()) {
    var f = it.next();
    if (f.isTrashed()) continue;
    var t = f.getBlob().getDataAsString('UTF-8').replace(/\s+/g, '');
    if (t) return t;
  }
  return '';
}

function pushToGithub_(path, html, token, message) {
  var url = 'https://api.github.com/repos/' + PUB.GITHUB_OWNER + '/' + PUB.GITHUB_REPO + '/contents/' + encodeURIComponent(path);
  var headers = { Authorization: 'Bearer ' + token, Accept: 'application/vnd.github+json', 'X-GitHub-Api-Version': '2022-11-28' };
  var sha = null;
  var g = UrlFetchApp.fetch(url + '?ref=' + PUB.GITHUB_BRANCH, { method: 'get', headers: headers, muteHttpExceptions: true });
  if (g.getResponseCode() === 200) sha = JSON.parse(g.getContentText()).sha;
  else if (g.getResponseCode() === 401 || g.getResponseCode() === 403) throw new Error('GitHub rejected the stored key (' + g.getResponseCode() + '). It may have expired. Run Deal Memo > Set GitHub key with a new token.');
  else if (g.getResponseCode() !== 404) throw new Error('GitHub lookup failed (' + g.getResponseCode() + '): ' + g.getContentText().slice(0, 200));
  var body = { message: message, branch: PUB.GITHUB_BRANCH, content: Utilities.base64Encode(html, Utilities.Charset.UTF_8) };
  if (sha) body.sha = sha;
  var p = UrlFetchApp.fetch(url, { method: 'put', headers: headers, contentType: 'application/json', payload: JSON.stringify(body), muteHttpExceptions: true });
  if (p.getResponseCode() !== 200 && p.getResponseCode() !== 201) throw new Error('GitHub publish failed (' + p.getResponseCode() + '): ' + p.getContentText().slice(0, 200));
}

/** Updates "[Name] — N units · $X.XXM · Y1 NOI $XXXK —" on the deal's existing line. Never adds a line. Never fatal. */
function updatePortfolioLine_(cfg, m) {
  try {
    var body = DocumentApp.openById(PUB.PORTFOLIO_DOC_ID).getBody();
    var paras = body.getParagraphs();
    var fresh = portfolioStats_(m);
    for (var i = 0; i < paras.length; i++) {
      var txt = paras[i].getText();
      if (txt.indexOf(cfg.name) === 0) {
        if (txt.indexOf(fresh) !== -1) return 'Deal Analyzer List: already current.';
        var found = paras[i].findText('\\d+ units · \\$[0-9.]+M · Y1 NOI \\$[0-9,]+K');
        if (!found) return 'Deal Analyzer List: found the ' + cfg.name + ' line but not its numbers; left as is.';
        found.getElement().asText().deleteText(found.getStartOffset(), found.getEndOffsetInclusive())
             .insertText(found.getStartOffset(), fresh);
        return 'Deal Analyzer List: numbers updated.';
      }
    }
    return 'Deal Analyzer List: no line for ' + cfg.name + ' yet (Claude adds it on first render).';
  } catch (e) { return 'Deal Analyzer List not updated (' + (e.message || e) + '). The memo itself is published.'; }
}

function portfolioStats_(m) {
  return m.units + ' units · ' + compact_(m.price) + ' · Y1 NOI $' + Math.round(m.noi[0] / 1000) + 'K';
}

/* ───────────────────────── model read (render spec cell map) ───────────────────────── */

function readModel_(ss) {
  var F = ss.getSheetByName(PUB.FIN), S = ss.getSheetByName(PUB.SUM);
  if (!F || !S) throw new Error('This Sheet needs tabs named "' + PUB.FIN + '" and "' + PUB.SUM + '".');
  // Label guards: stop rather than publish from the wrong rows if the template layout ever moves.
  label_(F, 'B9', 'Mortgage Principal');
  label_(F, 'B143', 'Annual Debt Service');
  label_(F, 'B32', 'Market Value Growth');
  label_(S, 'H5', 'Units');
  label_(S, 'H15', 'Annual Depreciation');
  label_(S, 'B20', 'Gross Potential');
  label_(S, 'B24', 'Net Operating Income');
  var row = function (sh, r) { return sh.getRange(r, 3, 1, 10).getValues()[0]; };
  var m = {
    price: F.getRange('C1').getValue(), down_pct: F.getRange('C7').getValue(), rate: F.getRange('C10').getValue(),
    amort: F.getRange('C11').getValue(), growth: F.getRange('C32').getValue(),
    boy_balances: row(F, 9), ds_y1: F.getRange('C143').getValue(), ds_y6: F.getRange('H143').getValue(),
    units: S.getRange('I5').getValue(), annual_dep: S.getRange('I15').getValue(),
    gpr: row(S, 20), vacancy: row(S, 21), egi: row(S, 22), opex: row(S, 23), noi: row(S, 24)
  };
  return checkModel_(m);
}

function label_(sh, a1, expect) {
  var v = String(sh.getRange(a1).getValue());
  if (v.toLowerCase().indexOf(expect.toLowerCase()) === -1)
    throw new Error('Layout check failed: ' + sh.getName() + '!' + a1 + ' should read "' + expect + '" but reads "' + v +
                    '". The template layout has changed; ask Claude to update the publisher before publishing.');
}

function checkModel_(m) {
  var isNum = function (x) { return typeof x === 'number' && isFinite(x); };
  var scalars = ['price', 'down_pct', 'rate', 'amort', 'growth', 'ds_y1', 'ds_y6', 'units', 'annual_dep'];
  for (var i = 0; i < scalars.length; i++) if (!isNum(m[scalars[i]])) throw new Error('Model value "' + scalars[i] + '" is not a number (error or blank cell in the Sheet).');
  var arrs = ['boy_balances', 'gpr', 'vacancy', 'egi', 'opex', 'noi'];
  for (var j = 0; j < arrs.length; j++) {
    var a = m[arrs[j]];
    if (!a || a.length !== 10) throw new Error('Model row "' + arrs[j] + '" does not have 10 years.');
    for (var k = 0; k < 10; k++) if (!isNum(a[k])) throw new Error('Model row "' + arrs[j] + '", year ' + (k + 1) + ' is not a number.');
  }
  if (m.price <= 0 || m.units <= 0) throw new Error('Price and unit count must be above zero.');
  if (m.down_pct <= 0 || m.down_pct >= 1) throw new Error('Down payment % must be between 0 and 100%.');
  if (m.rate <= 0 || m.rate > 0.25) throw new Error('Mortgage rate looks wrong: ' + m.rate);
  if (m.amort < 5 || m.amort > 40) throw new Error('Amortization years looks wrong: ' + m.amort);
  var loan = m.price * (1 - m.down_pct);
  if (Math.abs(m.boy_balances[0] - loan) > 1) throw new Error('Loan does not tie: price x (1 - down) = ' + money_(loan) + ' but Financials!C9 = ' + money_(m.boy_balances[0]) + '. (Improvements or a note in the model? Ask Claude to re-render.)');
  if (m.noi[0] <= 0) throw new Error('Year 1 NOI is zero or negative.');
  return m;
}

/* ───────────────────────── shell ───────────────────────── */

function parseShell_(text) {
  var re = /<!--DEAL_MEMO_SHELL_CONFIG\s*([\s\S]*?)\s*-->\n?/;
  var mm = re.exec(text);
  if (!mm) throw new Error('The memo shell has no settings block. Ask Claude to "render this deal" again.');
  var cfg = JSON.parse(mm[1]);
  var need = ['slug', 'name', 'output_name', 'price_min', 'price_max', 'price_step', 'capital_min', 'capital_step', 'capital_default'];
  for (var i = 0; i < need.length; i++) if (cfg[need[i]] === undefined || cfg[need[i]] === '') throw new Error('Memo shell setting "' + need[i] + '" is missing. Ask Claude to "render this deal" again.');
  if (!/^[a-z0-9]+$/.test(cfg.slug)) throw new Error('Memo shell slug must be lowercase letters and digits only.');
  return { cfg: cfg, html: text.replace(re, '') };
}

/** Slider settings: price window from the shell unless the price has moved outside it; refi window follows the rate. */
function sliderSettings_(m, cfg) {
  var price = Math.round(m.price), pmin = cfg.price_min, pmax = cfg.price_max;
  if (price < pmin || price > pmax) {
    pmin = Math.max(250000, Math.round((price - 1500000) / 250000) * 250000);
    pmax = Math.round((price + 1500000) / 250000) * 250000;
  }
  var bp = Math.round(m.rate * 10000);
  var tcl = Math.round(price * (round_(m.down_pct + PUB.CLOSING, 6)));
  var capDefault = Math.min(cfg.capital_default, tcl);
  return { price_min: pmin, price_max: pmax, price_step: cfg.price_step,
           capital: { min: Math.min(cfg.capital_min, capDefault), step: cfg.capital_step, 'default': capDefault },
           refi_min: bp - 125, refi_max: bp + 175, refi_step: 5 };
}

function narrativeDrift_(cfg, m) {
  var b = cfg.narrative_basis, out = [];
  if (!b) return out;
  if (b.price !== undefined && Math.round(b.price) !== Math.round(m.price)) out.push('price was ' + money_(b.price) + ', now ' + money_(m.price));
  if (b.rate !== undefined && Math.abs(b.rate - m.rate) > 0.000001) out.push('rate was ' + pct_(b.rate, 2) + ', now ' + pct_(m.rate, 2));
  if (b.down_pct !== undefined && Math.abs(b.down_pct - m.down_pct) > 0.000001) out.push('down payment was ' + pct_(b.down_pct) + ', now ' + pct_(m.down_pct));
  if (b.noi_y1 !== undefined && Math.abs(b.noi_y1 - m.noi[0]) > 500) out.push('Year 1 NOI was ' + money_(b.noi_y1) + ', now ' + money_(m.noi[0]));
  return out;
}

function fillShell_(shell, tokens) {
  var html = shell;
  for (var k in tokens) if (tokens.hasOwnProperty(k)) html = html.split('{{' + k + '}}').join(tokens[k]);
  return html;
}

function validateOutput_(html) {
  var left = html.match(/\{\{[A-Z0-9_]+\}\}/g);
  if (left) throw new Error('Unfilled placeholders remain: ' + left.slice(0, 5).join(', ') + '. Ask Claude to "render this deal" again.');
  var need = { 'snap-cell': [/<div class="snap-cell">/g, 12], 'partner-cell': [/<div class="partner-cell">/g, 6], 'risk-card': [/<div class="risk-card">/g, 6] };
  for (var n in need) { var c = (html.match(need[n][0]) || []).length; if (c !== need[n][1]) throw new Error('Structure check failed: ' + n + ' count is ' + c + ', expected ' + need[n][1] + '.'); }
  if (!/const BASE = \{/.test(html)) throw new Error('Structure check failed: data block missing.');
}

/* ───────────────────────── token math (render spec v1.7, dual lane) ───────────────────────── */

function pmtConst_(rate, years) { var r = rate / 12, n = years * 12; return r / (1 - Math.pow(1 + r, -n)) * 12; }
function balAfter_(loan, rate, years, months) {
  var r = rate / 12, n = years * 12, p = loan * r / (1 - Math.pow(1 + r, -n));
  return loan * Math.pow(1 + r, months) - p * (Math.pow(1 + r, months) - 1) / r;
}
function round_(x, d) { var f = Math.pow(10, d); return Math.round(x * f) / f; }
function num_(x) { return String(round_(Number(x), 10)); }
function compact_(n) { return '$' + (n / 1e6).toFixed(2) + 'M'; }
function kfmt_(n) {
  if (n >= 1e6) { var mm = n / 1e6; return Math.abs(mm - Math.round(mm)) < 1e-9 ? '$' + Math.round(mm) + 'M' : '$' + mm.toFixed(2) + 'M'; }
  return '$' + Math.round(n / 1e3) + 'K';
}
function arr_(vals, dec) { var o = []; for (var i = 0; i < vals.length; i++) o.push(fix_(vals[i], dec)); return '[' + o.join(', ') + ']'; }
function fix_(v, dec) { var s = Number(v).toFixed(dec); return (Number(s) === 0 && s.charAt(0) === '-') ? s.slice(1) : s; }
function iarr_(vals) { var o = []; for (var i = 0; i < vals.length; i++) o.push(String(Math.round(vals[i]))); return '[' + o.join(', ') + ']'; }
function pctTrim_(x) { var s = (x * 100).toFixed(2); s = s.replace(/0+$/, '').replace(/\.$/, ''); return s + '%'; }
function commas_(n) { return String(Math.round(n)).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function money_(n) { return '$' + commas_(n); }
function pct_(x, d) { return d === undefined ? pctTrim_(x) : (x * 100).toFixed(d) + '%'; }

function buildTokens_(m, ui) {
  var price = Math.round(m.price), down = m.down_pct, rate = m.rate, amort = Math.round(m.amort);
  var loan = round_(price * (1 - down), 2), cashPct = round_(down + PUB.CLOSING, 6), cashIn = price * cashPct;
  var i, k;
  // STD lane: Financials row 9 (BOY1..BOY10) -> EOY1..EOY9; EOY10 computed on the post-refi schedule
  var eoy = m.boy_balances.slice(1, 10), eoy5 = eoy[4];
  eoy.push(balAfter_(eoy5, rate, amort, 60));
  var stdBal = [], stdWf = [];
  for (i = 0; i < 10; i++) { stdBal.push(eoy[i] / loan); stdWf.push(((i === 0 ? loan : eoy[i - 1]) - eoy[i]) / cashIn); }
  // IO lane: 2 years interest-only, then (amort-2)-year schedule Y3-Y5, refi at EOY5 on a fresh amort
  var c28 = pmtConst_(rate, amort - 2);
  var io = [loan, loan];
  for (k = 1; k <= 3; k++) io.push(balAfter_(loan, rate, amort - 2, 12 * k));
  var io5 = io[4];
  for (k = 1; k <= 5; k++) io.push(balAfter_(io5, rate, amort, 12 * k));
  var ioBal = [], ioWf = [];
  for (i = 0; i < 10; i++) { ioBal.push(io[i] / loan); ioWf.push(((i === 0 ? loan : io[i - 1]) - io[i]) / price); }
  var ioPost = io5 * pmtConst_(rate, amort) / loan;
  var ioDs = [rate, rate, c28, c28, c28, ioPost, ioPost, ioPost, ioPost, ioPost];
  var depShield = round_(m.annual_dep * PUB.TAX, 2);
  var cap = ui.capital, presets = [100000, 250000, 500000, 750000, 1000000];
  var tcl = Math.round(cashIn), half = tcl / 2, btns = [];
  for (i = 0; i < presets.length; i++) {
    var v = presets[i];
    btns.push('          <button class="' + (v === cap['default'] ? 'preset-btn active' : 'preset-btn') + '" data-val="' + v + '">' + kfmt_(v) + '</button>');
  }
  btns.push('          <button class="preset-btn" data-val="' + num_(half) + '">' + compact_(half) + ' (half)</button>');
  btns.push('          <button class="preset-btn" data-val="' + tcl + '">' + compact_(tcl) + ' (100%)</button>');
  var presetsHtml = '<div class="preset-buttons">\n' + btns.join('\n') + '\n        </div>';
  var ticksHtml = '<div class="slider-ticks">\n            <span>' + kfmt_(cap.min) + '</span><span>' + kfmt_(500000) + '</span><span>' + kfmt_(1000000) +
                  '</span><span>' + compact_(half) + '</span><span>' + compact_(tcl) + ' (full)</span>\n          </div>';
  return {
    DUAL_BASE_PRICE: String(price), DUAL_BASE_PRICE_COMPACT: compact_(price),
    DUAL_PRICE_MIN: String(ui.price_min), DUAL_PRICE_MAX: String(ui.price_max), DUAL_PRICE_STEP: String(ui.price_step),
    DUAL_PRICE_MIN_COMPACT: compact_(ui.price_min), DUAL_PRICE_MAX_COMPACT: compact_(ui.price_max),
    DUAL_DOWN_PCT: num_(down), DUAL_CLOSING_RESERVES_PCT: num_(PUB.CLOSING), DUAL_LOAN_PCT: num_(round_(1 - down, 10)),
    DUAL_CASH_IN_PCT: num_(cashPct), DUAL_DEP_SHIELD_PER_DOLLAR: num_(depShield) + ' / ' + price,
    DUAL_DS_RATIO_Y1_Y5: Math.round(Math.abs(m.ds_y1)) + ' / ' + num_(loan),
    DUAL_DS_RATIO_Y6_Y10: Math.round(Math.abs(m.ds_y6)) + ' / ' + num_(loan),
    DUAL_STD_BAL_RATIOS_Y1_Y10: arr_(stdBal, 6), DUAL_STD_WF_PRINC_RATIOS_Y1_Y10: arr_(stdWf, 4),
    DUAL_IO_BAL_RATIOS_Y1_Y10: arr_(ioBal, 6), DUAL_IO_WF_PRINC_RATIOS_Y1_Y10: arr_(ioWf, 6), DUAL_IO_DS_BY_YEAR_RATIOS: arr_(ioDs, 6),
    DUAL_GPR_Y1_Y10: iarr_(m.gpr), DUAL_VACANCY_Y1_Y10: iarr_(m.vacancy), DUAL_EGI_Y1_Y10: iarr_(m.egi),
    DUAL_OPEX_Y1_Y10: iarr_(m.opex), DUAL_NOI_Y1_Y10: iarr_(m.noi),
    DUAL_MARKET_GROWTH: num_(m.growth), DUAL_EXIT_CAP: num_(PUB.EXIT_CAP), DUAL_EXIT_SELLING_COST_PCT: num_(PUB.SELL),
    DUAL_MORTGAGE_RATE: num_(rate), DUAL_AMORT_YEARS: String(amort),
    DUAL_DOWN_PCT_FORMATTED: pctTrim_(down), DUAL_MORTGAGE_RATE_PCT: (rate * 100).toFixed(2) + '%', DUAL_TAX_BRACKET_PCT: pctTrim_(PUB.TAX),
    DUAL_CAPITAL_MIN: String(cap.min), DUAL_CAPITAL_MAX: String(tcl), DUAL_CAPITAL_STEP: String(cap.step), DUAL_CAPITAL_DEFAULT: String(cap['default']),
    DUAL_CAPITAL_DEFAULT_FORMATTED: '$' + commas_(cap['default']), DUAL_INITIAL_SHARE_PCT: (cap['default'] / cashIn * 100).toFixed(2) + '%',
    DUAL_CAPITAL_TICKS_HTML: ticksHtml, DUAL_CAPITAL_PRESETS_HTML: presetsHtml,
    DUAL_REFI_RATE_MIN: String(ui.refi_min), DUAL_REFI_RATE_MAX: String(ui.refi_max), DUAL_REFI_RATE_STEP: String(ui.refi_step),
    DUAL_REFI_RATE_DEFAULT: String(Math.round(rate * 10000))
  };
}
