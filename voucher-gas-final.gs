var SHEET_ID = '1-0JMPMBtJTJ4WVUz_MnkzHYAkUqWmLmpsmtKTN69464';
var SHEET_NAME = 'Vouchers';
var REPORT_SHEET_ID = '1txRlZ0uuUg1vW5CwqTZ47x2CuVzADlVxIVbzq-vscW4';
var REPORT_SHEET_NAME = 'Report Voucher';
var GITHUB_BASE = 'https://pippopa1975.github.io/mak-voucher/';

function doGet(e) {
  var action = (e.parameter.action || '').toString().trim();
  var code   = (e.parameter.code  || '').toString().trim().toUpperCase();
  var vCode  = (e.parameter.v     || '').toString().trim().toUpperCase();
  var cb     = e.parameter.callback || 'callback';

  if (action === 'generate') {
    var nome       = e.parameter.nome       || '';
    var email      = e.parameter.email      || '';
    var telefono   = e.parameter.telefono   || '';
    var compleanno = e.parameter.compleanno || '';
    if (!nome || !compleanno) {
      return jsonpOut(cb, {status:'error', message:'Parametri mancanti'});
    }
    return jsonpOut(cb, generateVoucher(nome, email, telefono, compleanno));
  }

  if (action === 'check' || action === 'redeem') {
    var result = (action === 'check') ? checkVoucher(code) : redeemVoucher(code);
    return jsonpOut(cb, result);
  }

  // Nessuna action → serve la pagina HTML di validazione
  return HtmlService.createHtmlOutput(buildHtml(vCode))
    .setTitle('Mak Mixology — Voucher')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function serverCheck(code) {
  return checkVoucher(code.toString().trim().toUpperCase());
}
function serverRedeem(code) {
  return redeemVoucher(code.toString().trim().toUpperCase());
}

/* ── CORE LOGIC ─────────────────────────────────────────────────── */

function checkVoucher(code) {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() !== code) continue;
    var stato    = data[i][7].toString().trim();
    var scadenza = new Date(data[i][6]); scadenza.setHours(23,59,59,999);
    var oggi     = new Date(); oggi.setHours(0,0,0,0);
    if (stato === 'Riscattato') {
      return {status:'redeemed', nome:data[i][1], scadenza:fmtIT(new Date(data[i][6])), dataRiscatto:fmtIT(new Date(data[i][8]))};
    }
    if (scadenza < oggi) {
      sheet.getRange(i+1,8).setValue('Scaduto');
      return {status:'expired', nome:data[i][1], scadenza:fmtIT(new Date(data[i][6]))};
    }
    return {status:'valid', nome:data[i][1], scadenza:fmtIT(new Date(data[i][6]))};
  }
  return {status:'not_found'};
}

function redeemVoucher(code) {
  var sheet = getSheet();
  var data  = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() !== code) continue;
    var stato    = data[i][7].toString().trim();
    if (stato === 'Riscattato') return {success:false, message:'Già riscattato.'};
    var scadenza = new Date(data[i][6]); scadenza.setHours(23,59,59,999);
    if (scadenza < new Date()) { sheet.getRange(i+1,8).setValue('Scaduto'); return {success:false, message:'Voucher scaduto.'}; }
    sheet.getRange(i+1,8).setValue('Riscattato');
    sheet.getRange(i+1,9).setValue(fmtIT(new Date()));
    logRedeemToReport(code, fmtIT(new Date()));
    return {success:true, message:'Drink omaggio pronto!'};
  }
  return {success:false, message:'Voucher non trovato.'};
}

function generateVoucher(nome, email, telefono, compleanno) {
  var sheet    = getSheet();
  var nomePart = nome.replace(/[^A-Za-z]/g,'').toUpperCase().substring(0,4);
  while (nomePart.length < 4) nomePart += 'X';
  var bd = new Date(compleanno);
  var dd = ('0'+bd.getDate()).slice(-2);
  var mm = ('0'+(bd.getMonth()+1)).slice(-2);
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var rnd = ''; for (var i=0;i<4;i++) rnd += chars.charAt(Math.floor(Math.random()*chars.length));
  var codice = 'MAK-'+nomePart+'-'+dd+mm+'-'+rnd;
  var oggi     = new Date();
  var bdQuest  = new Date(oggi.getFullYear(), bd.getMonth(), bd.getDate());
  if (bdQuest < oggi) bdQuest.setFullYear(oggi.getFullYear()+1);
  var scadenza = new Date(bdQuest.getTime()); scadenza.setDate(scadenza.getDate()+10);
  sheet.appendRow([codice,nome,email,telefono,fmtIT(bd),fmtIT(new Date()),fmtIT(scadenza),'Valido','']);
  logEmitToReport(codice, nome, fmtIT(new Date()), fmtIT(scadenza));
  var validationUrl = GITHUB_BASE + '?v=' + codice;
  var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' + encodeURIComponent(validationUrl);
  return {codice:codice, qrUrl:qrUrl, scadenza:fmtIT(scadenza), validationUrl:validationUrl};
}

/* ── REPORT VOUCHER ─────────────────────────────────────────────── */

function getReportSheet() {
  try { return SpreadsheetApp.openById(REPORT_SHEET_ID).getSheetByName(REPORT_SHEET_NAME); }
  catch(e) { return null; }
}
function logEmitToReport(codice, nome, dataEmissione, scadenza) {
  var rs = getReportSheet(); if (!rs) return;
  rs.appendRow([codice, nome, dataEmissione, scadenza, '', 'Emesso']);
}
function logRedeemToReport(code, dataUtilizzo) {
  var rs = getReportSheet(); if (!rs) return;
  var data = rs.getDataRange().getValues();
  for (var j = 1; j < data.length; j++) {
    if (data[j][0].toString().trim().toUpperCase() === code) {
      rs.getRange(j+1,5).setValue(dataUtilizzo);
      rs.getRange(j+1,6).setValue('Riscattato');
      return;
    }
  }
  rs.appendRow([code,'','','',dataUtilizzo,'Riscattato']);
}

/* ── UTILS ──────────────────────────────────────────────────────── */

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var s  = ss.getSheetByName(SHEET_NAME);
  if (!s) {
    s = ss.insertSheet(SHEET_NAME);
    s.appendRow(['CodiceVoucher','Nome','Email','Telefono','DataCompleanno','DataEmissione','DataScadenza','Stato','DataRiscatto']);
    s.getRange(1,1,1,9).setFontWeight('bold').setBackground('#c9a84c');
  }
  return s;
}
function fmtIT(d) {
  if (!d || isNaN(d.getTime())) return '';
  return ('0'+d.getDate()).slice(-2)+'/'+('0'+(d.getMonth()+1)).slice(-2)+'/'+d.getFullYear();
}
function jsonpOut(cb, obj) {
  var out = ContentService.createTextOutput(cb+'('+JSON.stringify(obj)+')');
  out.setMimeType(ContentService.MimeType.JAVASCRIPT);
  return out;
}
function testCheck()    { Logger.log(JSON.stringify(checkVoucher('MAK-TEST-0101-XXXX'))); }
function testGenerate() { Logger.log(JSON.stringify(generateVoucher('Mario Rossi','t@t.com','333','1990-03-15'))); }

/* ── HTML PAGE (servita direttamente dall'Apps Script) ──────────── */

function buildHtml(preloadCode) {
  var GAS_URL = ScriptApp.getService().getUrl();
  var qrDataUrl = preloadCode
    ? 'https://api.qrserver.com/v1/create-qr-code/?size=220x220&color=c9a84c&bgcolor=1a1a1a&data=' + encodeURIComponent(GITHUB_BASE + '?v=' + preloadCode)
    : '';

  return '<!DOCTYPE html><html lang="it"><head>' +
  '<meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1">' +
  '<title>Mak Mixology — Voucher</title>' +
  '<style>' +
  '*{box-sizing:border-box;margin:0;padding:0}' +
  'body{background:#0a0a0a;font-family:-apple-system,sans-serif;color:#fff;min-height:100vh;display:flex;flex-direction:column;align-items:center;padding:28px 16px}' +
  '.logo-img{width:90px;margin-bottom:24px}' +
  '.card{background:#111;border:1px solid #c9a84c44;border-radius:18px;padding:28px 22px;max-width:380px;width:100%;text-align:center}' +
  'h2{color:#c9a84c;font-size:18px;letter-spacing:2px;margin-bottom:6px}' +
  '.code{font-size:20px;font-weight:700;letter-spacing:3px;color:#fff;margin:10px 0}' +
  '.sub{font-size:13px;color:#888;margin-bottom:18px}' +
  '.status-box{border-radius:12px;padding:18px 16px;margin:16px 0;font-size:15px;line-height:1.6;display:none}' +
  '.valid-box{background:#0d2b0d;border:1px solid #2d7a2d;color:#7fe87f}' +
  '.redeemed-box{background:#2b1a00;border:1px solid #c9a84c;color:#c9a84c}' +
  '.error-box{background:#2b0d0d;border:1px solid #7a2d2d;color:#e87f7f}' +
  '.loading-box{background:#1a1a1a;border:1px solid #555;color:#aaa}' +
  '.qr-wrap{margin:18px 0 6px}' +
  '.qr-wrap img{border-radius:10px;border:2px solid #c9a84c44}' +
  '.qr-label{font-size:12px;color:#666;margin-top:6px}' +
  '.staff-btn{margin-top:24px;background:none;border:1px solid #333;color:#555;font-size:12px;padding:8px 18px;border-radius:20px;cursor:pointer;width:100%}' +
  '.staff-btn:active{background:#1a1a1a}' +
  '.overlay{position:fixed;inset:0;background:rgba(0,0,0,.85);display:flex;align-items:center;justify-content:center;z-index:100;display:none}' +
  '.pin-box{background:#1a1a1a;border:1px solid #c9a84c;border-radius:16px;padding:28px 24px;width:300px;text-align:center}' +
  '.pin-box h3{color:#c9a84c;margin-bottom:16px;font-size:16px}' +
  '.pin-input{width:100%;background:#111;border:1px solid #444;color:#fff;font-size:24px;letter-spacing:6px;text-align:center;padding:12px;border-radius:8px;margin-bottom:14px}' +
  '.btn{width:100%;padding:13px;border-radius:10px;border:none;font-size:15px;font-weight:700;cursor:pointer;margin-top:6px}' +
  '.btn-confirm{background:#c9a84c;color:#000}' +
  '.btn-cancel{background:#333;color:#aaa}' +
  '.btn-redeem{background:#1a5c1a;color:#7fe87f;border:1px solid #2d7a2d;display:none;margin-top:14px}' +
  '</style></head><body>' +
  '<img class="logo-img" src="https://raw.githubusercontent.com/pippopa1975/mak-voucher/main/logo.png" onerror="this.style.display=\'none\'">' +
  '<div class="card">' +
  '  <h2>VOUCHER OMAGGIO</h2>' +
  '  <div class="code" id="codeDisplay">' + (preloadCode || '—') + '</div>' +
  '  <div class="sub">Drink di compleanno 🎂</div>' +
  '  <div class="status-box loading-box" id="statusBox" style="display:block">⏳ Verifica in corso…</div>' +
  (preloadCode && qrDataUrl
    ? '  <div class="qr-wrap" id="qrWrap" style="display:none"><img src="' + qrDataUrl + '" width="180" height="180"><div class="qr-label">📱 Fai scansionare al personale</div></div>'
    : '') +
  '  <button class="btn btn-redeem" id="btnRedeem" onclick="showPin()">✅ Segna come riscattato</button>' +
  '  <button class="staff-btn" id="btnStaff" onclick="showPin()" style="display:none">👤 Staff — Riscatta voucher</button>' +
  '</div>' +
  '<div class="overlay" id="overlay">' +
  '  <div class="pin-box">' +
  '    <h3>🔐 PIN Staff</h3>' +
  '    <input class="pin-input" type="password" inputmode="numeric" id="pinInput" maxlength="6" placeholder="• • • •">' +
  '    <button class="btn btn-confirm" onclick="checkPin()">Conferma</button>' +
  '    <button class="btn btn-cancel" onclick="hidePin()">Annulla</button>' +
  '  </div>' +
  '</div>' +
  '<script>' +
  'var VOUCHER_CODE = "' + preloadCode + '";' +
  'var PIN_CORRECT  = "1250";' +
  'var pinVerified  = false;' +

  'window.onload = function() {' +
  '  if (!VOUCHER_CODE) { showBox("loading-box","Nessun codice specificato."); return; }' +
  '  google.script.run.withSuccessHandler(onCheck).withFailureHandler(onErr).serverCheck(VOUCHER_CODE);' +
  '};' +

  'function onCheck(d) {' +
  '  if (d.status === "valid") {' +
  '    showBox("valid-box", "✅ VOUCHER VALIDO\\n👤 " + d.nome + "\\n📅 Scadenza: " + d.scadenza);' +
  '    var qr = document.getElementById("qrWrap"); if(qr) qr.style.display="block";' +
  '    document.getElementById("btnStaff").style.display = "block";' +
  '  } else if (d.status === "redeemed") {' +
  '    showBox("redeemed-box", "⚠️ GIÀ RISCATTATO\\n👤 " + d.nome + "\\n🗓 " + d.dataRiscatto);' +
  '  } else if (d.status === "expired") {' +
  '    showBox("error-box", "❌ VOUCHER SCADUTO\\n📅 " + d.scadenza);' +
  '  } else {' +
  '    showBox("error-box", "❌ VOUCHER NON TROVATO");' +
  '  }' +
  '}' +

  'function onRedeem(d) {' +
  '  if (d.success) {' +
  '    showBox("valid-box", "🎉 RISCATTATO!\\n" + d.message);' +
  '    document.getElementById("btnRedeem").style.display = "none";' +
  '    document.getElementById("btnStaff").style.display  = "none";' +
  '    var qr = document.getElementById("qrWrap"); if(qr) qr.style.display="none";' +
  '  } else {' +
  '    showBox("error-box", "❌ " + d.message);' +
  '  }' +
  '}' +

  'function onErr(e) { showBox("error-box", "Errore: " + (e.message || e)); }' +

  'function showBox(cls, msg) {' +
  '  var b = document.getElementById("statusBox");' +
  '  b.className = "status-box " + cls;' +
  '  b.style.display = "block";' +
  '  b.style.whiteSpace = "pre-line";' +
  '  b.textContent = msg;' +
  '}' +

  'function showPin()  { document.getElementById("overlay").style.display="flex"; document.getElementById("pinInput").value=""; document.getElementById("pinInput").focus(); }' +
  'function hidePin()  { document.getElementById("overlay").style.display="none"; }' +

  'function checkPin() {' +
  '  var p = document.getElementById("pinInput").value.trim();' +
  '  if (p !== PIN_CORRECT) { document.getElementById("pinInput").value=""; document.getElementById("pinInput").placeholder="PIN errato"; return; }' +
  '  hidePin();' +
  '  document.getElementById("btnRedeem").style.display = "block";' +
  '  document.getElementById("btnStaff").style.display  = "none";' +
  '}' +

  'function doRedeem() {' +
  '  document.getElementById("btnRedeem").style.display = "none";' +
  '  showBox("loading-box", "⏳ Riscatto in corso…");' +
  '  google.script.run.withSuccessHandler(onRedeem).withFailureHandler(onErr).serverRedeem(VOUCHER_CODE);' +
  '}' +
  '<\/script></body></html>';
}
