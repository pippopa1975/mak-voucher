var SHEET_ID = '1-0JMPMBtJTJ4WVUz_MnkzHYAkUqWmLmpsmtKTN69464';
var SHEET_NAME = 'Vouchers';

function doGet(e) {
  var action = (e.parameter.action || '').toString().trim();
  var code   = (e.parameter.code || '').toString().trim().toUpperCase();
  var vCode  = (e.parameter.v || '').toString().trim().toUpperCase();

  var cb = e.parameter.callback || 'callback';
  if (action === 'generate') {
    var nome = e.parameter.nome || '';
    var email = e.parameter.email || '';
    var telefono = e.parameter.telefono || '';
    var compleanno = e.parameter.compleanno || '';
    if (!nome || !compleanno) {
      var r = ContentService.createTextOutput(cb+'({"status":"error","message":"Parametri mancanti"})');
      r.setMimeType(ContentService.MimeType.JAVASCRIPT);
      return r;
    }
    var result = generateVoucher(nome, email, telefono, compleanno);
    var out = ContentService.createTextOutput(cb+'('+JSON.stringify(result)+')');
    out.setMimeType(ContentService.MimeType.JAVASCRIPT);
    return out;
  }
  if (action === 'check' || action === 'redeem') {
    var result = (action === 'check') ? checkVoucher(code) : redeemVoucher(code);
    var cb = e.parameter.callback || '';
    var out = cb ? cb + '(' + JSON.stringify(result) + ')' : JSON.stringify(result);
    var mime = cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
    return ContentService.createTextOutput(out).setMimeType(mime);
  }

  // Serve the HTML validation page
  var html = buildHtml(vCode);
  return HtmlService.createHtmlOutput(html)
    .setTitle('Mak Mixology — Voucher')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Called from the HTML page via google.script.run
function serverCheck(code) {
  return checkVoucher(code.toString().trim().toUpperCase());
}
function serverRedeem(code) {
  return redeemVoucher(code.toString().trim().toUpperCase());
}

function buildHtml(preloadCode) {
  return '<!DOCTYPE html><html lang="it"><head><meta charset="UTF-8">' +
  '<meta name="viewport" content="width=device-width,initial-scale=1">' +
  '<title>Mak Mixology — Voucher</title>' +
  '<style>' +
  'body{margin:0;background:#0a0a0a;font-family:sans-serif;color:#fff;display:flex;flex-direction:column;align-items:center;justify-content:flex-start;min-height:100vh;padding:24px 16px}' +
  '.logo{font-size:22px;font-weight:700;letter-spacing:3px;color:#c9a84c;margin-bottom:32px}' +
  '.card{background:#111;border:1px solid #c9a84c33;border-radius:16px;padding:28px 24px;max-width:380px;width:100%;text-align:center}' +
  'input{width:100%;box-sizing:border-box;background:#1a1a1a;border:1px solid #444;color:#fff;font-size:18px;letter-spacing:2px;text-transform:uppercase;padding:14px;border-radius:8px;margin-bottom:16px;text-align:center}' +
  '.btn{width:100%;padding:14px;border-radius:8px;border:none;font-size:16px;font-weight:700;cursor:pointer;margin-top:8px}' +
  '.btn-check{background:#c9a84c;color:#000}' +
  '.btn-redeem{background:#2d7a2d;color:#fff}' +
  '.status{margin-top:20px;padding:16px;border-radius:8px;font-size:15px;display:none}' +
  '.valid{background:#1a3a1a;border:1px solid #2d7a2d;color:#7fe87f}' +
  '.invalid{background:#3a1a1a;border:1px solid #7a2d2d;color:#e87f7f}' +
  '.warning{background:#3a2d1a;border:1px solid #c9a84c;color:#c9a84c}' +
  '.loading{opacity:.5;pointer-events:none}' +
  '</style></head><body>' +
  '<div class="logo">MAK MIXOLOGY</div>' +
  '<div class="card">' +
  '<p style="color:#888;font-size:13px;margin-top:0">Inserisci o scansiona il codice voucher</p>' +
  '<input type="text" id="code" placeholder="MAK-XXXX-0000-XXXX" maxlength="20" oninput="this.value=this.value.toUpperCase()" />' +
  '<button class="btn btn-check" onclick="doCheck()">VERIFICA VOUCHER</button>' +
  '<div class="status" id="status"></div>' +
  '<button class="btn btn-redeem" id="btnRedeem" style="display:none" onclick="doRedeem()">✅ RISCATTA VOUCHER</button>' +
  '</div>' +
  '<script>' +
  'var currentCode = "' + preloadCode + '";' +
  'if(currentCode){document.getElementById("code").value=currentCode;setTimeout(doCheck,400);}' +
  'function doCheck(){' +
  '  var c=document.getElementById("code").value.trim().toUpperCase();' +
  '  if(!c){showStatus("Inserisci un codice","invalid");return;}' +
  '  currentCode=c;showStatus("Verifica in corso…","warning");' +
  '  document.getElementById("btnRedeem").style.display="none";' +
  '  google.script.run.withSuccessHandler(onCheck).withFailureHandler(onErr).serverCheck(c);' +
  '}' +
  'function doRedeem(){' +
  '  if(!currentCode)return;' +
  '  showStatus("Riscatto in corso…","warning");' +
  '  document.getElementById("btnRedeem").style.display="none";' +
  '  google.script.run.withSuccessHandler(onRedeem).withFailureHandler(onErr).serverRedeem(currentCode);' +
  '}' +
  'function onCheck(d){' +
  '  if(d.status==="valid"){showStatus("✅ VOUCHER VALIDO\\nNome: "+d.nome+"\\nScadenza: "+d.scadenza,"valid");document.getElementById("btnRedeem").style.display="block";}' +
  '  else if(d.status==="redeemed"){showStatus("⚠️ GIÀ RISCATTATO\\nNome: "+d.nome+"\\nData riscatto: "+d.dataRiscatto,"warning");}' +
  '  else if(d.status==="expired"){showStatus("❌ VOUCHER SCADUTO\\nScadenza: "+d.scadenza,"invalid");}' +
  '  else{showStatus("❌ VOUCHER NON TROVATO","invalid");}' +
  '}' +
  'function onRedeem(d){' +
  '  if(d.success){showStatus("🎉 RISCATTATO CON SUCCESSO!\\n"+d.message,"valid");}' +
  '  else{showStatus("❌ "+d.message,"invalid");}' +
  '}' +
  'function onErr(e){showStatus("Errore: "+e.message,"invalid");}' +
  'function showStatus(msg,cls){' +
  '  var el=document.getElementById("status");' +
  '  el.textContent=msg;el.className="status "+cls;el.style.display="block";' +
  '  el.style.whiteSpace="pre-line";' +
  '}' +
  '<\/script></body></html>';
}

function checkVoucher(code) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === code) {
      var stato = data[i][7].toString().trim();
      var scadenza = new Date(data[i][6]);
      var oggi = new Date(); oggi.setHours(0,0,0,0);
      scadenza.setHours(23,59,59,999);
      if (stato === 'Riscattato') {
        return { status: 'redeemed', nome: data[i][1], scadenza: fmtIT(new Date(data[i][6])), dataRiscatto: fmtIT(new Date(data[i][8])) };
      }
      if (scadenza < oggi) {
        sheet.getRange(i+1,8).setValue('Scaduto');
        return { status: 'expired', nome: data[i][1], scadenza: fmtIT(new Date(data[i][6])) };
      }
      return { status: 'valid', nome: data[i][1], scadenza: fmtIT(new Date(data[i][6])) };
    }
  }
  return { status: 'not_found' };
}

function redeemVoucher(code) {
  var sheet = getSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === code) {
      var stato = data[i][7].toString().trim();
      if (stato === 'Riscattato') return { success: false, message: 'Gia riscattato.' };
      var scadenza = new Date(data[i][6]); scadenza.setHours(23,59,59,999);
      if (scadenza < new Date()) { sheet.getRange(i+1,8).setValue('Scaduto'); return { success: false, message: 'Voucher scaduto.' }; }
      sheet.getRange(i+1,8).setValue('Riscattato');
      sheet.getRange(i+1,9).setValue(fmtIT(new Date()));
      return { success: true, message: 'Drink omaggio pronto!' };
    }
  }
  return { success: false, message: 'Voucher non trovato.' };
}

// NOTA: il parametro "compleanno" deve essere in formato ISO YYYY-MM-DD (es. "1975-09-14").
// new Date(compleanno) richiede questo formato per parsare correttamente la data.
function generateVoucher(nome, email, telefono, compleanno) {
  var sheet = getSheet();
  var nomePart = nome.replace(/[^A-Za-z]/g,'').toUpperCase().substring(0,4);
  while (nomePart.length < 4) nomePart += 'X';
  var bd = new Date(compleanno);
  var dd = ('0'+bd.getDate()).slice(-2);
  var mm = ('0'+(bd.getMonth()+1)).slice(-2);
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var rnd = ''; for (var i=0;i<4;i++) rnd += chars.charAt(Math.floor(Math.random()*chars.length));
  var codice = 'MAK-'+nomePart+'-'+dd+mm+'-'+rnd;
  // Scadenza: 10 giorni dall'anniversario dell'anno CORRENTE (non dall'anno di nascita)
  var oggi = new Date();
  var bdQuest = new Date(oggi.getFullYear(), bd.getMonth(), bd.getDate());
  if (bdQuest < oggi) bdQuest.setFullYear(oggi.getFullYear() + 1);
  var scadenza = new Date(bdQuest.getTime());
  scadenza.setDate(scadenza.getDate() + 10);
  sheet.appendRow([codice, nome, email, telefono, fmtIT(bd), fmtIT(new Date()), fmtIT(scadenza), 'Valido', '']);
  // QR points to the script itself (no CORS issue)
  var validationUrl = 'https://pippopa1975.github.io/mak-voucher/?v=' + codice;
  var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' + encodeURIComponent(validationUrl);
  return { codice: codice, qrUrl: qrUrl, scadenza: fmtIT(scadenza), validationUrl: validationUrl };
}

function getSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var s = ss.getSheetByName(SHEET_NAME);
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

function testCheck() { Logger.log(JSON.stringify(checkVoucher('MAK-MARI-1409-VQLR'))); }

function testGenerate() { Logger.log(JSON.stringify(generateVoucher('Mario Rossi','test@test.com','3331234567','1990-03-15'))); }
