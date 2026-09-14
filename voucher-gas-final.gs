var SHEET_ID = '1-0JMPMBtJTJ4WVUz_MnkzHYAkUqWmLmpsmtKTN69464';
var SHEET_NAME = 'Vouchers';
var REPORT_SHEET_ID = '1txRlZ0uuUg1vW5CwqTZ47x2CuVzADlVxIVbzq-vscW4';
var REPORT_SHEET_NAME = 'Report Voucher';

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
    var cbVal = e.parameter.callback || '';
    var out = cbVal ? cbVal + '(' + JSON.stringify(result) + ')' : JSON.stringify(result);
    var mime = cbVal ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
    return ContentService.createTextOutput(out).setMimeType(mime);
  }

  var html = buildHtml(vCode);
  return HtmlService.createHtmlOutput(html)
    .setTitle('Mak Mixology — Voucher')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function serverCheck(code) {
  return checkVoucher(code.toString().trim().toUpperCase());
}
function serverRedeem(code) {
  return redeemVoucher(code.toString().trim().toUpperCase());
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
      // === REPORT VOUCHER: aggiorna data utilizzo e stato ===
      logRedeemToReport(code, fmtIT(new Date()));
      return { success: true, message: 'Drink omaggio pronto!' };
    }
  }
  return { success: false, message: 'Voucher non trovato.' };
}

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
  var oggi = new Date();
  var bdQuest = new Date(oggi.getFullYear(), bd.getMonth(), bd.getDate());
  if (bdQuest < oggi) bdQuest.setFullYear(oggi.getFullYear() + 1);
  var scadenza = new Date(bdQuest.getTime());
  scadenza.setDate(scadenza.getDate() + 10);
  sheet.appendRow([codice, nome, email, telefono, fmtIT(bd), fmtIT(new Date()), fmtIT(scadenza), 'Valido', '']);
  // === REPORT VOUCHER: logga l'emissione ===
  logEmitToReport(codice, nome, fmtIT(new Date()), fmtIT(scadenza));
  var validationUrl = 'https://pippopa1975.github.io/mak-voucher/?v=' + codice;
  var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' + encodeURIComponent(validationUrl);
  return { codice: codice, qrUrl: qrUrl, scadenza: fmtIT(scadenza), validationUrl: validationUrl };
}

// === FUNZIONI REPORT VOUCHER ===
function getReportSheet() {
  try {
    var ss = SpreadsheetApp.openById(REPORT_SHEET_ID);
    return ss.getSheetByName(REPORT_SHEET_NAME);
  } catch(e) { return null; }
}

function logEmitToReport(codice, nome, dataEmissione, scadenza) {
  var rs = getReportSheet();
  if (!rs) return;
  rs.appendRow([codice, nome, dataEmissione, scadenza, '', 'Emesso']);
}

function logRedeemToReport(code, dataUtilizzo) {
  var rs = getReportSheet();
  if (!rs) return;
  var data = rs.getDataRange().getValues();
  for (var j = 1; j < data.length; j++) {
    if (data[j][0].toString().trim().toUpperCase() === code) {
      rs.getRange(j+1, 5).setValue(dataUtilizzo); // Data Utilizzo
      rs.getRange(j+1, 6).setValue('Riscattato');  // Stato
      return;
    }
  }
  // Se non trovato (voucher emessi prima dell'aggiornamento), aggiunge riga
  rs.appendRow([code, '', '', '', dataUtilizzo, 'Riscattato']);
}

// === UTILS ===
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

function buildHtml(preloadCode) {
  return '<!DOCTYPE html><html lang="it"><head><meta charset="UTF-8">' +
  '<meta name="viewport" content="width=device-width,initial-scale=1">' +
  '<title>Mak Mixology — Voucher</title></head><body>' +
  '<p>Usa il link GitHub Pages per la validazione.</p>' +
  '</body></html>';
}

function testCheck() { Logger.log(JSON.stringify(checkVoucher('MAK-MARI-1409-VQLR'))); }
function testGenerate() { Logger.log(JSON.stringify(generateVoucher('Mario Rossi','test@test.com','3331234567','1990-03-15'))); }
