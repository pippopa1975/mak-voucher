var SHEET_ID = '1-0JMPMBtJTJ4WVUz_MnkzHYAkUqWmLmpsmtKTN69464';
var SHEET_NAME = 'Vouchers';

function doGet(e) {
  var action = (e.parameter.action || '').toString().trim();
  var code = (e.parameter.code || '').toString().trim().toUpperCase();
  var callback = e.parameter.callback || 'callback';
  var result;

  if (action === 'check' && code) {
    result = checkVoucher(code);
  } else if (action === 'redeem' && code) {
    result = redeemVoucher(code);
  } else if (action === 'generate') {
    var nome = e.parameter.nome || '';
    var email = e.parameter.email || '';
    var telefono = e.parameter.telefono || '';
    var compleanno = e.parameter.compleanno || '';
    result = generateVoucher(nome, email, telefono, compleanno);
  } else {
    result = { error: true, message: 'Parametri mancanti' };
  }

  return ContentService
    .createTextOutput(callback + '(' + JSON.stringify(result) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
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
      return { success: true, message: 'Voucher riscattato con successo!' };
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
  var scadenza = new Date(bd.getTime()); scadenza.setDate(scadenza.getDate()+10);
  sheet.appendRow([codice, nome, email, telefono, fmtIT(bd), fmtIT(new Date()), fmtIT(scadenza), 'Valido', '']);
  var pageBase = 'https://pippopa1975.github.io/mak-voucher/';
  var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data='+encodeURIComponent(pageBase+'?v='+codice);
  return { codice: codice, qrUrl: qrUrl, scadenza: fmtIT(scadenza) };
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

function testCheck() {
  Logger.log(JSON.stringify(checkVoucher('TEST')));
}
