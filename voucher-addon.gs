// ===== VOUCHER QR — MAK MIXOLOGY =====
var VOUCHER_SHEET_ID = '1-0JMPMBtJTJ4WVUz_MnkzHYAkUqWmLmpsmtKTN69464';
var VOUCHER_SHEET_NAME = 'Vouchers';
var VOUCHER_PAGE_BASE = 'https://pippopa1975.github.io/mak-voucher/';

function handleVoucherRequest(e) {
  var action = e.parameter.action;
  var code = (e.parameter.code || '').toString().trim().toUpperCase();
  var callback = e.parameter.callback || 'callback';
  var result;
  if (!code) {
    result = { status: 'not_found', message: 'Codice mancante' };
  } else if (action === 'check') {
    result = checkVoucher(code);
  } else if (action === 'redeem') {
    result = redeemVoucher(code);
  } else {
    result = { error: true, message: 'Azione non valida' };
  }
  return ContentService.createTextOutput(callback + '(' + JSON.stringify(result) + ')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function checkVoucher(code) {
  var sheet = getVoucherSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === code) {
      var stato = data[i][7].toString().trim();
      var scadenza = new Date(data[i][6]);
      var scadenzaStr = formatDateIT(scadenza);
      var oggi = new Date(); oggi.setHours(0,0,0,0);
      if (stato === 'Riscattato') {
        return { status: 'redeemed', nome: data[i][1].toString(), scadenza: scadenzaStr, dataRiscatto: formatDateIT(new Date(data[i][8])) };
      }
      scadenza.setHours(23,59,59,999);
      if (scadenza < oggi) {
        sheet.getRange(i+1, 8).setValue('Scaduto');
        return { status: 'expired', nome: data[i][1].toString(), scadenza: scadenzaStr };
      }
      return { status: 'valid', nome: data[i][1].toString(), scadenza: scadenzaStr };
    }
  }
  return { status: 'not_found', message: 'Voucher non trovato' };
}

function redeemVoucher(code) {
  var sheet = getVoucherSheet();
  var data = sheet.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === code) {
      var stato = data[i][7].toString().trim();
      if (stato === 'Riscattato') return { success: false, message: 'Voucher gia riscattato.' };
      var scadenza = new Date(data[i][6]); scadenza.setHours(23,59,59,999);
      var oggi = new Date(); oggi.setHours(0,0,0,0);
      if (scadenza < oggi) { sheet.getRange(i+1,8).setValue('Scaduto'); return { success: false, message: 'Voucher scaduto.' }; }
      sheet.getRange(i+1,8).setValue('Riscattato');
      sheet.getRange(i+1,9).setValue(formatDateIT(new Date()));
      return { success: true, message: 'Voucher riscattato!' };
    }
  }
  return { success: false, message: 'Voucher non trovato.' };
}

function generateVoucher(nome, email, telefono, dataCompleanno) {
  var sheet = getVoucherSheet();
  var nomePart = nome.replace(/[^A-Za-z]/g,'').toUpperCase().substring(0,4);
  while (nomePart.length < 4) nomePart += 'X';
  var bd = new Date(dataCompleanno);
  var dd = ('0'+bd.getDate()).slice(-2);
  var mm = ('0'+(bd.getMonth()+1)).slice(-2);
  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var rnd = ''; for (var i=0;i<4;i++) rnd += chars.charAt(Math.floor(Math.random()*chars.length));
  var codice = 'MAK-'+nomePart+'-'+dd+mm+'-'+rnd;
  var scadenza = new Date(bd.getTime()); scadenza.setDate(scadenza.getDate()+10);
  var oggi = new Date();
  sheet.appendRow([codice, nome, email, telefono, formatDateIT(bd), formatDateIT(oggi), formatDateIT(scadenza), 'Valido', '']);
  var validationUrl = VOUCHER_PAGE_BASE+'?v='+codice;
  var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data='+encodeURIComponent(validationUrl);
  return { codice: codice, qrUrl: qrUrl, scadenza: formatDateIT(scadenza), validationUrl: validationUrl };
}

function getVoucherSheet() {
  var ss = SpreadsheetApp.openById(VOUCHER_SHEET_ID);
  var sheet = ss.getSheetByName(VOUCHER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(VOUCHER_SHEET_NAME);
    sheet.appendRow(['CodiceVoucher','Nome','Email','Telefono','DataCompleanno','DataEmissione','DataScadenza','Stato','DataRiscatto']);
    sheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#c9a84c');
  }
  return sheet;
}

function formatDateIT(date) {
  if (!date) return '';
  var d = (date instanceof Date) ? date : new Date(date);
  if (isNaN(d.getTime())) return '';
  return ('0'+d.getDate()).slice(-2)+'/'+('0'+(d.getMonth()+1)).slice(-2)+'/'+d.getFullYear();
}
