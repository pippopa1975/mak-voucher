/**
 * Mak Mixology — Voucher QR System
 * =================================
 * Aggiungere questo file al progetto Google Apps Script esistente.
 * Il foglio Google Sheets (ID: 1-0JMPMBtJTJ4WVUz_MnkzHYAkUqWmLmpsmtKTN69464)
 * deve avere un foglio chiamato "Vouchers" con le colonne:
 * A=CodiceVoucher, B=Nome, C=Email, D=Telefono, E=DataCompleanno,
 * F=DataEmissione, G=DataScadenza, H=Stato, I=DataRiscatto
 *
 * NOTA: Aggiornare la funzione doGet() esistente per includere
 * il routing verso le funzioni voucher (vedi commento sotto).
 */

// ===== CONFIGURAZIONE =====
var VOUCHER_SHEET_ID = '1-0JMPMBtJTJ4WVUz_MnkzHYAkUqWmLmpsmtKTN69464';
var VOUCHER_SHEET_NAME = 'Vouchers';
var VOUCHER_PAGE_BASE = 'https://pippopa1975.github.io/mak-voucher/';

// ===== AGGIORNAMENTO doGet =====
// Aggiungere QUESTO BLOCCO all'inizio della funzione doGet(e) esistente,
// PRIMA del codice delle prenotazioni:
//
// function doGet(e) {
//   var action = (e.parameter.action || '').toString();
//   if (action === 'check' || action === 'redeem') {
//     return handleVoucherRequest(e);
//   }
//   // ... codice prenotazioni esistente sotto ...
// }

/**
 * Gestisce le richieste voucher (check e redeem) via JSONP.
 */
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

  var jsonpResponse = callback + '(' + JSON.stringify(result) + ')';
  return ContentService.createTextOutput(jsonpResponse)
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}

/**
 * Verifica lo stato di un voucher.
 */
function checkVoucher(code) {
  var sheet = getVoucherSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === code) {
      var stato = data[i][7].toString().trim();
      var scadenza = data[i][6];
      var scadenzaStr = formatDateISO(scadenza);
      var oggi = new Date();
      oggi.setHours(0, 0, 0, 0);

      if (stato === 'Riscattato') {
        return {
          status: 'redeemed',
          nome: data[i][1].toString(),
          scadenza: scadenzaStr,
          dataRiscatto: formatDateISO(data[i][8])
        };
      }

      var scadenzaDate = new Date(scadenza);
      scadenzaDate.setHours(23, 59, 59, 999);
      if (scadenzaDate < oggi) {
        // Aggiorna stato nel foglio
        sheet.getRange(i + 1, 8).setValue('Scaduto');
        return {
          status: 'expired',
          nome: data[i][1].toString(),
          scadenza: scadenzaStr
        };
      }

      return {
        status: 'valid',
        nome: data[i][1].toString(),
        scadenza: scadenzaStr
      };
    }
  }

  return { status: 'not_found', message: 'Voucher non trovato' };
}

/**
 * Riscatta un voucher (segna come usato).
 */
function redeemVoucher(code) {
  var sheet = getVoucherSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 1; i < data.length; i++) {
    if (data[i][0].toString().trim().toUpperCase() === code) {
      var stato = data[i][7].toString().trim();

      if (stato === 'Riscattato') {
        return { success: false, message: 'Voucher già riscattato.' };
      }

      var scadenza = new Date(data[i][6]);
      scadenza.setHours(23, 59, 59, 999);
      var oggi = new Date();
      oggi.setHours(0, 0, 0, 0);

      if (scadenza < oggi) {
        sheet.getRange(i + 1, 8).setValue('Scaduto');
        return { success: false, message: 'Voucher scaduto.' };
      }

      // Segna come riscattato
      var now = new Date();
      sheet.getRange(i + 1, 8).setValue('Riscattato');
      sheet.getRange(i + 1, 9).setValue(formatDateISO(now));
      return { success: true, message: 'Voucher riscattato con successo!' };
    }
  }

  return { success: false, message: 'Voucher non trovato.' };
}

/**
 * Genera un nuovo voucher per un cliente (da usare nel flusso compleanno).
 * Ritorna: { codice, qrUrl, scadenza }
 */
function generateVoucher(nome, email, telefono, dataCompleanno) {
  var sheet = getVoucherSheet();

  // Genera codice univoco: MAK-[NOME4]-[DDMM]-[RANDOM4]
  var nomePart = nome.replace(/[^A-Za-z]/g, '').toUpperCase().substring(0, 4);
  while (nomePart.length < 4) nomePart += 'X';

  var birthday = new Date(dataCompleanno);
  var dd = ('0' + birthday.getDate()).slice(-2);
  var mm = ('0' + (birthday.getMonth() + 1)).slice(-2);
  var datePart = dd + mm;

  var chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  var randomPart = '';
  for (var i = 0; i < 4; i++) {
    randomPart += chars.charAt(Math.floor(Math.random() * chars.length));
  }

  var codice = 'MAK-' + nomePart + '-' + datePart + '-' + randomPart;

  // Verifica unicità
  var data = sheet.getDataRange().getValues();
  var exists = false;
  for (var j = 1; j < data.length; j++) {
    if (data[j][0].toString().trim().toUpperCase() === codice) {
      exists = true;
      break;
    }
  }
  if (exists) {
    // Rigenera con diverso random
    randomPart = '';
    for (var k = 0; k < 4; k++) {
      randomPart += chars.charAt(Math.floor(Math.random() * chars.length));
    }
    codice = 'MAK-' + nomePart + '-' + datePart + '-' + randomPart;
  }

  // Calcola scadenza: compleanno + 10 giorni
  var scadenza = new Date(birthday.getTime());
  scadenza.setDate(scadenza.getDate() + 10);

  var oggi = new Date();

  // Scrivi riga nel foglio
  sheet.appendRow([
    codice,
    nome,
    email,
    telefono,
    formatDateISO(birthday),
    formatDateISO(oggi),
    formatDateISO(scadenza),
    'Valido',
    ''
  ]);

  // URL per la pagina di validazione
  var validationUrl = VOUCHER_PAGE_BASE + '?v=' + codice;

  // URL per il QR code
  var qrUrl = 'https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=' + encodeURIComponent(validationUrl);

  return {
    codice: codice,
    qrUrl: qrUrl,
    scadenza: formatDateISO(scadenza),
    validationUrl: validationUrl
  };
}

/**
 * Restituisce il foglio "Vouchers", creandolo se non esiste.
 */
function getVoucherSheet() {
  var ss = SpreadsheetApp.openById(VOUCHER_SHEET_ID);
  var sheet = ss.getSheetByName(VOUCHER_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(VOUCHER_SHEET_NAME);
    sheet.appendRow([
      'CodiceVoucher', 'Nome', 'Email', 'Telefono',
      'DataCompleanno', 'DataEmissione', 'DataScadenza',
      'Stato', 'DataRiscatto'
    ]);
    // Formattazione header
    var headerRange = sheet.getRange(1, 1, 1, 9);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#c9a84c');
    headerRange.setFontColor('#000000');
    // Larghezza colonne
    sheet.setColumnWidth(1, 200);
    sheet.setColumnWidth(2, 160);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 140);
    sheet.setColumnWidth(5, 120);
    sheet.setColumnWidth(6, 120);
    sheet.setColumnWidth(7, 120);
    sheet.setColumnWidth(8, 100);
    sheet.setColumnWidth(9, 120);
  }
  return sheet;
}

/**
 * Formatta una data come YYYY-MM-DD.
 */
function formatDateISO(date) {
  if (!date) return '';
  if (typeof date === 'string') {
    if (/^\d{4}-\d{2}-\d{2}/.test(date)) return date.substring(0, 10);
    date = new Date(date);
  }
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  var y = date.getFullYear();
  var m = ('0' + (date.getMonth() + 1)).slice(-2);
  var d = ('0' + date.getDate()).slice(-2);
  return y + '-' + m + '-' + d;
}

/**
 * Funzione di test: genera un voucher di prova.
 * Eseguire manualmente da Apps Script per verificare il setup.
 */
function testGenerateVoucher() {
  var result = generateVoucher(
    'Mario Rossi',
    'mario@test.com',
    '+393331234567',
    '2026-09-14'
  );
  Logger.log(JSON.stringify(result));
}
