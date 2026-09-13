# 🎟️ Mak Mixology — Sistema Voucher QR

Sistema di voucher con QR code per drink omaggio di compleanno.

## Come funziona

1. **Generazione**: Quando invii un messaggio di buon compleanno a un cliente, il sistema genera un voucher con QR code univoco
2. **Consegna**: Il QR code viene inviato al cliente (WhatsApp/email)
3. **Riscatto**: Il cliente mostra il QR allo staff → lo staff scansiona → vede lo stato → riscatta con un tap

## 📋 Setup Google Apps Script

### Passo 1: Apri il progetto esistente

Vai a: [Google Apps Script](https://script.google.com) e apri il progetto collegato al form prenotazioni.

> URL attuale del web app: `https://script.google.com/macros/s/AKfycbw1Dv8C9hXfXC5NdYe3--LmZTpDuER2PYstaOxqiotXf-t5kWoECfMxZzvX1nQQLI-vcA/exec`

### Passo 2: Aggiungi il file voucher

1. Nel menu a sinistra, clicca il **+** accanto a "File"
2. Seleziona **Script**
3. Rinomina il file in `voucher` (diventerà `voucher.gs`)
4. Copia e incolla tutto il contenuto del file `apps-script-voucher.gs`
5. Salva (Ctrl+S)

### Passo 3: Aggiorna la funzione doGet

Apri il file principale (quello con `doGet`) e **aggiungi queste righe all'inizio** della funzione `doGet(e)`:

```javascript
function doGet(e) {
  // === VOUCHER ROUTING (aggiunta) ===
  var action = (e.parameter.action || '').toString();
  if (action === 'check' || action === 'redeem') {
    return handleVoucherRequest(e);
  }
  // === FINE VOUCHER ROUTING ===

  // ... tutto il codice prenotazioni esistente resta sotto, invariato ...
}
```

### Passo 4: Ridistribuisci

1. Vai su **Distribuzione** → **Gestisci distribuzioni**
2. Clicca sull'icona ✏️ (modifica) sulla distribuzione esistente
3. In **Versione** seleziona **Nuova versione**
4. Clicca **Distribuzione**
5. L'URL rimane lo stesso!

### Passo 5: Testa

1. Esegui la funzione `testGenerateVoucher` dall'editor (menu Esegui → testGenerateVoucher)
2. Controlla il foglio "Vouchers" — dovrebbe apparire una riga di test
3. Copia il codice generato e apri: `https://pippopa1975.github.io/mak-voucher/?v=CODICE`
4. Verifica che lo stato appaia come "Valido"

## 📊 Foglio Google Sheets

Il foglio **Vouchers** viene creato automaticamente nello stesso Google Sheet delle prenotazioni con queste colonne:

| Colonna | Contenuto |
|---------|-----------|
| A | CodiceVoucher (es. MAK-MARC-1409-7X3K) |
| B | Nome cliente |
| C | Email |
| D | Telefono |
| E | Data compleanno |
| F | Data emissione |
| G | Data scadenza (compleanno +10 giorni) |
| H | Stato (Valido / Riscattato / Scaduto) |
| I | Data riscatto |

## 🔗 Formato QR Code

Ogni voucher genera un QR code via API gratuita:
```
https://api.qrserver.com/v1/create-qr-code/?size=300x300&data=https://pippopa1975.github.io/mak-voucher/?v=MAK-XXXX-DDMM-RAND
```

## 🎨 Pagina Staff (GitHub Pages)

URL: `https://pippopa1975.github.io/mak-voucher/`

La pagina mostra:
- ✅ **VALIDO** — Pulsante verde per riscattare
- 🔴 **GIÀ RISCATTATO** — Data del riscatto
- ❌ **SCADUTO** — Data di scadenza
- 🚫 **NON VALIDO** — Codice non trovato

---

*Mak Mixology — Galleria delle Vittorie, Palermo*
