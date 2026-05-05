function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback ? e.parameter.callback : null;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const trades = [];

  function getHour(cell) {
    if (!cell || cell === '') return null;
    try {
      const iso = new Date(cell).toISOString();
      const parts = iso.split('T')[1].split(':');
      const utcH = parseInt(parts[0]);
      const utcM = parseInt(parts[1]);
      return (utcH - 8 + 24) % 24 + utcM / 60;
    } catch(e) { return null; }
  }

  function getHourStr(cell) {
    if (!cell || cell === '') return null;
    try {
      const iso = new Date(cell).toISOString();
      const parts = iso.split('T')[1].split(':');
      const utcH = (parseInt(parts[0]) - 8 + 24) % 24;
      const utcM = parseInt(parts[1]);
      return String(utcH).padStart(2,'0') + ':' + String(utcM).padStart(2,'0');
    } catch(e) { return null; }
  }

  function getDuration(cell) {
    if (!cell || cell === '') return null;
    try {
      const iso = new Date(cell).toISOString();
      const parts = iso.split('T')[1].split(':');
      const utcH = parseInt(parts[0]);
      const utcM = parseInt(parts[1]);
      const totalMin = utcH * 60 + utcM - 480;
      return totalMin > 0 && totalMin < 480 ? totalMin : null;
    } catch(e) { return null; }
  }

  function getUrl(cell) {
    if (!cell || cell === '') return null;
    const s = String(cell).trim();
    if (s === '') return null;
    return s.startsWith('http') ? s : 'https://' + s;
  }

  function normPair(p) {
    if (!p) return '';
    return String(p).trim()
      .replace('EURUSD', 'EUR/USD')
      .replace('GBPUSD', 'GBP/USD')
      .replace('XAUUSD', 'XAU/USD');
  }

  const VALID_SENS = new Set(['Seguro - Confiado','Convencido - Calma','Dudoso - Inseguro','Fomo - Acelerado','Venganza - Rabia','Miedo - Parálisis']);

  function cleanSens(val) {
    const s = String(val || '').trim();
    return VALID_SENS.has(s) ? s : '';
  }

  // ZONAS
  const wsZ = ss.getSheetByName('ZONAS');
  const dataZ = wsZ.getRange(9, 1, wsZ.getLastRow() - 8, 24).getValues();
  dataZ.forEach(row => {
    const res = row[13];
    if (!['TP','SL','BE'].includes(res)) return;
    const pnl = parseFloat(row[12]) || 0;
    const date = row[4];
    if (!date) return;
    trades.push({
      sheet: 'ZONAS',
      date: Utilities.formatDate(new Date(date), 'Europe/Madrid', 'yyyy-MM-dd'),
      result: res, pnl: pnl,
      open_hour: getHour(row[6]),
      open_str: getHourStr(row[6]),
      dur: getDuration(row[8]),
      setup: row[3] || '',
      pair: normPair(row[2]),
      zone: String(row[10] || ''),
      entry: 'Stop Limit',
      sensacion: cleanSens(row[21]),
      url1: getUrl(row[22]),
      url2: null,
      reflexion: String(row[23] || '')
    });
  });

  // LIQUIDEZ
  const wsL = ss.getSheetByName('LIQUIDEZ');
  const dataL = wsL.getRange(7, 1, wsL.getLastRow() - 6, 27).getValues();
  dataL.forEach(row => {
    const res = row[15];
    if (!['TP','SL','BE'].includes(res)) return;
    const pnl = parseFloat(row[14]) || 0;
    const date = row[4];
    if (!date) return;
    const u1 = getUrl(row[24]);
    const u2 = getUrl(row[25]);
    trades.push({
      sheet: 'LIQUIDEZ',
      date: Utilities.formatDate(new Date(date), 'Europe/Madrid', 'yyyy-MM-dd'),
      result: res, pnl: pnl,
      open_hour: getHour(row[6]),
      open_str: getHourStr(row[6]),
      dur: getDuration(row[8]),
      setup: row[3] || '',
      pair: normPair(row[2]),
      zone: String(row[9] || ''),
      entry: String(row[12] || ''),
      sensacion: cleanSens(row[23]),
      url1: u1 || u2,
      url2: u1 && u2 ? u2 : null,
      reflexion: String(row[26] || '')
    });
  });

  // NASDAQ
  const wsN = ss.getSheetByName('NASDAQ');
  const dataN = wsN.getRange(9, 1, wsN.getLastRow() - 8, 26).getValues();
  dataN.forEach(row => {
    const res = row[14];
    if (!['TP','SL','BE'].includes(res)) return;
    const pnl = parseFloat(row[13]) || 0;
    const date = row[3];
    if (!date) return;
    const u1 = getUrl(row[23]);
    const u2 = getUrl(row[24]);
    trades.push({
      sheet: 'NASDAQ',
      date: Utilities.formatDate(new Date(date), 'Europe/Madrid', 'yyyy-MM-dd'),
      result: res, pnl: pnl,
      open_hour: getHour(row[5]),
      open_str: getHourStr(row[5]),
      dur: getDuration(row[7]),
      setup: row[2] || '',
      pair: 'NQ',
      zone: String(row[8] || ''),
      entry: String(row[11] || ''),
      sensacion: cleanSens(row[22]),
      url1: u1 || u2,
      url2: u1 && u2 ? u2 : null,
      reflexion: String(row[25] || '')
    });
  });

  const json = JSON.stringify({ trades });

  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Tradinverso')
    .addItem('Abrir Reflexión', 'abrirReflexion')
    .addToUi();
}

function abrirReflexion() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const sheetName = sheet.getName();
  const reflexionCols = { 'ZONAS': 24, 'LIQUIDEZ': 27, 'NASDAQ': 26 };
  const col = reflexionCols[sheetName];
  if (!col) {
    SpreadsheetApp.getUi().alert('Solo disponible en ZONAS, LIQUIDEZ y NASDAQ.');
    return;
  }
  const row = cell.getRow();
  if (row < 7) {
    SpreadsheetApp.getUi().alert('Selecciona una fila de trade.');
    return;
  }
  const currentText = sheet.getRange(row, col).getValue() || '';
  const tradeNum = sheet.getRange(row, 2).getValue() || row;
  const html = HtmlService.createHtmlOutput(`
    <!DOCTYPE html><html><head>
    <style>
      *{box-sizing:border-box;margin:0;padding:0;}
      body{font-family:'Google Sans',Arial,sans-serif;background:#0F1420;color:#D0D4E8;
           padding:16px;height:100vh;display:flex;flex-direction:column;gap:12px;}
      .header{font-size:13px;font-weight:600;color:#F0F2F8;padding-bottom:10px;
              border-bottom:1px solid rgba(255,255,255,0.1);}
      .sub{font-size:11px;color:#8A96B8;font-family:monospace;}
      textarea{flex:1;width:100%;background:#1A2235;color:#E0E4F0;
               border:1px solid rgba(255,255,255,0.15);border-radius:8px;
               padding:12px;font-size:13px;line-height:1.6;resize:none;outline:none;}
      textarea:focus{border-color:rgba(91,141,239,0.5);}
      .btns{display:flex;gap:8px;justify-content:flex-end;}
      button{padding:8px 20px;border-radius:6px;font-size:12px;font-weight:500;cursor:pointer;border:none;}
      .cancel{background:rgba(255,255,255,0.08);color:#8A96B8;}
      .save{background:#00D4AA;color:#000;}
    </style></head><body>
    <div class="header">Reflexión — Trade ${tradeNum}</div>
    <div class="sub">${sheetName} · Fila ${row}</div>
    <textarea id="txt" placeholder="Escribe tu reflexión aquí...">${currentText}</textarea>
    <div class="btns">
      <button class="cancel" onclick="google.script.host.close()">Cancelar</button>
      <button class="save" onclick="guardar()">Guardar</button>
    </div>
    <script>
      document.getElementById('txt').focus();
      function guardar(){
        google.script.run
          .withSuccessHandler(()=>google.script.host.close())
          .guardarReflexion(document.getElementById('txt').value);
      }
    <\/script>
    </body></html>
  `).setWidth(520).setHeight(420).setTitle('Reflexión del trade');
  SpreadsheetApp.getUi().showModalDialog(html, 'Reflexión');
}

function guardarReflexion(texto) {
  const sheet = SpreadsheetApp.getActiveSheet();
  const cell = sheet.getActiveCell();
  const sheetName = sheet.getName();
  const reflexionCols = { 'ZONAS': 24, 'LIQUIDEZ': 27, 'NASDAQ': 26 };
  const col = reflexionCols[sheetName];
  if (!col) return;
  sheet.getRange(cell.getRow(), col).setValue(texto);
}
