/**
 * Europa 2026 — Google Apps Script Web App
 *
 * INSTRUCCIONES:
 * 1. Abre tu Google Sheet → Extensiones → Apps Script
 * 2. Borrá el contenido existente y pegá todo este archivo
 * 3. Guardá (Ctrl+S)
 * 4. Clic en "Implementar" → "Nueva implementación"
 *    - Tipo: App web
 *    - Ejecutar como: Yo
 *    - Acceso: Cualquier usuario (incluso anónimo)
 * 5. Copiá la URL generada
 * 6. En index.html, reemplazá 'TU_APPS_SCRIPT_URL_AQUI' con esa URL
 */

function doGet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {};

  ss.getSheets().forEach(function(sheet) {
    const name = sheet.getName();
    const rows = sheet.getDataRange().getDisplayValues();
    result[name] = parseTab(rows);
  });

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function parseTab(rows) {
  // Busca la fila de encabezado (primera fila con "DIA" en col A)
  var headerIdx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i][0]).trim().toUpperCase() === 'DIA') {
      headerIdx = i;
      break;
    }
  }
  if (headerIdx === -1) return [];

  var headers = rows[headerIdx].map(function(h) { return String(h).trim().toUpperCase(); });
  var COL = {
    DIA:         headers.indexOf('DIA'),
    HORA:        headers.indexOf('HORA'),
    ZONA:        headers.indexOf('ZONA'),
    ACTIVIDAD:   headers.indexOf('ACTIVIDAD'),
    TIEMPO:      headers.indexOf('TIEMPO'),
    COMENTARIOS: headers.indexOf('COMENTARIOS'),
    COSTO:       headers.indexOf('$'),
  };

  var daysMap = {};
  var daysOrder = [];

  for (var i = headerIdx + 1; i < rows.length; i++) {
    var r = rows[i];
    var dia = String(r[COL.DIA] !== undefined ? r[COL.DIA] : '').trim();

    // Saltar filas vacías o filas de encabezado repetidas
    if (!dia || dia.toUpperCase() === 'DIA') continue;

    if (!daysMap[dia]) {
      daysMap[dia] = { date: dia, label: dia, items: [] };
      daysOrder.push(dia);
    }

    var zona = String(r[COL.ZONA] !== undefined ? r[COL.ZONA] : '').trim();
    if (!zona) continue;

    var actividad   = String(r[COL.ACTIVIDAD]   !== undefined ? r[COL.ACTIVIDAD]   : '').trim();
    var comentarios = String(r[COL.COMENTARIOS] !== undefined ? r[COL.COMENTARIOS] : '').trim();
    var detail = [actividad, comentarios].filter(Boolean).join(' · ');

    var costoRaw = String(r[COL.COSTO] !== undefined ? r[COL.COSTO] : '').trim();
    var costo = costoRaw.replace(/^[-\\]+$/, '').trim();

    daysMap[dia].items.push({
      type:     zona.toLowerCase() === 'traslado' ? 'transfer' : 'event',
      time:     String(r[COL.HORA] !== undefined ? r[COL.HORA] : '').trim(),
      activity: zona,
      detail:   detail,
      cost:     costo,
      duration: String(r[COL.TIEMPO] !== undefined ? r[COL.TIEMPO] : '').trim(),
    });
  }

  return daysOrder.map(function(dia) { return daysMap[dia]; });
}
