/***************************************************************
 * API Builder + Listener (Bonos Dashboard)
 * Sheet: 01 DashBoard  ->  API (solo lectura para la web)
 *
 * Endpoints:
 *  GET  ?action=setup_api
 *  GET  ?action=refresh_api
 *  GET  ?action=install_triggers&minutes=5
 *  GET  ?action=remove_triggers
 *
 * (Opcional) POST { action: "setup_api" | "refresh_api" | ... }
 ***************************************************************/

const SPREADSHEET_ID = "1bJHM84cRxKQRJjcikVydCQeMgxVWoZmi34M4sGD-jVw";

const SHEET_DASH = "01 DashBoard";
const SHEET_API  = "API";

// Layout API (según tu captura)
const API_LAYOUT = {
  // KPI Key/Value
  KPI_KEYS_RANGE: "A1:B14", // headers A(Key), B(Value) + keys A2:A14

  // Bloques visibles
  SIN_ASIGNAR_TOPLEFT: "C1",  // C:E (3 cols)
  PENDIENTES_TOPLEFT:  "H1",  // H:K (4 cols)
  PLAN_TOPLEFT:        "M1",  // M:N (2 cols)

  // RECOMENDADO: hacer PREPARAR contiguo O:R (4 cols)
  // (Fecha, Distribuidor, $ Min, $ Max)
  PREPARAR_TOPLEFT:    "O1",  // O:R (4 cols)

  // Tablas de soporte (ocultas / fuera de vista) para KPIs lookup
  // TKC: 2 cols (Estado, Cantidad) -> V:W
  // Flotas: 2 cols (Estado, Cantidad) -> Y:Z
  TKC_SUPPORT_TOPLEFT:   "V1",
  FLOTA_SUPPORT_TOPLEFT: "Y1"
};

// Keys KPI (ya los tienes, pero el script los asegura)
const KPI_KEYS = [
  "updated_at",
  "tkc_cancelada",
  "tkc_en_distribucion",
  "tkc_entregada",
  "tkc_lista_distribuir",
  "flota_confirmada",
  "flota_ordenado_desp_distrib",
  "flota_total",
  "sin_asignar_total",
  "pendientes_flota_total",
  "plan_total",
  "preparar_total_min",
  "preparar_total_max"
];

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ------------------------- ENDPOINTS -------------------------

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
  try {
    let result;

    if (action === "setup_api") result = setupApi_();
    else if (action === "refresh_api") result = refreshApi_();
    else if (action === "install_triggers") {
      const minutes = Number(e.parameter.minutes || 5);
      result = installTriggers_(minutes);
    }
    else if (action === "remove_triggers") result = removeTriggers_();
    else result = { status: "error", message: `Acción '${action}' no válida.` };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents || "{}");
    const action = body.action || "";
    let result;

    if (action === "setup_api") result = setupApi_();
    else if (action === "refresh_api") result = refreshApi_();
    else if (action === "install_triggers") result = installTriggers_(Number(body.minutes || 5));
    else if (action === "remove_triggers") result = removeTriggers_();
    else result = { status: "error", message: `Acción POST '${action}' no válida.` };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ------------------------- MENÚ -------------------------

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("API Dashboard")
    .addItem("1) Setup API", "setupApi_")
    .addItem("2) Refresh API", "refreshApi_")
    .addSeparator()
    .addItem("Instalar trigger (5 min)", "installTrigger5_")
    .addItem("Eliminar triggers", "removeTriggers_")
    .addToUi();
}

function installTrigger5_() {
  installTriggers_(5);
}

// ------------------------- CORE -------------------------

function setupApi_() {
  const ss = getSpreadsheet_();
  const dash = ss.getSheetByName(SHEET_DASH);
  if (!dash) throw new Error(`No existe la hoja '${SHEET_DASH}'.`);

  let api = ss.getSheetByName(SHEET_API);
  if (!api) api = ss.insertSheet(SHEET_API);

  // 1) Headers + keys KPI
  buildKpiArea_(api);

  // 2) Headers de los bloques (normaliza layout)
  buildApiHeaders_(api);

  // 3) Detecta bloques en dashboard y escribe fórmulas
  linkBlocksFromDashboard_(dash, api);

  // 4) Fórmulas KPI (usan tablas de soporte ya importadas)
  writeKpiFormulas_(api);

  return { status: "success", message: "API creada y enlazada al dashboard." };
}

function refreshApi_() {
  const ss = getSpreadsheet_();
  const dash = ss.getSheetByName(SHEET_DASH);
  const api  = ss.getSheetByName(SHEET_API);
  if (!dash || !api) return { status: "error", message: "Faltan hojas. Ejecuta setup_api primero." };

  linkBlocksFromDashboard_(dash, api);
  writeKpiFormulas_(api);

  return { status: "success", message: "API refrescada." };
}

// ------------------------- BUILDERS -------------------------

function buildKpiArea_(api) {
  // Encabezados
  api.getRange("A1").setValue("A (Key)");
  api.getRange("B1").setValue("B (Value)");

  // Keys
  const keyValues = KPI_KEYS.map(k => [k]);
  api.getRange(2, 1, keyValues.length, 1).setValues(keyValues);

  // Limpia valores (col B) para que no queden basura
  api.getRange(2, 2, keyValues.length, 1).clearContent();
}

function buildApiHeaders_(api) {
  // SIN ASIGNAR (C:E)
  api.getRange(API_LAYOUT.SIN_ASIGNAR_TOPLEFT).setValue("Distribuidor");
  api.getRange("D1").setValue("Estado");
  api.getRange("E1").setValue("Cantidad");

  // PENDIENTES (H:K)
  api.getRange(API_LAYOUT.PENDIENTES_TOPLEFT).setValue("Distribuidor");
  api.getRange("I1").setValue("Fecha");
  api.getRange("J1").setValue("Id Orden");
  api.getRange("K1").setValue("Cantidad");

  // PLAN (M:N)
  api.getRange(API_LAYOUT.PLAN_TOPLEFT).setValue("Distribuidor");
  api.getRange("N1").setValue("Importe");

  // PREPARAR (O:R) — contiguo (sin huecos)
  api.getRange(API_LAYOUT.PREPARAR_TOPLEFT).setValue("Fecha");
  api.getRange("P1").setValue("Distribuidor");
  api.getRange("Q1").setValue("$ Min");
  api.getRange("R1").setValue("$ Max");

  // SOPORTE TKC (V:W)
  api.getRange(API_LAYOUT.TKC_SUPPORT_TOPLEFT).setValue("ESTADO");
  api.getRange("W1").setValue("Cantidad");

  // SOPORTE FLOTA (Y:Z)
  api.getRange(API_LAYOUT.FLOTA_SUPPORT_TOPLEFT).setValue("ESTADO");
  api.getRange("Z1").setValue("Cantidad");
}

// ------------------------- LINKING (AUTO-DETECT) -------------------------

function linkBlocksFromDashboard_(dash, api) {
  // Lee una “ventana” del dashboard para buscar títulos y headers
  const scan = readScanWindow_(dash, 400, 40); // 400 filas x 40 cols (A:AN aprox)

  // 1) Reporte TKC (tabla Estado/Cantidad) -> soporte V:W
  const tkcRange = findTableUnderTitle_(dash, scan, "reporte tkc", ["estado", "cantidad"]);
  if (tkcRange) {
    writeTableImportFormula_(api, API_LAYOUT.TKC_SUPPORT_TOPLEFT, tkcRange.a1, 2);
  }

  // 2) Reporte Flotas -> soporte Y:Z
  const flotaRange = findTableUnderTitle_(dash, scan, "reporte flotas", ["estado", "cantidad"]);
  if (flotaRange) {
    writeTableImportFormula_(api, API_LAYOUT.FLOTA_SUPPORT_TOPLEFT, flotaRange.a1, 2);
  }

  // 3) Órdenes sin asignar -> C:E
  const sinAsignar = findTableUnderTitle_(dash, scan, "ordenes sin asignar", ["distribuidor", "estado", "cantidad"]);
  if (sinAsignar) {
    writeTableImportFormula_(api, API_LAYOUT.SIN_ASIGNAR_TOPLEFT, sinAsignar.a1, 3);
  }

  // 4) Pendientes en flota -> H:K
  const pendientes = findTableUnderTitle_(dash, scan, "ordenes pendientes en flota", ["distribuidor", "fecha", "id", "cantidad"]);
  if (pendientes) {
    writeTableImportFormula_(api, API_LAYOUT.PENDIENTES_TOPLEFT, pendientes.a1, 4);
  }

  // 5) Plan $ -> M:N
  const plan = findTableUnderTitle_(dash, scan, "plan $", ["distribuidor", "sum", "importe"]);
  if (plan) {
    // Normalmente el pivot pone “SUM of IMPORTE”, lo aceptamos por contains.
    writeTableImportFormula_(api, API_LAYOUT.PLAN_TOPLEFT, plan.a1, 2);
  }

  // 6) $ a preparar -> O:R
  const preparar = findTableByHeaders_(dash, scan, ["fecha", "distribuidor", "min", "max"]);
  if (preparar) {
    writeTableImportFormula_(api, API_LAYOUT.PREPARAR_TOPLEFT, preparar.a1, 4);
  }
}

/**
 * Importa una tabla del dashboard hacia API con filtro de filas vacías.
 * targetTopLeft: A1 en API
 * sourceA1Range: rango A1 en dashboard (ej "D5:F30")
 * width: columnas de la tabla a mantener
 */
function writeTableImportFormula_(api, targetTopLeft, sourceA1Range, width) {
  const dashName = SHEET_DASH.replace(/'/g, "''");
  const formula =
    `=LET(t,INDIRECT("'${dashName}'!${sourceA1Range}"),` +
    `FILTER(t, INDEX(t,,1)<>"" ))`;

  api.getRange(targetTopLeft).setFormula(formula);
}

// ------------------------- KPI FORMULAS -------------------------

function writeKpiFormulas_(api) {
  // Asumimos que las tablas de soporte se importan con headers en fila 1
  // TKC soporte: V:W (Estado/Cantidad), datos desde fila 2
  // Flotas soporte: Y:Z

  // updated_at: intenta agarrar timestamp desde dashboard (si lo quieres fijo, cambia la fórmula a una celda concreta)
  api.getRange("B2").setFormula(`=NOW()`); // OJO: si tienes celda de timestamp real, cambia esto.

  // TKC
  api.getRange("B3").setFormula(`=IFERROR(XLOOKUP("Cancelada",$V$2:$V$50,$W$2:$W$50),0)`);
  api.getRange("B4").setFormula(`=IFERROR(XLOOKUP("En distribución",$V$2:$V$50,$W$2:$W$50),0)`);
  api.getRange("B5").setFormula(`=IFERROR(XLOOKUP("Entregada",$V$2:$V$50,$W$2:$W$50),0)`);
  api.getRange("B6").setFormula(`=IFERROR(XLOOKUP("Lista para distribuir",$V$2:$V$50,$W$2:$W$50),0)`);

  // Flotas
  api.getRange("B7").setFormula(`=IFERROR(XLOOKUP("Confirmada",$Y$2:$Y$50,$Z$2:$Z$50),0)`);
  api.getRange("B8").setFormula(`=IFERROR(XLOOKUP("Ordenado Desp. y Distrib.",$Y$2:$Y$50,$Z$2:$Z$50),0)`);
  api.getRange("B9").setFormula(`=IFERROR(XLOOKUP("Grand Total",$Y$2:$Y$50,$Z$2:$Z$50),0)`);

  // Totales desde bloques visibles en API
  api.getRange("B10").setFormula(`=IFERROR(SUM($E$2:$E),0)`);           // sin_asignar_total
  api.getRange("B11").setFormula(`=IFERROR(SUM($K$2:$K),0)`);           // pendientes_flota_total
  api.getRange("B12").setFormula(`=IFERROR(SUM($N$2:$N),0)`);           // plan_total
  api.getRange("B13").setFormula(`=IFERROR(SUM($Q$2:$Q),0)`);           // preparar_total_min
  api.getRange("B14").setFormula(`=IFERROR(SUM($R$2:$R),0)`);           // preparar_total_max
}

// ------------------------- TRIGGERS -------------------------

function installTriggers_(minutes) {
  if (!minutes || minutes < 1) minutes = 5;

  removeTriggers_();

  ScriptApp.newTrigger("refreshApi_")
    .timeBased()
    .everyMinutes(minutes)
    .create();

  return { status: "success", message: `Trigger instalado cada ${minutes} min.` };
}

function removeTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => ScriptApp.deleteTrigger(t));
  return { status: "success", message: "Triggers eliminados." };
}

// ------------------------- SCAN + DETECTION HELPERS -------------------------

function readScanWindow_(sheet, maxRows, maxCols) {
  const range = sheet.getRange(1, 1, maxRows, maxCols);
  const values = range.getDisplayValues();
  return { values, maxRows, maxCols };
}

function normalize_(s) {
  return String(s || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "") // quita acentos
    .trim();
}

function findCellByText_(scan, needle) {
  const n = normalize_(needle);
  const { values, maxRows, maxCols } = scan;

  for (let r = 0; r < maxRows; r++) {
    for (let c = 0; c < maxCols; c++) {
      if (normalize_(values[r][c]) === n) return { r, c };
    }
  }
  return null;
}

function findRowWithHeadersNear_(scan, anchor, headers, rowWindow = 10, colWindow = 20) {
  const { values, maxRows, maxCols } = scan;
  const startR = Math.max(0, anchor.r);
  const endR = Math.min(maxRows - 1, anchor.r + rowWindow);
  const startC = Math.max(0, anchor.c);
  const endC = Math.min(maxCols - 1, anchor.c + colWindow);

  const wanted = headers.map(h => normalize_(h));

  for (let r = startR; r <= endR; r++) {
    // Busca el header row por contains en columnas cercanas
    const row = values[r].slice(startC, endC + 1).map(normalize_);
    const ok = wanted.every(w => row.some(cell => cell.includes(w)));
    if (ok) {
      // encuentra el primer col donde aparece el primer header
      const firstHeader = wanted[0];
      let headerCol = startC;
      for (let c = startC; c <= endC; c++) {
        if (normalize_(values[r][c]).includes(firstHeader)) { headerCol = c; break; }
      }
      return { headerRow: r, headerCol };
    }
  }
  return null;
}

function detectTableBounds_(dash, headerRow0, headerCol0, width, maxDown = 200) {
  // headerRow0, headerCol0 son 0-based
  const headerRow = headerRow0 + 1; // a 1-based
  const headerCol = headerCol0 + 1;

  // baja hasta primera fila vacía en la 1ra columna de la tabla
  let lastRow = headerRow;
  for (let i = 1; i <= maxDown; i++) {
    const r = headerRow + i;
    const v = dash.getRange(r, headerCol).getDisplayValue();
    if (normalize_(v) === "") { break; }
    lastRow = r;
  }

  const lastCol = headerCol + width - 1;
  return { headerRow, headerCol, lastRow, lastCol, a1: a1_(headerRow, headerCol, lastRow, lastCol) };
}

function a1_(r1, c1, r2, c2) {
  const start = columnToLetter_(c1) + r1;
  const end = columnToLetter_(c2) + r2;
  return `${start}:${end}`;
}

function columnToLetter_(col) {
  let temp = "";
  while (col > 0) {
    let rem = (col - 1) % 26;
    temp = String.fromCharCode(65 + rem) + temp;
    col = Math.floor((col - 1) / 26);
  }
  return temp;
}

/**
 * Busca una tabla que está debajo de un título (ej. "Reporte TKC")
 * y que tenga ciertos headers.
 */
function findTableUnderTitle_(dash, scan, title, headers) {
  const anchor = findCellByText_(scan, title);
  if (!anchor) return null;

  // Encuentra fila de headers cerca del título
  const hdr = findRowWithHeadersNear_(scan, anchor, headers, 15, 25);
  if (!hdr) return null;

  // Decide ancho por headers
  const width = headers.length >= 4 ? 4 : (headers.length >= 3 ? 3 : 2);

  return detectTableBounds_(dash, hdr.headerRow, hdr.headerCol, width, 300);
}

/**
 * Busca una tabla por headers en cualquier lugar del scan.
 * Útil para "$ a preparar" si el título cambia.
 */
function findTableByHeaders_(dash, scan, headers) {
  const { values, maxRows, maxCols } = scan;
  const wanted = headers.map(h => normalize_(h));
  const width = headers.length;

  for (let r = 0; r < maxRows; r++) {
    const rowNorm = values[r].map(normalize_);
    const ok = wanted.every(w => rowNorm.some(cell => cell.includes(w)));
    if (ok) {
      // primera columna donde aparece el primer header
      const firstHeader = wanted[0];
      let c0 = 0;
      for (let c = 0; c < maxCols; c++) {
        if (normalize_(values[r][c]).includes(firstHeader)) { c0 = c; break; }
      }
      return detectTableBounds_(dash, r, c0, width, 300);
    }
  }
  return null;
}
