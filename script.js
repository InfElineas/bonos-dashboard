// ===================== CONFIG =====================

// Spreadsheet (solo informativo ya; el frontend NO lo lee directo)
const SPREADSHEET_ID = "1bJHM84cRxKQRJjcikVydCQeMgxVWoZmi34M4sGD-jVw";

// RANGOS A1 en la hoja API (ajustables)
const RANGES = {
  kpi: "API!A1:B20",
  sinAsignar: "API!C1:E200",
  pendientesFlota: "API!H1:K500",
  plan: "API!M1:N200",
  preparar: "API!T1:W500",

  // opcional si luego lo quieres mostrar
  // tkc: "API!V1:W200",
  // flotas: "API!Y1:Z200",
};

// Apps Script WebApp (API Builder)
const WEBAPP_URL =
  "https://script.google.com/macros/s/AKfycbzoovic1iHFl4AADdKBv_G_du-bIO2tK_IV2vQJvzcc5m53FwZvfLHc1dnHk4K7pG38/exec";


// ===================== HELPERS =====================

function escapeHtml(s) {
  return String(s ?? "").replace(/[&<>"']/g, m => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    '"': "&quot;",
    "'": "&#039;"
  }[m]));
}

function money(n) {
  const v = Number(String(n).replace(/[^0-9.-]/g, ""));
  if (Number.isNaN(v)) return String(n ?? "");
  return v.toLocaleString("en-US", { style: "currency", currency: "USD", maximumFractionDigits: 0 });
}

function notify(msg, type = "ok") {
  if (type === "err") console.error(msg);
  else console.log(msg);
  // Si quieres UI real, aquí conectas tu toast.
}


// ===================== WEBAPP CALLS =====================

async function callWebApp(action, params = {}) {
  if (!WEBAPP_URL) throw new Error("Falta WEBAPP_URL");

  const url = new URL(WEBAPP_URL);
  url.searchParams.set("action", action);

  Object.entries(params).forEach(([k, v]) => url.searchParams.set(k, String(v)));

  const res = await fetch(url.toString(), { method: "GET" });
  if (!res.ok) throw new Error(`WebApp respondió ${res.status}`);

  const data = await res.json();
  if (data.status && data.status !== "success") {
    throw new Error(data.message || "WebApp devolvió error");
  }
  return data;
}

async function fetchApiRange(rangeA1) {
  const data = await callWebApp("get_api", { range: rangeA1 });
  // data.values: matriz 2D (displayValues)
  return data.values || [];
}


// ===================== RENDER =====================

function renderKpisFromValues(values) {
  // values = [[Key, Value], ...] (incluye headers)
  const map = {};
  values.slice(1).forEach(row => {
    const k = row?.[0];
    const v = row?.[1];
    if (k) map[String(k).trim()] = v;
  });

  const kpis = [
    { key: "tkc_entregada", label: "TKC Entregadas", fmt: v => v },
    { key: "tkc_en_distribucion", label: "TKC En distribución", fmt: v => v },
    { key: "sin_asignar_total", label: "Sin asignar", fmt: v => v },
    { key: "pendientes_flota_total", label: "Pendientes flota", fmt: v => v },
    { key: "plan_total", label: "Plan $", fmt: money },
    { key: "preparar_total_max", label: "$ a preparar MAX", fmt: money },
  ];

  const grid = document.getElementById("kpiGrid");
  if (grid) {
    grid.innerHTML = kpis.map(k => `
      <div class="kpi">
        <div class="k">${escapeHtml(k.label)}</div>
        <div class="v">${escapeHtml(String(k.fmt(map[k.key] ?? "—")))}</div>
      </div>
    `).join("");
  }

  const updated = map["updated_at"] ?? "—";
  const last = document.getElementById("lastUpdate");
  if (last) last.textContent = `Actualización: ${updated}`;
}

function renderTableFromValues(tableId, values) {
  // values = matriz 2D (incluye headers)
  const el = document.getElementById(tableId);
  if (!el) return;

  if (!values || values.length === 0) {
    el.innerHTML = "";
    return;
  }

  const cols = values[0] || [];
  const rows = values.slice(1)
    .filter(r => r && r.some(v => String(v ?? "").trim() !== ""));

  const thead = `<thead><tr>${cols.map(c => `<th>${escapeHtml(c)}</th>`).join("")}</tr></thead>`;
  const tbody = `<tbody>${
    rows.map(r => `<tr>${r.map(v => `<td>${escapeHtml(String(v ?? ""))}</td>`).join("")}</tr>`).join("")
  }</tbody>`;

  el.innerHTML = thead + tbody;
}


// ===================== ACTIONS =====================

async function setupApiFromWeb() {
  const btn = document.getElementById("setupApiBtn");
  if (btn) btn.disabled = true;

  try {
    notify("Configurando API…");
    const result = await callWebApp("setup_api");
    notify(result.message || "API configurada.");
    await loadBonosDashboard();
  } catch (e) {
    console.error(e);
    alert(`Setup falló: ${e.message}`);
  } finally {
    if (btn) btn.disabled = false;
  }
}

async function refreshApiFromWeb({ silent = false } = {}) {
  const btn = document.getElementById("refreshDashboard");
  if (btn) btn.disabled = true;

  try {
    if (!silent) notify("Actualizando API…");
    const result = await callWebApp("refresh_api");
    if (!silent) notify(result.message || "API actualizada.");
    await loadBonosDashboard();
  } catch (e) {
    console.error(e);
    if (!silent) alert(`Refresh falló: ${e.message}`);
    // Intentamos render con lo último disponible
    await loadBonosDashboard();
  } finally {
    if (btn) btn.disabled = false;
  }
}


// ===================== DASHBOARD LOAD =====================

async function loadBonosDashboard() {
  try {
    const [kpi, sinA, pend, plan, prep] = await Promise.all([
      fetchApiRange(RANGES.kpi),
      fetchApiRange(RANGES.sinAsignar),
      fetchApiRange(RANGES.pendientesFlota),
      fetchApiRange(RANGES.plan),
      fetchApiRange(RANGES.preparar),
    ]);

    renderKpisFromValues(kpi);
    renderTableFromValues("sinAsignarTable", sinA);
    renderTableFromValues("pendientesFlotaTable", pend);
    renderTableFromValues("planTable", plan);
    renderTableFromValues("prepararTable", prep);

  } catch (e) {
    console.error(e);
    alert(e.message || "Error cargando dashboard desde WebApp.");
  }
}


// ===================== BOOTSTRAP =====================

document.addEventListener("DOMContentLoaded", () => {
  const refreshBtn = document.getElementById("refreshDashboard");
  if (refreshBtn) refreshBtn.addEventListener("click", () => refreshApiFromWeb());

  const setupBtn = document.getElementById("setupApiBtn");
  if (setupBtn) setupBtn.addEventListener("click", setupApiFromWeb);

  // Primera carga: refresca API y pinta
  refreshApiFromWeb({ silent: true });

  // Auto-refresh opcional
  // setInterval(() => refreshApiFromWeb({ silent: true }), 5 * 60 * 1000);
});
