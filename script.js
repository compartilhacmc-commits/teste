// ======================
// CONFIG DA PLANILHA
// ======================
const SHEET_ID = "1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4";
// Aba (gid 278071504). Vamos puxar via "gviz" pelo NOME da aba.
// Se der erro por nome, me diga o nome exato da aba, que ajusto.
const SHEET_NAME = "Pendências - Consultas Especializadas - Distrito Eldorado"; // tentativa pelo título visível
// Fallback: se não funcionar, você pode colocar o nome exato da guia (tab) do Google Sheets.

const gvizUrl = (sheetName) =>
  `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?sheet=${encodeURIComponent(sheetName)}&tqx=out:json`;

let RAW = [];
let FILTERED = [];
let totalLoaded = 0;

let tsUnidade, tsCbo, tsStatus;
let chartUnidade, chartCbo;

const el = (id) => document.getElementById(id);

const COLS = {
  numSolic: "N° Solicitação",
  dataSolic: "Data da Solicitação",
  prontuario: "Nº Prontuário",
  telefone: "Telefone",
  unidade: "Unidade Solicitante",
  cbo: "Cbo Especialidade",
  dataInicio: "Data Início da Pendência",
  status: "Status",
  motivo: "Motivo da Pendência",
};

function parsePtDate(d) {
  // Espera dd/mm/yyyy
  if (!d || typeof d !== "string") return null;
  const m = d.trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (!m) return null;
  const [_, dd, mm, yyyy] = m;
  return new Date(Number(yyyy), Number(mm) - 1, Number(dd));
}

function daysBetween(a, b) {
  const ms = 24 * 60 * 60 * 1000;
  return Math.floor((b - a) / ms);
}

function safeStr(v) {
  return (v ?? "").toString().trim();
}

async function fetchSheet() {
  // Lê Google Sheets via gviz JSON
  const url = gvizUrl(SHEET_NAME);
  const res = await fetch(url, { cache: "no-store" });
  const text = await res.text();

  // gviz devolve: google.visualization.Query.setResponse({...});
  const json = JSON.parse(text.substring(text.indexOf("{"), text.lastIndexOf("}") + 1));

  const cols = json.table.cols.map(c => c.label);
  const rows = json.table.rows.map(r => r.c.map(c => (c ? c.v : "")));

  const objects = rows.map(arr => {
    const obj = {};
    cols.forEach((name, i) => obj[name] = arr[i]);
    return obj;
  });

  return objects;
}

function uniq(arr) {
  return Array.from(new Set(arr)).filter(Boolean).sort((a,b)=>a.localeCompare(b,"pt-BR"));
}

function buildSelectOptions(selectEl, values) {
  selectEl.innerHTML = "";
  values.forEach(v => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    selectEl.appendChild(opt);
  });
}

function initTomSelect() {
  const base = {
    plugins: ["remove_button", "clear_button"],
    create: false,
    persist: false,
    closeAfterSelect: false,
    hideSelected: true
  };

  tsUnidade = new TomSelect("#fUnidade", base);
  tsCbo = new TomSelect("#fEspecialidade", base);
  tsStatus = new TomSelect("#fStatus", base);

  const onChange = () => applyAndRender();
  tsUnidade.on("change", onChange);
  tsCbo.on("change", onChange);
  tsStatus.on("change", onChange);
}

function getFilterValues(ts) {
  const v = ts.getValue();
  return Array.isArray(v) ? v : (v ? [v] : []);
}

function applyFilters() {
  const fUn = new Set(getFilterValues(tsUnidade));
  const fCb = new Set(getFilterValues(tsCbo));
  const fSt = new Set(getFilterValues(tsStatus));

  FILTERED = RAW.filter(r => {
    const unidade = safeStr(r[COLS.unidade]);
    const cbo = safeStr(r[COLS.cbo]);
    const status = safeStr(r[COLS.status]);

    const okUn = fUn.size ? fUn.has(unidade) : true;
    const okCb = fCb.size ? fCb.has(cbo) : true;
    const okSt = fSt.size ? fSt.has(status) : true;

    return okUn && okCb && okSt;
  });
}

function computeKpis() {
  const today = new Date();
  const total = FILTERED.length;

  let v15 = 0, v30 = 0;
  for (const r of FILTERED) {
    const di = parsePtDate(safeStr(r[COLS.dataInicio]));
    if (!di) continue;
    const age = daysBetween(di, today); // dias desde início
    // "vencendo em 15 dias" -> até 15 dias desde início
    if (age >= 0 && age <= 15) v15++;
    if (age >= 0 && age <= 30) v30++;
  }

  const pct = totalLoaded ? Math.round((total / totalLoaded) * 100) : 0;

  el("kpiTotal").textContent = total.toLocaleString("pt-BR");
  el("kpi15").textContent = v15.toLocaleString("pt-BR");
  el("kpi30").textContent = v30.toLocaleString("pt-BR");
  el("kpiPct").textContent = `${pct}%`;
}

function groupCount(key) {
  const m = new Map();
  for (const r of FILTERED) {
    const k = safeStr(r[key]) || "(Vazio)";
    m.set(k, (m.get(k) || 0) + 1);
  }
  return Array.from(m.entries()).sort((a,b)=>b[1]-a[1]).slice(0, 12);
}

const valueLabelPlugin = {
  id: "valueLabelPlugin",
  afterDatasetsDraw(chart) {
    const { ctx } = chart;
    ctx.save();
    const meta = chart.getDatasetMeta(0);
    const data = chart.data.datasets[0].data;

    ctx.font = "700 12px system-ui, -apple-system, Segoe UI, Roboto";
    ctx.fillStyle = "#ffffff";
    ctx.textAlign = "center";
    ctx.textBaseline = "middle";

    meta.data.forEach((bar, i) => {
      const val = data[i];
      const x = bar.x;
      const y = bar.y + (bar.base - bar.y) / 2; // meio da barra
      ctx.fillText(String(val), x, y);
    });

    ctx.restore();
  }
};

function renderCharts() {
  const u = groupCount(COLS.unidade);
  const c = groupCount(COLS.cbo);

  const labelsU = u.map(x => x[0]);
  const dataU = u.map(x => x[1]);

  const labelsC = c.map(x => x[0]);
  const dataC = c.map(x => x[1]);

  const commonOptions = (color) => ({
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: { display: false },
      tooltip: { enabled: true }
    },
    scales: {
      x: {
        ticks: { maxRotation: 35, minRotation: 35, color: "#334155", font: { weight: "700", size: 11 } },
        grid: { display: false }
      },
      y: {
        beginAtZero: true,
        ticks: { color: "#64748b" },
        grid: { color: "rgba(15,23,42,.08)" }
      }
    }
  });

  // Unidade (Azul)
  if (chartUnidade) chartUnidade.destroy();
  chartUnidade = new Chart(el("chartUnidade"), {
    type: "bar",
    data: {
      labels: labelsU,
      datasets: [{
        data: dataU,
        backgroundColor: "rgba(37,99,235,.95)",
        borderRadius: 10,
        maxBarThickness: 48
      }]
    },
    options: commonOptions("#2563eb"),
    plugins: [valueLabelPlugin]
  });

  // CBO (Verde)
  if (chartCbo) chartCbo.destroy();
  chartCbo = new Chart(el("chartCbo"), {
    type: "bar",
    data: {
      labels: labelsC,
      datasets: [{
        data: dataC,
        backgroundColor: "rgba(22,163,74,.95)",
        borderRadius: 10,
        maxBarThickness: 48
      }]
    },
    options: commonOptions("#16a34a"),
    plugins: [valueLabelPlugin]
  });
}

function renderTable() {
  const body = el("tblBody");
  body.innerHTML = "";

  const rows = FILTERED.slice(0, 500); // limite visual (pode aumentar)
  for (const r of rows) {
    const tr = document.createElement("tr");
    const cells = [
      r[COLS.numSolic],
      r[COLS.dataSolic],
      r[COLS.prontuario],
      r[COLS.telefone],
      r[COLS.unidade],
      r[COLS.cbo],
      r[COLS.dataInicio],
      r[COLS.status]
    ];
    cells.forEach(v => {
      const td = document.createElement("td");
      td.textContent = safeStr(v);
      tr.appendChild(td);
    });
    body.appendChild(tr);
  }

  el("countRows").textContent =
    `Mostrando ${Math.min(FILTERED.length, rows.length)} de ${FILTERED.length} registros (após filtros).`;
}

function applyAndRender() {
  applyFilters();
  computeKpis();
  renderCharts();
  renderTable();
}

function clearFilters() {
  tsUnidade.clear(true);
  tsCbo.clear(true);
  tsStatus.clear(true);
  applyAndRender();
}

function exportExcel() {
  const out = FILTERED.map(r => ({
    "Número da Solicitação": safeStr(r[COLS.numSolic]),
    "Data da Solicitação": safeStr(r[COLS.dataSolic]),
    "Nº Prontuário": safeStr(r[COLS.prontuario]),
    "Telefone": safeStr(r[COLS.telefone]),
    "Unidade Solicitante": safeStr(r[COLS.unidade]),
    "CBO Especialidade": safeStr(r[COLS.cbo]),
    "Data Início da Pendência": safeStr(r[COLS.dataInicio]),
    "Status": safeStr(r[COLS.status]),
  }));

  const ws = XLSX.utils.json_to_sheet(out);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Pendencias");
  XLSX.writeFile(wb, "pendencias_filtradas.xlsx");
}

function setLastUpdate() {
  const d = new Date();
  el("lastUpdate").textContent = `Última atualização: ${d.toLocaleString("pt-BR")}`;
}

async function init() {
  try {
    el("connStatus").textContent = "Conectando...";
    const data = await fetchSheet();

    RAW = data;
    totalLoaded = RAW.length;

    // opções dos filtros
    const unidades = uniq(RAW.map(r => safeStr(r[COLS.unidade])));
    const cbos = uniq(RAW.map(r => safeStr(r[COLS.cbo])));
    const status = uniq(RAW.map(r => safeStr(r[COLS.status])));

    buildSelectOptions(el("fUnidade"), unidades);
    buildSelectOptions(el("fEspecialidade"), cbos);
    buildSelectOptions(el("fStatus"), status);

    initTomSelect();
    applyAndRender();

    el("connStatus").textContent = "Conectado";
    setLastUpdate();

  } catch (e) {
    console.error(e);
    el("connStatus").textContent = "Erro ao conectar";
  }
}

document.addEventListener("DOMContentLoaded", () => {
  el("btnRefresh").addEventListener("click", () => location.reload());
  el("btnClear").addEventListener("click", clearFilters);
  el("btnExcel").addEventListener("click", exportExcel);
  init();
});
