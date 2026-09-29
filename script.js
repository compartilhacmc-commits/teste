'use strict';

Chart.register(ChartDataLabels);

// ============================================================
// PLUGIN CENTRAL TEXT – para gráficos do tipo Doughnut
// ============================================================
const centerTextPlugin = {
  id: 'centerText',
  beforeDraw(chart) {
    const opts = chart.config?.options?.plugins?.centerText;
    if (!opts || !opts.enabled) return;
    const { ctx, chartArea } = chart;
    if (!chartArea) return;
    const cx = (chartArea.left + chartArea.right) / 2;
    const cy = (chartArea.top + chartArea.bottom) / 2;
    ctx.save();
    ctx.font = `800 ${opts.fontSize || 22}px Inter, sans-serif`;
    ctx.fillStyle = opts.valueColor || '#1e3a5f';
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    ctx.fillText(opts.value, cx, cy - 9);
    ctx.font = `500 10px Inter, sans-serif`;
    ctx.fillStyle = opts.labelColor || '#7a8fa6';
    ctx.fillText(opts.label || '', cx, cy + 12);
    ctx.restore();
  }
};
Chart.register(centerTextPlugin);

// ============================================================
// CONFIGURAÇÕES
// ============================================================
const SHEET_ID  = '14DiFK9EW36s8ntkukyhiRMJxcX0ghG5XTzWb1TwpI2Q';
const SHEET_GID = '0';
const CSV_URL   = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${SHEET_GID}`;
const CACHE_KEY = 'cmcDashboardData_v2';
const CACHE_EXPIRY_HOURS = 0.5;

// ============================================================
// MAPEAMENTO DE OPERADORES
// ============================================================
const OPERADORES = {
  '336': 'Giovanna Borello',
  '531': 'Rosangela de Jesus',
  '594': 'Ariana Trindade',
  '536': 'Magaly Mendes',
  '591': 'Ana P Araujo',
  '535': 'Cristina Fonseca',
  '534': 'Erica Lucia',
  '533': 'Natalia Bretas',
  '595': 'Eliane Pereira',
  '540': 'Carolina de Avelar',
  '541': 'Daniela Fonseca',
  '532': 'Graziele Alves'
};

const DISTRITO_MAP = {
  // ============================================================
  // ELDORADO
  // ============================================================
  'UNIDADE BASICA DE SAUDE JARDIM BANDEIRANTES': 'ELDORADO',
  'UNIDADE BASICA DE SAUDE AGUA BRANCA': 'ELDORADO',
  'UNIDADE BASICA DE SAUDE XV': 'ELDORADO',
  'UNIDADE BASICA DE SAUDE CSU ELDORADO': 'ELDORADO',
  'UNIDADE BASICA DE SAUDE PARQUE SAO JOAO': 'ELDORADO',
  'UNIDADE BASICA DE SAUDE JARDIM ELDORADO': 'ELDORADO',
  'UNIDADE BASICA DE SAUDE NOVO ELDORADO': 'ELDORADO',
  'UNIDADE BASICA DE SAUDE SANTA CRUZ': 'ELDORADO',
  'UNIDADE BASICA DE SAUDE PEROBAS': 'ELDORADO',
  'UNIDADE BASICA DE SAUDE BELA VISTA': 'ELDORADO',

  // ============================================================
  // INDUSTRIAL
  // ============================================================
  'UNIDADE BASICA DE SAUDE INDUSTRIAL III SECAO': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE JARDIM INDUSTRIAL': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE VILA SAO PAULO': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE SANDOVAL DE AZEVEDO': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE JOAO EVANGELISTA': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE BANDEIRANTES': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE AMAZONAS': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE AMAZONAS I': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE VILA DINIZ': 'INDUSTRIAL',

  // ============================================================
  // NACIONAL
  // ============================================================
  'UNIDADE BASICA DE SAUDE NACIONAL': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE ILDA EFIGENIA': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE JOAQUIM MURTINHO': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE XANGRILA': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE AMENDOEIRAS': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE ESTRELA DALVA': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE TIJUCA': 'NACIONAL',
  'UNIDADE DE SAUDE DA FAMILIA ESTRELA DALVA USF 54': 'NACIONAL',
  'UNIDADE DE SAUDE DA FAMILIA NACIONAL II USF 23': 'NACIONAL',

  // ============================================================
  // PETROLANDIA
  // ============================================================
  'UNIDADE BASICA DE SAUDE PETROLANDIA': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE TROPICAL II': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE TROPICAL I': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE SAPUCAIAS': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE SAPUCAIAS II': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE DUQUE DE CAXIAS': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE SAO LUIZ II': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE SAO LUIZ I': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE CAMPO ALTO': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE ESTANCIAS IMPERIAIS': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE NASCENTES IMPERIAIS': 'PETROLANDIA',

  // ============================================================
  // RESSACA
  // ============================================================
  'UNIDADE BASICA DE SAUDE CAMPINA VERDE': 'RESSACA',
  'UNIDADE BASICA DE SAUDE LAGUNA': 'RESSACA',
  'UNIDADE BASICA DE SAUDE ARPOADOR': 'RESSACA',
  'UNIDADE BASICA DE SAUDE SAO JOAQUIM': 'RESSACA',
  'UNIDADE BASICA DE SAUDE PARQUE TURISTA': 'RESSACA',
  'UNIDADE BASICA DE SAUDE VILA PEROLA': 'RESSACA',
  'UNIDADE BASICA DE SAUDE NOVO PROGRESSO II': 'RESSACA',
  'UNIDADE BASICA DE SAUDE NOVO PROGRESSO I': 'RESSACA',
  'UNIDADE BASICA DE SAUDE COLORADO': 'RESSACA',
  'UNIDADE BASICA DE SAUDE CANDIDA FERREIRA': 'RESSACA',
  'UNIDADE DE SAUDE DA FAMILIA VILA PEROLA II USF 84': 'RESSACA',
  'UNIDADE BASICA DE SAUDE PRESIDENTE KENNEDY': 'RESSACA',
  'UNIDADE BASICA DE SAUDE OITIS': 'RESSACA',
  'UNIDADE BASICA DE SAUDE MORADA NOVA': 'RESSACA',
  'UNIDADE DE SAUDE DA FAMILIA JARDIM LAGUNA III USF 42': 'RESSACA',
  'UNIDADE DE SAUDE DA FAMILIA PARQUE RECREIO USF 35': 'RESSACA',

  // ============================================================
  // RIACHO
  // ============================================================
  'UNIDADE BASICA DE SAUDE INCONFIDENTES': 'RIACHO',
  'UNIDADE BASICA DE SAUDE RIACHO': 'RIACHO',
  'UNIDADE BASICA DE SAUDE FLAMENGO': 'RIACHO',
  'UNIDADE BASICA DE SAUDE NOVO RIACHO': 'RIACHO',
  'UNIDADE BASICA DE SAUDE DURVAL DE BARROS': 'RIACHO',
  'UNIDADE BASICA DE SAUDE MONTE CASTELO': 'RIACHO',

  // ============================================================
  // SEDE
  // ============================================================
  'UNIDADE BASICA DE SAUDE CHACARAS': 'SEDE',
  'UNIDADE BASICA DE SAUDE CANADA': 'SEDE',
  'UNIDADE BASICA DE SAUDE CENTRO (CAD)': 'SEDE',
  'UBS BERNARDO MONTEIRO/MOACIR PINTO MOREIRA': 'SEDE',
  'UNIDADE BASICA DE SAUDE BERNARDO MONTEIRO/MOACIR PINTO MOREIRA': 'SEDE',
  'UNIDADE BASICA DE SAUDE LINDA VISTA': 'SEDE',
  'UNIDADE BASICA DE SAUDE SANTA HELENA': 'SEDE',
  'UNIDADE BASICA DE SAUDE VILA ITALIA': 'SEDE',
  'UNIDADE BASICA DE SAUDE BEATRIZ USF 72': 'SEDE',
  'UNIDADE BASICA DE SAUDE MARIA DA CONCEICAO': 'SEDE',
  'UNIDADE BASICA DE SAUDE FUNCIONARIOS': 'SEDE',
  'UNIDADE BASICA DE SAUDE PRAIA': 'SEDE',
  'UBS UNIDADE XVI (SEDE)': 'SEDE',
  'UNIDADE BASICA DE SAUDE SESC': 'SEDE',
  'UNIDADE BASICA DE SAUDE ALVORADA': 'SEDE',
  'UNIDADE BASICA DE SAUDE BARRAGINHA': 'SEDE',
  'UNIDADE BASICA DE SAUDE PRESIDENTE VARGAS': 'SEDE',
  'UNIDADE DE SAUDE DA FAMILIA FUNCIONARIOS II USF 55': 'SEDE',
  'UNIDADE DE SAUDE DA FAMILIA SAO MATEUS USF 28': 'SEDE',
  'UNIDADE DE SAUDE DA FAMILIA ARVOREDO USF 32': 'SEDE',
  'UNIDADE DE SAUDE DA FAMILIA GUANABARA USF 37': 'SEDE',
  'UNIDADE DE SAUDE DA FAMILIA MILANES USF 44': 'SEDE',

  // ============================================================
  // VARGEM DAS FLORES
  // ============================================================
  'UNIDADE BASICA DE SAUDE VILA RENASCER': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE NOVA CONTAGEM': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE NOVA CONTAGEM I': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE VILA SOLEDADE': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE ESTALEIRO': 'VARGEM DAS FLORES',
  'CERESP CONTAGEM': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE RETIRO II': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE RETIRO': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE IPE AMARELO': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE SAO JUDAS TADEU': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE VILA ESPERANCA': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE DARCY RIBEIRO': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE ICAIVERA': 'VARGEM DAS FLORES',
  'CONTAGEM PENITENCIARIA NELSON HUNGRIA': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE TUPA': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE TUPÃƒ': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE LIBERDADE II': 'VARGEM DAS FLORES',
  'UNIDADE DE SAUDE DA FAMILIA VILA SOLEDADE II USF 10': 'VARGEM DAS FLORES',
  'UNIDADE DE SAUDE DA FAMILIA RETIRO III USF 03': 'VARGEM DAS FLORES',
  'UNIDADE DE SAUDE DA FAMILIA SAO JUDAS TADEU I USF 04': 'VARGEM DAS FLORES',
  'UNIDADE DE REFERENCIA A SAUDE DA FAMILIA VARGEM DAS FLORES': 'VARGEM DAS FLORES',
};

const CAE_UNITS = ['CAE IRIA DINIZ', 'CAE RESSACA', 'CEAPS'];
const MESES_PT  = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                   'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

// ============================================================
// PALETAS DE CORES
// ============================================================
const PALETTE_AZUL = [
  '#0a1628', '#0f1f3a', '#142844', '#1a3256',
  '#1e3a5f', '#254a7a', '#2d5494', '#3a6bb0',
  '#4a90d9', '#6aaff0', '#8fc7f5', '#b5dffa',
  '#d4edfd', '#eaf6ff'
];

const PALETTE_ROSA = [
  '#4a0d2f', '#641240', '#7f1852', '#9a1e64',
  '#b42476', '#cb3388', '#de4b99', '#e966a9',
  '#f181b9', '#f69cc9', '#fab7d8', '#fdd1e7',
  '#fee5f2', '#fff3f8'
];

const PALETTE_ABSENTEISMO = [
  '#7c2d12', '#9a3412', '#c2410c', '#ea580c',
  '#f97316', '#fb923c', '#fdba74', '#fed7aa',
  '#ffedd5', '#fff7ed', '#fffaf5', '#fffcfa',
  '#fffefd', '#ffffff'
];

const PALETTE_RECEPCIONADOS = [
  '#0a1a10', '#0f2618', '#143020', '#1a3d28',
  '#1f4d30', '#26603a', '#2e7042', '#3a8a52',
  '#48a064', '#5ab87a', '#78d096', '#a0e0b4',
  '#c8f0d4', '#e8f8ec'
];

const PALETTE_ATENDIDOS = [
  '#0a1a2a', '#0f2840', '#143656', '#1a446c',
  '#1f5282', '#266098', '#2e6eae', '#3a7cc4',
  '#4a8ada', '#5a98e8', '#78a8f0', '#a0c0f5',
  '#c8d8fa', '#e8f0fd'
];

const PALETTE_CANCELADOS = [
  '#1a0a0a', '#2d1010', '#3d1515', '#4d1a1a',
  '#6b2020', '#8a2a2a', '#a83232', '#c93d3d',
  '#e04a4a', '#ea6a6a', '#f08c8c', '#f5b0b0',
  '#fad0d0', '#fce8e8'
];

const PALETTE_TRANSFERIDOS = [
  '#140a1a', '#1e0f28', '#281430', '#321a3d',
  '#3d1f4d', '#4a2660', '#582e70', '#6a3a8a',
  '#7c48a0', '#945ab8', '#b078d0', '#c8a0e0',
  '#e0c8f0', '#f0e8f8'
];

const TOOLTIP_BASE = {
  backgroundColor: 'rgba(20,40,68,0.92)',
  titleFont:  { family: 'Inter', size: 12, weight: '700' },
  bodyFont:   { family: 'Inter', size: 11 },
  padding: 12, cornerRadius: 10,
  callbacks: { label: ctx => ` ${fmt(ctx.raw)}` }
};

// ============================================================
// ESTADO GLOBAL
// ============================================================
let allData       = [];
let filteredData  = [];
let tableData     = [];
let tableSearched = [];
let currentPage   = 1;
let sortColIdx    = -1;
let sortAscFlag   = true;

let tableDataResumido = [];
let tableSearchedResumido = [];
let currentPageResumido = 1;
let sortColIdxResumido = -1;
let sortAscFlagResumido = true;

const multiSelectState = {};

// Instâncias dos gráficos
let chartDistrito, chartTipoAtendimento, chartEspecialidade, chartPrestador;
let chartSituacao, chartMeses;
let chartAbsenteismoEsp, chartAbsenteismoDist, chartAbsenteismoMensal, chartAbsenteismoPrestador;
let chartCancelamentosDist, chartCancelamentosEsp, chartCancelamentosPrestador, chartCancelamentosMensal;
let chartRecepcionadosDistrito, chartRecepcionadosEspecialidade, chartRecepcionadosPrestador, chartRecepcionadosMensal;
let chartAtendidosDistrito, chartAtendidosEspecialidade, chartAtendidosPrestador, chartAtendidosMensal;
let chartTransferidosDistrito, chartTransferidosEspecialidade, chartTransferidosPrestador, chartTransferidosMensal;
let chartPrimeiraConsultaDistrito, chartRetornoDistrito, chartComparativoDistrito, chartDistritoRosca;

// ============================================================
// UTILITÁRIOS
// ============================================================
function norm(str) {
  if (!str) return '';
  return str.toString().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '')
    .replace(/\s+/g, ' ').trim();
}

function titleCase(str) {
  if (!str) return '';
  const artigos = ['DE','DA','DO','DAS','DOS','E','A','O','EM','NO','NA','NOS','NAS','POR','COM','PARA'];
  return str.toString().toLowerCase().split(' ').map((w, i) => {
    if (i === 0 || !artigos.includes(w.toUpperCase()))
      return w.charAt(0).toUpperCase() + w.slice(1);
    return w;
  }).join(' ');
}

function formatCBO(nomeCBO, especialidade) {
  if (!nomeCBO) return '';
  let name = nomeCBO.toString().trim();
  name = name.replace(/^M[eé]dico[:\-\s]*/i, '').trim();
  name = titleCase(name);
  if (especialidade) {
    const esp    = especialidade.toString().trim();
    const addons = ['Adulto','Infantil','Geral','Pediátrico','Pediátrica'];
    for (const addon of addons) {
      if (norm(esp).includes(norm(addon)) && !norm(name).includes(norm(addon))) {
        name += ' ' + addon; break;
      }
    }
  }
  return name;
}

function formatProfissional(nome) {
  if (!nome) return '';
  return titleCase(nome.toString().trim().replace(/^\d+\s*[-–]?\s*/, '').trim());
}

function getUnidade(row) {
  const solicitante = (row['UNIDADE SOLICITANTE'] || '').toString().trim();
  const solNorm     = norm(solicitante);
  for (const cae of CAE_UNITS) {
    if (solNorm === norm(cae)) {
      const ref = (row['UNIDADE DE REFERÊNCIA'] || row['UNIDADE DE REFERENCIA'] || '').toString().trim();
      if (ref) return ref;
      return solicitante;
    }
  }
  return solicitante;
}

function getDistrito(unidade) {
  if (!unidade) return 'NÃO IDENTIFICADO';
  const key = norm(unidade);
  for (const [mapKey, distrito] of Object.entries(DISTRITO_MAP)) {
    if (norm(mapKey) === key) return distrito;
  }
  for (const [mapKey, distrito] of Object.entries(DISTRITO_MAP)) {
    if (key.includes(norm(mapKey)) || norm(mapKey).includes(key)) return distrito;
  }
  return 'OUTROS';
}

function getTipoAtendimento(val) {
  const v = (val || '').toString().trim().toUpperCase();
  if (v === 'P') return 'Primeira Consulta';
  if (v === 'R') return 'Retorno';
  return val || '–';
}

function getSituacaoLabel(val) {
  const map = {
    AGE:'Agendados',
    CAN:'Cancelados',
    FAL:'Faltosos',
    REC:'Recepcionados',
    TRA:'Transferidos',
    ATD:'Atendidos'
  };
  return map[(val||'').toUpperCase()] || val || '–';
}

function getOperador(codigo) {
  const key = (codigo || '').toString().trim();
  return OPERADORES[key] || 'Outros';
}

function fmt(n) { return (n || 0).toLocaleString('pt-BR'); }

function parseDate(str) {
  if (!str) return null;
  str = str.toString().trim();
  let m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2]-1, +m[1]);
  m = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2]-1, +m[3]);
  return null;
}

function isSameDay(d1, d2) {
  if (!d1 || !d2) return false;
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth()    === d2.getMonth()    &&
         d1.getDate()     === d2.getDate();
}

function compareDate(selectedDate, rowDate) {
  if (!selectedDate || !rowDate) return false;
  const d1 = selectedDate.getFullYear() + '-' +
             String(selectedDate.getMonth() + 1).padStart(2, '0') + '-' +
             String(selectedDate.getDate()).padStart(2, '0');
  const d2 = rowDate.getFullYear() + '-' +
             String(rowDate.getMonth() + 1).padStart(2, '0') + '-' +
             String(rowDate.getDate()).padStart(2, '0');
  return d1 === d2;
}

// ============================================================
// MULTI SELECT
// ============================================================
function initMultiSelect(id, values) {
  multiSelectState[id] = { selected: [], open: false, values: values || [] };
  renderMultiSelect(id);
}

function renderMultiSelect(id) {
  const state = multiSelectState[id];
  if (!state) return;
  const wrapper = document.getElementById(id);
  if (!wrapper) return;

  const placeholder = wrapper.querySelector('.multi-select-placeholder');
  const optionsContainer = wrapper.querySelector('.multi-select-options');

  if (state.selected.length === 0) {
    placeholder.textContent = placeholder.getAttribute('data-default') || 'Todos';
    placeholder.classList.remove('selected');
  } else if (state.selected.length <= 3) {
    placeholder.textContent = state.selected.join(', ');
    placeholder.classList.add('selected');
  } else {
    placeholder.textContent = `${state.selected.length} selecionados`;
    placeholder.classList.add('selected');
  }

  if (optionsContainer) {
    const searchInput = wrapper.querySelector('.multi-select-search input');
    const searchVal = searchInput ? searchInput.value.toLowerCase() : '';
    const filteredValues = searchVal ? state.values.filter(v => v.toLowerCase().includes(searchVal)) : state.values;

    optionsContainer.innerHTML = filteredValues.map(v => `
      <label class="multi-select-option">
        <input type="checkbox" value="${v}" ${state.selected.includes(v) ? 'checked' : ''} onchange="toggleMultiSelectOption('${id}', '${v.replace(/'/g, "\\'")}', this.checked)" />
        <span>${v}</span>
      </label>
    `).join('');
  }
}

function toggleMultiSelect(id) {
  const state = multiSelectState[id];
  if (!state) return;
  const wrapper = document.getElementById(id);
  if (!wrapper) return;

  Object.keys(multiSelectState).forEach(key => {
    if (key !== id) {
      multiSelectState[key].open = false;
      const w = document.getElementById(key);
      if (w) w.classList.remove('open');
    }
  });

  state.open = !state.open;
  if (state.open) {
    wrapper.classList.add('open');
    const searchInput = wrapper.querySelector('.multi-select-search input');
    if (searchInput) { searchInput.value = ''; }
    renderMultiSelect(id);
  } else {
    wrapper.classList.remove('open');
  }
}

function toggleMultiSelectOption(id, value, checked) {
  const state = multiSelectState[id];
  if (!state) return;
  if (checked) {
    if (!state.selected.includes(value)) state.selected.push(value);
  } else {
    state.selected = state.selected.filter(v => v !== value);
  }
  renderMultiSelect(id);
  applyFilters();
}

function selectAllMultiSelect(id) {
  const state = multiSelectState[id];
  if (!state) return;
  state.selected = [...state.values];
  renderMultiSelect(id);
  applyFilters();
}

function clearMultiSelect(id) {
  const state = multiSelectState[id];
  if (!state) return;
  state.selected = [];
  renderMultiSelect(id);
  applyFilters();
}

function filterMultiSelectOptions(id, query) {
  const state = multiSelectState[id];
  if (!state) return;
  if (!query) {
    state.values = [...state._allValues];
  } else {
    const q = query.toLowerCase();
    state.values = state._allValues.filter(v => v.toLowerCase().includes(q));
  }
  renderMultiSelect(id);
}

function getMultiSelectValues(id) {
  const state = multiSelectState[id];
  return state ? state.selected : [];
}

document.addEventListener('click', function(e) {
  if (!e.target.closest('.multi-select-wrapper')) {
    Object.keys(multiSelectState).forEach(key => {
      multiSelectState[key].open = false;
      const w = document.getElementById(key);
      if (w) w.classList.remove('open');
    });
  }
});

// ============================================================
// TABS / ABAS
// ============================================================
function initTabs() {
  const tabBtns = document.querySelectorAll('.tab-btn');
  tabBtns.forEach(btn => {
    btn.addEventListener('click', function() {
      const tabId = this.getAttribute('data-tab');
      tabBtns.forEach(b => b.classList.remove('active'));
      document.querySelectorAll('.tab-panel').forEach(p => p.classList.remove('active'));
      this.classList.add('active');
      const panel = document.getElementById('tab-' + tabId);
      if (panel) {
        panel.classList.add('active');
        renderChartsForTab(tabId);
      }
    });
  });
}

function renderChartsForTab(tabId) {
  destroyAllCharts();
  switch(tabId) {
    case 'visao-geral':
      renderChartDistrito();
      renderChartTipoAtendimento();
      renderChartEspecialidade();
      renderChartPrestador();
      renderChartSituacao();
      renderChartMeses();
      break;
    case 'tabela':
      renderTable();
      break;
    case 'profissionais-atendimento':
      renderTableResumido();
      break;
    case 'agendamentos-distrito':
      renderChartPrimeiraConsultaDistrito();
      renderChartRetornoDistrito();
      renderChartComparativoDistrito();
      renderChartDistritoRosca();
      break;
    case 'absenteismo':
      renderChartAbsenteismoEsp();
      renderChartAbsenteismoDist();
      renderChartAbsenteismoMensal();
      renderChartAbsenteismoPrestador();
      break;
    case 'recepcionados':
      renderChartRecepcionadosDistrito();
      renderChartRecepcionadosEspecialidade();
      renderChartRecepcionadosPrestador();
      renderChartRecepcionadosMensal();
      break;
    case 'atendidos':
      renderChartAtendidosDistrito();
      renderChartAtendidosEspecialidade();
      renderChartAtendidosPrestador();
      renderChartAtendidosMensal();
      break;
    case 'cancelados':
      renderChartCancelamentosDist();
      renderChartCancelamentosEsp();
      renderChartCancelamentosPrestador();
      renderChartCancelamentosMensal();
      break;
    case 'transferidos':
      renderChartTransferidosDistrito();
      renderChartTransferidosEspecialidade();
      renderChartTransferidosPrestador();
      renderChartTransferidosMensal();
      break;
  }
}

// ============================================================
// DESTRUIR TODOS OS GRÁFICOS
// ============================================================
function destroyAllCharts() {
  const refs = [
    'chartDistrito', 'chartTipoAtendimento', 'chartEspecialidade', 'chartPrestador',
    'chartSituacao', 'chartMeses', 'chartAbsenteismoEsp', 'chartAbsenteismoDist',
    'chartAbsenteismoMensal', 'chartAbsenteismoPrestador', 'chartCancelamentosDist',
    'chartCancelamentosEsp', 'chartCancelamentosPrestador', 'chartCancelamentosMensal',
    'chartRecepcionadosDistrito', 'chartRecepcionadosEspecialidade', 'chartRecepcionadosPrestador',
    'chartRecepcionadosMensal', 'chartAtendidosDistrito', 'chartAtendidosEspecialidade',
    'chartAtendidosPrestador', 'chartAtendidosMensal', 'chartTransferidosDistrito',
    'chartTransferidosEspecialidade', 'chartTransferidosPrestador', 'chartTransferidosMensal',
    'chartPrimeiraConsultaDistrito', 'chartRetornoDistrito', 'chartComparativoDistrito', 'chartDistritoRosca'
  ];
  refs.forEach(ref => {
    if (window[ref]) {
      try { window[ref].destroy(); } catch(e) {}
      window[ref] = null;
    }
  });
}

// ============================================================
// CARREGAR DADOS
// ============================================================
async function loadData(forceRefresh = false) {
  const btnRefresh = document.getElementById('btnRefresh');
  const icon = document.getElementById('refreshIcon');

  if (btnRefresh) btnRefresh.disabled = true;
  if (icon) icon.classList.add('spinning');
  showLoading(true);
  setStatus('Carregando...', false);

  try {
    if (!forceRefresh) {
      const cached = getCachedData();
      if (cached) {
        console.log('📦 Dados carregados do cache.');
        allData = cached.data;
        onDataLoaded();
        setStatus('Conectado (cache)', true);
        updateLastUpdate(new Date(cached.timestamp));
        showLoading(false);
        if (icon) icon.classList.remove('spinning');
        if (btnRefresh) btnRefresh.disabled = false;
        fetchDataInBackground();
        return;
      }
    }

    console.log('🔄 Buscando dados novos da planilha...');
    const freshData = await fetchFreshData();
    if (freshData && freshData.length > 0) {
      allData = freshData;
      const timestamp = new Date().toISOString();
      setCachedData(allData, timestamp);
      onDataLoaded();
      setStatus('Conectado', true);
      updateLastUpdate(new Date());
      console.log(`✅ ${freshData.length} registros carregados!`);
    } else {
      throw new Error('Falha ao buscar dados da planilha.');
    }
  } catch (err) {
    console.error('❌ Erro ao carregar dados:', err);

    const expiredCache = getCachedData(true);
    if (expiredCache && expiredCache.data && expiredCache.data.length > 0) {
      allData = expiredCache.data;
      onDataLoaded();
      setStatus('⚠️ Usando cache antigo (offline)', false);
      updateLastUpdate(new Date(expiredCache.timestamp));
      showError('Não foi possível atualizar. Usando dados em cache.');
    } else {
      showError('❌ Erro ao carregar dados. Verifique sua conexão.');
      const msg = document.getElementById('statusText');
      if (msg) msg.textContent = '⚠️ Sem conexão com a planilha';
    }
  } finally {
    showLoading(false);
    if (icon) icon.classList.remove('spinning');
    if (btnRefresh) btnRefresh.disabled = false;
  }
}

async function fetchFreshData() {
  try {
    const url = CSV_URL + '&t=' + Date.now();
    console.log('📡 Fazendo requisição para:', url);

    const response = await fetch(url);

    if (!response.ok) {
      throw new Error(`HTTP ${response.status} - ${response.statusText}`);
    }

    const text = await response.text();
    console.log(`📄 Resposta recebida: ${text.length} caracteres`);

    if (!text || text.trim().length === 0) {
      throw new Error('Resposta vazia da planilha');
    }

    const firstLine = text.split('\n')[0];
    if (!firstLine || !firstLine.includes(',')) {
      console.warn('⚠️ Possível erro na resposta.');
      const cleanedText = text.replace(/^[^{]*\{[^}]*\}\s*/, '');
      if (cleanedText !== text) {
        console.log('🔄 Texto limpo, tentando novamente...');
        return parseCSV(cleanedText);
      }
    }

    return parseCSV(text);

  } catch (err) {
    console.error('❌ Erro ao buscar dados:', err);
    throw err;
  }
}

function parseCSV(text) {
  return new Promise((resolve, reject) => {
    Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      transformHeader: h => h.trim(),
      complete(results) {
        if (results.errors && results.errors.length > 0) {
          console.warn('⚠️ Erros no parse:', results.errors);
        }
        const data = normalizeData(results.data);
        console.log(`📊 ${data.length} registros normalizados`);
        resolve(data);
      },
      error(err) {
        console.error('❌ Erro no Papa Parse:', err);
        reject(err);
      }
    });
  });
}

async function fetchDataInBackground() {
  try {
    console.log('🔄 Atualização em segundo plano iniciada...');
    const freshData = await fetchFreshData();
    if (freshData && freshData.length > 0) {
      const currentDataStr = JSON.stringify(allData);
      const newDataStr = JSON.stringify(freshData);

      if (currentDataStr !== newDataStr) {
        console.log('🔄 Dados atualizados em segundo plano!');
        allData = freshData;
        const timestamp = new Date().toISOString();
        setCachedData(allData, timestamp);
        onDataLoaded();
        setStatus('🔄 Dados atualizados!', true);
        updateLastUpdate(new Date());
        showToast('📊 Dados da planilha foram atualizados!');
      } else {
        console.log('✅ Dados já estão atualizados.');
        setCachedData(allData, new Date().toISOString());
      }
    }
  } catch (e) {
    console.warn('⚠️ Atualização em segundo plano falhou:', e.message);
  }
}

function setCachedData(data, timestamp) {
  try {
    const dataStr = JSON.stringify(data);
    const sizeInMB = (dataStr.length * 2) / (1024 * 1024);
    console.log(`📊 Tamanho dos dados: ${sizeInMB.toFixed(2)} MB`);

    if (sizeInMB > 4) {
      console.warn('⚠️ Dados muito grandes para cache (>4MB).');
      return;
    }

    const cacheObject = {
      data: data,
      timestamp: timestamp || new Date().toISOString()
    };
    localStorage.setItem(CACHE_KEY, JSON.stringify(cacheObject));
    console.log(`💾 Cache salvo com ${data.length} registros`);
  } catch (e) {
    console.warn('⚠️ Não foi possível salvar no cache:', e.message);
  }
}

function getCachedData(ignoreExpiry = false) {
  try {
    const cached = localStorage.getItem(CACHE_KEY);
    if (!cached) {
      console.log('📭 Nenhum cache encontrado.');
      return null;
    }

    const cacheObject = JSON.parse(cached);

    if (!cacheObject.data || cacheObject.data.length === 0) {
      console.log('📭 Cache vazio.');
      return null;
    }

    if (ignoreExpiry) {
      console.log(`📦 Usando cache (ignorando expiração): ${cacheObject.data.length} registros`);
      return cacheObject;
    }

    const now = new Date();
    const cacheDate = new Date(cacheObject.timestamp);
    const hoursDiff = (now - cacheDate) / (1000 * 60 * 60);

    console.log(`⏱️ Cache tem ${hoursDiff.toFixed(1)} horas de idade`);

    if (hoursDiff <= 0.5) {
      console.log(`✅ Cache válido: ${cacheObject.data.length} registros`);
      return cacheObject;
    }

    console.log('⏰ Cache expirado.');
    return null;
  } catch (e) {
    console.warn('⚠️ Erro ao ler cache:', e);
    return null;
  }
}

function showToast(message, type = 'info') {
  const toast = document.createElement('div');
  const colors = {
    info: '#2980b9',
    success: '#27ae60',
    warning: '#f39c12',
    error: '#e74c3c'
  };

  toast.style.cssText = `
    position: fixed;
    bottom: 24px;
    right: 24px;
    z-index: 9999;
    background: ${colors[type] || colors.info};
    color: #fff;
    border-radius: 12px;
    padding: 14px 22px;
    font-family: 'Inter', sans-serif;
    font-size: 0.85rem;
    font-weight: 600;
    box-shadow: 0 6px 24px rgba(0,0,0,0.3);
    display: flex;
    align-items: center;
    gap: 10px;
    animation: slideUp 0.3s ease;
    max-width: 400px;
  `;

  const icons = {
    info: 'ℹ️',
    success: '✅',
    warning: '⚠️',
    error: '❌'
  };

  toast.innerHTML = `${icons[type] || 'ℹ️'} ${message}`;
  document.body.appendChild(toast);

  setTimeout(() => {
    toast.style.opacity = '0';
    toast.style.transition = 'opacity 0.3s ease';
    setTimeout(() => toast.remove(), 300);
  }, 5000);
}

const style = document.createElement('style');
style.textContent = `
  @keyframes slideUp {
    from { opacity: 0; transform: translateY(20px); }
    to { opacity: 1; transform: translateY(0); }
  }
`;
document.head.appendChild(style);

function onDataLoaded() {
  populateMultiSelectOptions();
  applyFilters();
}

// ============================================================
// NORMALIZAR DADOS
// ============================================================
function normalizeData(rows) {
  return rows.map(row => {
    const get = (...keys) => {
      for (const k of keys) {
        if (row[k] !== undefined && row[k] !== null && row[k] !== '') return row[k];
        const kNorm = norm(k);
        for (const [rk, rv] of Object.entries(row)) {
          if (norm(rk) === kNorm && rv !== undefined && rv !== null && rv !== '') return rv;
        }
      }
      return '';
    };

    const unidadeSolicitante = getUnidade({
      'UNIDADE SOLICITANTE':   get('UNIDADE SOLICITANTE'),
      'UNIDADE DE REFERÊNCIA': get('UNIDADE DE REFERÊNCIA','UNIDADE DE REFERENCIA'),
    });

    const nomeCBO       = get('NOME CBO');
    const especialidade = get('ESPECIALIDADE');
    const cbof          = formatCBO(nomeCBO, especialidade);
    const profissional  = formatProfissional(get('NOME PROFISSIONAL'));
    const tipoAtend     = getTipoAtendimento(get('TIPO DE ATENDIMENTO'));
    const situacao      = (get('SITUAÇÃO','SITUACAO') || '').toUpperCase().trim();
    const operCod       = get('OPERADOR AGENDAMENTO');

    const dataCriacao       = get('DATA CRIAÇÃO DO AGENDAMENTO', 'DATA CRIACAO DO AGENDAMENTO');
    const dataCriacaoParsed = parseDate(dataCriacao);

    const dataAgenda       = get('DATA AGENDA','DATA_AGENDA');
    const dataAgendaParsed = parseDate(dataAgenda);

    const unidadeExec = get('UNIDADE EXECUTANTE');
    const distrito    = getDistrito(unidadeSolicitante);

    let mesLabel = '';
    if (dataAgendaParsed) {
      mesLabel = MESES_PT[dataAgendaParsed.getMonth()] + '/' + dataAgendaParsed.getFullYear();
    }

    return {
      _raw:              row,
      unidadeExecutante: unidadeExec,
      unidadeSolicitante,
      cbo:               cbof,
      especialidade:     especialidade ? titleCase(especialidade.toString()) : '',
      nomeCBO,
      profissional,
      tipoAtendimento:   tipoAtend,
      situacao,
      situacaoLabel:     getSituacaoLabel(situacao),
      operadorCod:       operCod,
      operador:          getOperador(operCod),
      dataCriacao,
      dataCriacaoParsed,
      dataAgenda,
      dataAgendaParsed,
      mesAgendamento:    mesLabel,
      distrito,
      valor:     parseFloat((get('VALOR')||'0').toString().replace(',','.')) || 0,
      quantidade: parseInt(get('QUANTIDADE')||'1') || 1,
    };
  }).filter(r => r.unidadeExecutante || r.profissional);
}

// ============================================================
// POPULAR MULTI SELECTS
// ============================================================
function populateMultiSelectOptions() {
  const prestadores    = [...new Set(allData.map(r => r.unidadeExecutante).filter(Boolean))].sort();
  const especialidades = [...new Set(allData.map(r => r.cbo).filter(Boolean))].sort();
  const profissionais  = [...new Set(allData.map(r => r.profissional).filter(Boolean))].sort();
  const unidades       = [...new Set(allData.map(r => r.unidadeSolicitante).filter(Boolean))].sort();
  const distritos      = [...new Set(allData.map(r => r.distrito).filter(Boolean))].sort();
  const tiposServico   = ['Primeira Consulta', 'Retorno'];

  const mesesSet = {};
  allData.forEach(r => {
    if (r.dataAgendaParsed && r.mesAgendamento) {
      const d = r.dataAgendaParsed;
      const key = d.getFullYear() * 100 + d.getMonth();
      mesesSet[key] = r.mesAgendamento;
    }
  });
  const mesesOrdenados = Object.entries(mesesSet).sort((a,b) => +a[0] - +b[0]).map(e => e[1]);

  initMultiSelect('multiSelectMes', mesesOrdenados);
  initMultiSelect('multiSelectPrestador', prestadores);
  initMultiSelect('multiSelectEspecialidade', especialidades);
  initMultiSelect('multiSelectTipoServico', tiposServico);
  initMultiSelect('multiSelectProfissional', profissionais);
  initMultiSelect('multiSelectUnidade', unidades);
  initMultiSelect('multiSelectDistrito', distritos);

  Object.keys(multiSelectState).forEach(key => {
    multiSelectState[key]._allValues = [...multiSelectState[key].values];
    const placeholder = document.querySelector(`#${key} .multi-select-placeholder`);
    if (placeholder) {
      if (key === 'multiSelectMes') placeholder.setAttribute('data-default', 'Todos os meses');
      else if (key === 'multiSelectTipoServico') placeholder.setAttribute('data-default', 'Todos');
      else if (key === 'multiSelectEspecialidade' || key === 'multiSelectUnidade') placeholder.setAttribute('data-default', 'Todas');
      else placeholder.setAttribute('data-default', 'Todos');
    }
  });
}

// ============================================================
// APLICAR FILTROS
// ============================================================
function applyFilters() {
  const prestadoresSelecionados    = getMultiSelectValues('multiSelectPrestador');
  const especialidadesSelecionadas = getMultiSelectValues('multiSelectEspecialidade');
  const tiposServicoSelecionados   = getMultiSelectValues('multiSelectTipoServico');
  const profissionaisSelecionados  = getMultiSelectValues('multiSelectProfissional');
  const mesesSelecionados          = getMultiSelectValues('multiSelectMes');
  const unidadesSelecionadas       = getMultiSelectValues('multiSelectUnidade');
  const distritosSelecionados      = getMultiSelectValues('multiSelectDistrito');
  const dataCriacaoSelecionada     = window._fpInicio ? window._fpInicio.selectedDates[0] : null;

  filteredData = allData.filter(r => {
    if (prestadoresSelecionados.length > 0 && !prestadoresSelecionados.includes(r.unidadeExecutante)) return false;
    if (especialidadesSelecionadas.length > 0 && !especialidadesSelecionadas.includes(r.cbo)) return false;
    if (tiposServicoSelecionados.length > 0 && !tiposServicoSelecionados.includes(r.tipoAtendimento)) return false;
    if (profissionaisSelecionados.length > 0 && !profissionaisSelecionados.includes(r.profissional)) return false;
    if (mesesSelecionados.length > 0 && !mesesSelecionados.includes(r.mesAgendamento)) return false;
    if (unidadesSelecionadas.length > 0 && !unidadesSelecionadas.includes(r.unidadeSolicitante)) return false;
    if (distritosSelecionados.length > 0 && !distritosSelecionados.includes(r.distrito)) return false;

    if (dataCriacaoSelecionada) {
      if (!r.dataCriacaoParsed) return false;
      if (!compareDate(dataCriacaoSelecionada, r.dataCriacaoParsed)) return false;
    }
    return true;
  });

  updateKPIs();
  buildTableData();
  buildTableDataResumido();
  currentPage = 1;
  currentPageResumido = 1;

  const activeTab = document.querySelector('.tab-btn.active');
  if (activeTab) {
    renderChartsForTab(activeTab.getAttribute('data-tab'));
  } else {
    renderChartsForTab('visao-geral');
  }

  const tabelaPanel = document.getElementById('tab-tabela');
  if (tabelaPanel && tabelaPanel.classList.contains('active')) {
    renderTable();
  }

  const resumidoPanel = document.getElementById('tab-profissionais-atendimento');
  if (resumidoPanel && resumidoPanel.classList.contains('active')) {
    renderTableResumido();
  }
}

function clearFilters() {
  Object.keys(multiSelectState).forEach(key => {
    multiSelectState[key].selected = [];
    multiSelectState[key].values = [...multiSelectState[key]._allValues];
    renderMultiSelect(key);
  });
  if (window._fpInicio) window._fpInicio.clear();
  applyFilters();
}

// ============================================================
// KPIS
// ============================================================
function updateKPIs() {
  const rec = filteredData.filter(r => r.situacao === 'REC').length;
  const fal = filteredData.filter(r => r.situacao === 'FAL').length;
  const can = filteredData.filter(r => r.situacao === 'CAN').length;
  const tra = filteredData.filter(r => r.situacao === 'TRA').length;
  const atd = filteredData.filter(r => r.situacao === 'ATD').length;
  const agendados = rec + fal + atd;

  animateCount('kpiAgendados', agendados);
  animateCount('kpiRecepcionados', rec);
  animateCount('kpiAtendidos', atd);
  animateCount('kpiFaltosos', fal);
  animateCount('kpiCancelados', can);
  animateCount('kpiTransferidos', tra);
}

function animateCount(id, target) {
  const el = document.getElementById(id);
  if (!el) return;
  const current = parseInt(el.textContent.replace(/\D/g,'')) || 0;
  if (current === target) { el.textContent = fmt(target); return; }
  const step = Math.max(1, Math.round(Math.abs(target - current) / 20));
  let val = current;
  const inc = target > current ? step : -step;
  const timer = setInterval(() => {
    val += inc;
    if ((inc > 0 && val >= target) || (inc < 0 && val <= target)) { val = target; clearInterval(timer); }
    el.textContent = fmt(val);
  }, 16);
}

// ============================================================
// HELPERS DE GRÁFICO
// ============================================================
function countBy(data, keyFn) {
  const map = {};
  data.forEach(r => { const k = keyFn(r); if (!k) return; map[k] = (map[k]||0) + 1; });
  return map;
}

function sortedEntries(obj, limit = 0) {
  let entries = Object.entries(obj).sort((a,b) => b[1]-a[1]);
  if (limit > 0) entries = entries.slice(0, limit);
  return entries;
}

function destroyChart(ref) { if (ref) { try { ref.destroy(); } catch(e) {} } }

// ============================================================
// GRÁFICOS - VISÃO GERAL
// ============================================================

function renderChartDistrito() {
  const ctx = document.getElementById('chartDistrito')?.getContext('2d');
  if (!ctx) return;
  const counts = countBy(filteredData, r => r.distrito);
  const entries = sortedEntries(counts);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  const total = data.reduce((a,b) => a+b, 0);
  const colors = labels.map((_, i) => PALETTE_AZUL[i % PALETTE_AZUL.length]);

  destroyChart(chartDistrito);
  chartDistrito = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Agendamentos',
        data,
        backgroundColor: colors.map(c => c + 'dd'),
        borderColor: colors.map(c => c),
        borderWidth: 2,
        borderRadius: 8,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          ...TOOLTIP_BASE,
          callbacks: {
            label: ctx => ` ${fmt(ctx.raw)} agendamentos`,
            afterLabel: ctx => ` ${total > 0 ? (ctx.raw/total*100).toFixed(1) : 0}% do total`
          }
        },
        datalabels: {
          anchor: 'center',
          align: 'center',
          color: '#fff',
          textStrokeColor: 'rgba(0,0,0,0.30)',
          textStrokeWidth: 2,
          font: { family: 'Inter', size: 13, weight: 'bold' },
          formatter: val => val > 0 ? fmt(val) : ''
        }
      },
      scales: {
        x: {
          ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 30 },
          grid: { display: false }
        },
        y: {
          beginAtZero: true,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' },
          grid: { display: false }
        }
      }
    }
  });
}

function renderChartTipoAtendimento() {
  const ctx = document.getElementById('chartTipoAtendimento')?.getContext('2d');
  if (!ctx) return;
  const pc = filteredData.filter(r => r.tipoAtendimento === 'Primeira Consulta').length;
  const ret = filteredData.filter(r => r.tipoAtendimento === 'Retorno').length;

  destroyChart(chartTipoAtendimento);
  chartTipoAtendimento = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['1ª Consulta', 'Retorno'],
      datasets: [{
        label: 'Quantidade',
        data: [pc, ret],
        backgroundColor: ['#1e3a5f', '#4a90d9'],
        borderColor: ['#1e3a5f', '#4a90d9'],
        borderWidth: 2,
        borderRadius: 8,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { ...TOOLTIP_BASE, callbacks: { label: ctx => ` ${fmt(ctx.raw)} agendamentos` } },
        datalabels: {
          anchor: 'center',
          align: 'center',
          color: '#fff',
          font: { family: 'Inter', size: 18, weight: 'bold' },
          formatter: val => val > 0 ? fmt(val) : ''
        }
      },
      scales: {
        x: {
          ticks: { font: { family: 'Inter', size: 15, weight: '700' }, color: '#1e3a5f' },
          grid: { display: false }
        },
        y: {
          beginAtZero: true,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' },
          grid: { display: false }
        }
      }
    }
  });
}

function renderChartEspecialidade() {
  const ctx = document.getElementById('chartEspecialidade')?.getContext('2d');
  if (!ctx) return;
  const counts = countBy(filteredData, r => r.cbo);
  const entries = sortedEntries(counts, 15);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  const colors = labels.map((_, i) => PALETTE_AZUL[i % PALETTE_AZUL.length]);

  destroyChart(chartEspecialidade);
  chartEspecialidade = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Agendamentos',
        data,
        backgroundColor: colors.map(c => c + 'dd'),
        borderColor: colors.map(c => c),
        borderWidth: 2,
        borderRadius: 5,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: 'y',
      plugins: {
        legend: { display: false },
        tooltip: TOOLTIP_BASE,
        datalabels: {
          anchor: 'end',
          align: 'end',
          color: '#1e3a5f',
          font: { family: 'Inter', size: 13, weight: '800' },
          formatter: val => fmt(val)
        }
      },
      layout: { padding: { right: 54 } },
      scales: {
        y: {
          ticks: {
            font: { family: 'Inter', size: 10 },
            color: '#3d5166',
            callback(val) {
              const label = this.getLabelForValue(val);
              return label && label.length > 24 ? label.substring(0,22)+'…' : label;
            }
          },
          grid: { display: false }
        },
        x: {
          beginAtZero: true,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' },
          grid: { display: false }
        }
      }
    }
  });
}

function renderChartPrestador() {
  const ctx = document.getElementById('chartPrestador')?.getContext('2d');
  if (!ctx) return;
  const counts = countBy(filteredData, r => r.unidadeExecutante);
  const entries = sortedEntries(counts, 10);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  const colors = labels.map((_, i) => PALETTE_AZUL[(i+3) % PALETTE_AZUL.length]);

  destroyChart(chartPrestador);
  chartPrestador = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Agendamentos',
        data,
        backgroundColor: colors.map(c => c + 'dd'),
        borderColor: colors.map(c => c),
        borderWidth: 2,
        borderRadius: 6,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      indexAxis: 'y',
      plugins: {
        legend: { display: false },
        tooltip: TOOLTIP_BASE,
        datalabels: {
          anchor: 'end',
          align: 'end',
          color: '#1e3a5f',
          font: { family: 'Inter', size: 13, weight: '800' },
          formatter: val => fmt(val)
        }
      },
      layout: { padding: { right: 54 } },
      scales: {
        y: {
          ticks: {
            font: { family: 'Inter', size: 9 },
            color: '#3d5166',
            callback(val) {
              const label = this.getLabelForValue(val);
              return label && label.length > 26 ? label.substring(0,24)+'…' : label;
            }
          },
          grid: { display: false }
        },
        x: {
          beginAtZero: true,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' },
          grid: { display: false }
        }
      }
    }
  });
}

function renderChartSituacao() {
  const ctx = document.getElementById('chartSituacao')?.getContext('2d');
  if (!ctx) return;
  const age = filteredData.filter(r => r.situacao === 'AGE').length;
  const rec = filteredData.filter(r => r.situacao === 'REC').length;
  const atd = filteredData.filter(r => r.situacao === 'ATD').length;
  const fal = filteredData.filter(r => r.situacao === 'FAL').length;
  const can = filteredData.filter(r => r.situacao === 'CAN').length;
  const tra = filteredData.filter(r => r.situacao === 'TRA').length;
  const labels = ['Agendados', 'Recepcionados', 'Atendidos', 'Faltosos', 'Cancelados', 'Transferidos'];
  const data = [age, rec, atd, fal, can, tra];
  const total = data.reduce((a,b) => a+b, 0);
  const coresAzul = ['#0a1628', '#1a3256', '#2d5494', '#4a90d9', '#6aaff0', '#8fc7f5'];

  destroyChart(chartSituacao);
  chartSituacao = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Quantidade',
        data,
        backgroundColor: coresAzul.map(c => c + 'cc'),
        borderColor: coresAzul,
        borderWidth: 2,
        borderRadius: 8,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          ...TOOLTIP_BASE,
          callbacks: {
            label: ctx => ` ${fmt(ctx.raw)} consultas`,
            afterLabel: ctx => ` ${total > 0 ? (ctx.raw/total*100).toFixed(1) : 0}% do total`
          }
        },
        datalabels: {
          anchor: 'center',
          align: 'center',
          color: '#fff',
          textStrokeColor: 'rgba(0,0,0,0.30)',
          textStrokeWidth: 2,
          font: { family: 'Inter', size: 13, weight: 'bold' },
          formatter: val => val > 0 ? fmt(val) : ''
        }
      },
      scales: {
        x: {
          ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 35 },
          grid: { display: false }
        },
        y: {
          beginAtZero: true,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' },
          grid: { display: false }
        }
      }
    }
  });
}

function renderChartMeses() {
  const ctx = document.getElementById('chartMeses')?.getContext('2d');
  if (!ctx) return;
  const counts = countBy(filteredData, r => r.mesAgendamento);
  const entries = sortedEntries(counts);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  const total = data.reduce((a,b) => a+b, 0);
  const coresAzul = [
    '#0a1628', '#0f1f3a', '#142844', '#1a3256',
    '#1e3a5f', '#254a7a', '#2d5494', '#3a6bb0',
    '#4a90d9', '#6aaff0', '#8fc7f5', '#b5dffa'
  ];

  destroyChart(chartMeses);
  chartMeses = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: labels.map((_, i) => coresAzul[i % coresAzul.length]),
        borderColor: '#fff',
        borderWidth: 3,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'right',
          labels: {
            font: { family: 'Inter', size: 11, weight: '600' },
            color: '#3d5166',
            padding: 14,
          }
        },
        tooltip: {
          ...TOOLTIP_BASE,
          callbacks: {
            label: ctx => {
              const value = ctx.raw;
              const pct = total > 0 ? (value/total*100).toFixed(1) : 0;
              return ` ${fmt(value)} agendamentos (${pct}%)`;
            }
          }
        },
        datalabels: {
          color: '#fff',
          font: { family: 'Inter', size: 11, weight: 'bold' },
          formatter: (val) => total > 0 ? (val/total*100).toFixed(0) + '%' : ''
        },
        centerText: {
          enabled: true,
          value: fmt(total),
          label: 'Total',
          fontSize: 24,
          valueColor: '#1e3a5f',
          labelColor: '#7a8fa6'
        }
      }
    }
  });
}

// ============================================================
// GRÁFICOS - AGENDAMENTOS POR DISTRITO
// ============================================================

function renderChartPrimeiraConsultaDistrito() {
  const ctx = document.getElementById('chartPrimeiraConsultaDistrito')?.getContext('2d');
  if (!ctx) return;
  const pcData = filteredData.filter(r => r.tipoAtendimento === 'Primeira Consulta');
  const counts = countBy(pcData, r => r.distrito);
  const entries = sortedEntries(counts);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  const colors = labels.map((_, i) => PALETTE_ROSA[i % PALETTE_ROSA.length]);

  destroyChart(chartPrimeiraConsultaDistrito);
  chartPrimeiraConsultaDistrito = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '1ª Consulta',
        data,
        backgroundColor: colors.map(c => c + 'dd'),
        borderColor: colors.map(c => c),
        borderWidth: 2,
        borderRadius: 8,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: TOOLTIP_BASE,
        datalabels: {
          anchor: 'center',
          align: 'center',
          color: '#fff',
          font: { family: 'Inter', size: 13, weight: 'bold' },
          formatter: val => val > 0 ? fmt(val) : ''
        }
      },
      scales: {
        x: {
          ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 30 },
          grid: { display: false }
        },
        y: {
          beginAtZero: true,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' },
          grid: { display: false }
        }
      }
    }
  });
}

function renderChartRetornoDistrito() {
  const ctx = document.getElementById('chartRetornoDistrito')?.getContext('2d');
  if (!ctx) return;
  const retData = filteredData.filter(r => r.tipoAtendimento === 'Retorno');
  const counts = countBy(retData, r => r.distrito);
  const entries = sortedEntries(counts);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  const colors = labels.map((_, i) => PALETTE_ROSA[(i+3) % PALETTE_ROSA.length]);

  destroyChart(chartRetornoDistrito);
  chartRetornoDistrito = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Retorno',
        data,
        backgroundColor: colors.map(c => c + 'dd'),
        borderColor: colors.map(c => c),
        borderWidth: 2,
        borderRadius: 8,
        borderSkipped: false,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: TOOLTIP_BASE,
        datalabels: {
          anchor: 'center',
          align: 'center',
          color: '#fff',
          font: { family: 'Inter', size: 13, weight: 'bold' },
          formatter: val => val > 0 ? fmt(val) : ''
        }
      },
      scales: {
        x: {
          ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 30 },
          grid: { display: false }
        },
        y: {
          beginAtZero: true,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' },
          grid: { display: false }
        }
      }
    }
  });
}

function renderChartComparativoDistrito() {
  const ctx = document.getElementById('chartComparativoDistrito')?.getContext('2d');
  if (!ctx) return;
  const distritos = [...new Set(filteredData.map(r => r.distrito).filter(Boolean))].sort();
  const pcCounts = distritos.map(d => filteredData.filter(r => r.distrito === d && r.tipoAtendimento === 'Primeira Consulta').length);
  const retCounts = distritos.map(d => filteredData.filter(r => r.distrito === d && r.tipoAtendimento === 'Retorno').length);

  destroyChart(chartComparativoDistrito);
  chartComparativoDistrito = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: distritos,
      datasets: [
        {
          label: '1ª Consulta',
          data: pcCounts,
          backgroundColor: '#8a2a48',
          borderColor: '#8a2a48',
          borderWidth: 2,
          borderRadius: 6,
        },
        {
          label: 'Retorno',
          data: retCounts,
          backgroundColor: '#f08cb0',
          borderColor: '#f08cb0',
          borderWidth: 2,
          borderRadius: 6,
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'top',
          labels: {
            font: { family: 'Inter', size: 12, weight: '600' },
            color: '#3d5166',
            usePointStyle: true,
            pointStyleWidth: 10,
          }
        },
        tooltip: TOOLTIP_BASE,
        datalabels: {
          anchor: 'center',
          align: 'center',
          color: '#fff',
          font: { family: 'Inter', size: 11, weight: 'bold' },
          formatter: val => val > 0 ? fmt(val) : ''
        }
      },
      scales: {
        x: {
          ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 35 },
          grid: { display: false }
        },
        y: {
          beginAtZero: true,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' },
          grid: { display: false }
        }
      }
    }
  });
}

function renderChartDistritoRosca() {
  const ctx = document.getElementById('chartDistritoRosca')?.getContext('2d');
  if (!ctx) return;
  const counts = countBy(filteredData, r => r.distrito);
  const entries = sortedEntries(counts);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  const total = data.reduce((a,b) => a+b, 0);
  const colors = labels.map((_, i) => PALETTE_ROSA[i % PALETTE_ROSA.length]);

  destroyChart(chartDistritoRosca);
  chartDistritoRosca = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: colors,
        borderColor: '#fff',
        borderWidth: 3,
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          position: 'right',
          labels: {
            font: { family: 'Inter', size: 11, weight: '600' },
            color: '#3d5166',
            padding: 14,
          }
        },
        tooltip: {
          ...TOOLTIP_BASE,
          callbacks: {
            label: ctx => ` ${fmt(ctx.raw)} agendamentos (${total > 0 ? (ctx.raw/total*100).toFixed(1) : 0}%)`
          }
        },
        datalabels: {
          color: '#fff',
          font: { family: 'Inter', size: 11, weight: 'bold' },
          formatter: (val) => total > 0 ? (val/total*100).toFixed(0) + '%' : ''
        },
        centerText: {
          enabled: true,
          value: fmt(total),
          label: 'Total',
          fontSize: 22,
          valueColor: '#8a2a48',
          labelColor: '#c93d70'
        }
      }
    }
  });
}

// ============================================================
// GRÁFICOS - ABSENTEÍSMO
// ============================================================

function renderChartAbsenteismoEsp() {
  const ctx = document.getElementById('chartAbsenteismoEsp')?.getContext('2d');
  if (!ctx) return;
  const map = {};
  filteredData.forEach(r => {
    const key = r.cbo || '–';
    if (!map[key]) map[key] = { fal: 0, rec: 0, can: 0, atd: 0 };
    if (r.situacao === 'FAL') map[key].fal++;
    else if (r.situacao === 'REC') map[key].rec++;
    else if (r.situacao === 'CAN') map[key].can++;
    else if (r.situacao === 'ATD') map[key].atd++;
  });
  const entries = Object.entries(map)
    .filter(([,v]) => (v.rec + v.fal + v.can + v.atd) > 0)
    .map(([k,v]) => ({ label: k, pct: parseFloat((v.fal / (v.rec + v.fal + v.can + v.atd) * 100).toFixed(1)), fal: v.fal, total: v.rec + v.fal + v.can + v.atd }))
    .sort((a,b) => b.pct - a.pct).slice(0, 15);
  const labels = entries.map(e => e.label);
  const data = entries.map(e => e.pct);
  const bgs = data.map((v, i) => PALETTE_ABSENTEISMO[i % PALETTE_ABSENTEISMO.length] + 'cc');
  const borders = data.map((v, i) => PALETTE_ABSENTEISMO[i % PALETTE_ABSENTEISMO.length]);
  destroyChart(chartAbsenteismoEsp);
  chartAbsenteismoEsp = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: '% Absenteísmo', data, backgroundColor: bgs, borderColor: borders, borderWidth: 2, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: {
        legend: { display: false },
        tooltip: { ...TOOLTIP_BASE, callbacks: { label: (ctx) => { const e = entries[ctx.dataIndex]; return [` ${ctx.raw}% de absenteísmo`, ` ${fmt(e.fal)} faltosos de ${fmt(e.total)} registros`]; } } },
        datalabels: { anchor: 'end', align: 'end', clamp: true, color: '#3d5166', font: { family: 'Inter', size: 10, weight: 'bold' }, formatter: val => val + '%' }
      },
      layout: { padding: { right: 44 } },
      scales: {
        y: { ticks: { font: { family: 'Inter', size: 10 }, color: '#3d5166', callback(val) { const label = this.getLabelForValue(val); return label && label.length > 22 ? label.substring(0,20)+'…' : label; } }, grid: { display: false } },
        x: { beginAtZero: true, max: 100, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6', callback: v => v + '%' }, grid: { display: false } }
      }
    }
  });
}

function renderChartAbsenteismoDist() {
  const ctx = document.getElementById('chartAbsenteismoDist')?.getContext('2d');
  if (!ctx) return;
  const map = {};
  filteredData.forEach(r => {
    const key = r.distrito || 'OUTROS';
    if (!map[key]) map[key] = { fal: 0, rec: 0, can: 0, atd: 0 };
    if (r.situacao === 'FAL') map[key].fal++;
    else if (r.situacao === 'REC') map[key].rec++;
    else if (r.situacao === 'CAN') map[key].can++;
    else if (r.situacao === 'ATD') map[key].atd++;
  });
  const entries = Object.entries(map)
    .filter(([,v]) => (v.rec + v.fal + v.can + v.atd) > 0)
    .map(([k,v]) => ({ label: k, pct: parseFloat((v.fal / (v.rec + v.fal + v.can + v.atd) * 100).toFixed(1)), fal: v.fal, total: v.rec + v.fal + v.can + v.atd }))
    .sort((a,b) => b.pct - a.pct);
  const labels = entries.map(e => e.label);
  const data = entries.map(e => e.pct);
  const bgs = data.map((v, i) => PALETTE_ABSENTEISMO[i % PALETTE_ABSENTEISMO.length] + 'cc');
  const borders = data.map((v, i) => PALETTE_ABSENTEISMO[i % PALETTE_ABSENTEISMO.length]);
  destroyChart(chartAbsenteismoDist);
  chartAbsenteismoDist = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: '% Absenteísmo', data, backgroundColor: bgs, borderColor: borders, borderWidth: 2, borderRadius: 7, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: { ...TOOLTIP_BASE, callbacks: { label: (ctx) => { const e = entries[ctx.dataIndex]; return [` ${ctx.raw}% de absenteísmo`, ` ${fmt(e.fal)} faltosos de ${fmt(e.total)} registros`]; } } },
        datalabels: { anchor: 'end', align: 'end', clamp: true, color: '#1a2a3a', font: { family: 'Inter', size: 11, weight: 'bold' }, formatter: val => val + '%' }
      },
      layout: { padding: { top: 28 } },
      scales: {
        x: { ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 35 }, grid: { display: false } },
        y: { beginAtZero: true, max: 100, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6', callback: v => v + '%' }, grid: { display: false } }
      }
    }
  });
}

function renderChartAbsenteismoMensal() {
  const ctx = document.getElementById('chartAbsenteismoMensal')?.getContext('2d');
  if (!ctx) return;
  const map = {};
  filteredData.forEach(r => {
    const key = r.mesAgendamento || '–';
    if (!map[key]) map[key] = { fal: 0, total: 0 };
    map[key].total++;
    if (r.situacao === 'FAL') map[key].fal++;
  });
  const mesesOrdenados = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const sortedKeys = Object.keys(map).sort((a, b) => {
    const [mesA, anoA] = a.split('/');
    const [mesB, anoB] = b.split('/');
    const indexMesA = mesesOrdenados.indexOf(mesA);
    const indexMesB = mesesOrdenados.indexOf(mesB);
    if (anoA !== anoB) return parseInt(anoA) - parseInt(anoB);
    return indexMesA - indexMesB;
  });
  const labels = sortedKeys;
  const data = sortedKeys.map(k => map[k].total > 0 ? parseFloat((map[k].fal / map[k].total * 100).toFixed(1)) : 0);
  destroyChart(chartAbsenteismoMensal);
  chartAbsenteismoMensal = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets: [{ label: '% Absenteísmo', data, borderColor: '#e74c3c', backgroundColor: 'rgba(231,76,60,0.1)', borderWidth: 3, pointRadius: 6, pointBackgroundColor: '#e74c3c', tension: 0.4, fill: true }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { ...TOOLTIP_BASE, callbacks: { label: ctx => ` ${ctx.raw}% de absenteísmo` } }, datalabels: { anchor: 'end', align: 'end', color: '#c0392b', font: { family: 'Inter', size: 11, weight: 'bold' }, formatter: val => val + '%' } },
      scales: {
        x: { ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 30 }, grid: { display: false } },
        y: { beginAtZero: true, max: 100, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6', callback: v => v + '%' }, grid: { display: false } }
      }
    }
  });
}

function renderChartAbsenteismoPrestador() {
  const ctx = document.getElementById('chartAbsenteismoPrestador')?.getContext('2d');
  if (!ctx) return;
  const map = {};
  filteredData.forEach(r => {
    const key = r.unidadeExecutante || '–';
    if (!map[key]) map[key] = { fal: 0, rec: 0, can: 0, atd: 0 };
    if (r.situacao === 'FAL') map[key].fal++;
    else if (r.situacao === 'REC') map[key].rec++;
    else if (r.situacao === 'CAN') map[key].can++;
    else if (r.situacao === 'ATD') map[key].atd++;
  });
  const entries = Object.entries(map)
    .filter(([,v]) => (v.rec + v.fal + v.can + v.atd) > 0)
    .map(([k,v]) => ({ label: k, pct: parseFloat((v.fal / (v.rec + v.fal + v.can + v.atd) * 100).toFixed(1)), fal: v.fal, total: v.rec + v.fal + v.can + v.atd }))
    .sort((a,b) => b.pct - a.pct).slice(0, 10);
  const labels = entries.map(e => e.label);
  const data = entries.map(e => e.pct);
  destroyChart(chartAbsenteismoPrestador);
  chartAbsenteismoPrestador = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: '% Absenteísmo', data, backgroundColor: PALETTE_ABSENTEISMO.map(c => c + 'cc'), borderColor: PALETTE_ABSENTEISMO, borderWidth: 2, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: TOOLTIP_BASE, datalabels: { anchor: 'end', align: 'end', color: '#3d5166', font: { family: 'Inter', size: 10, weight: 'bold' }, formatter: val => val + '%' } },
      layout: { padding: { right: 44 } },
      scales: {
        y: { ticks: { font: { family: 'Inter', size: 10 }, color: '#3d5166' }, grid: { display: false } },
        x: { beginAtZero: true, max: 100, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6', callback: v => v + '%' }, grid: { display: false } }
      }
    }
  });
}

// ============================================================
// GRÁFICOS - RECEPCIONADOS
// ============================================================

function renderChartRecepcionadosDistrito() {
  const ctx = document.getElementById('chartRecepcionadosDistrito')?.getContext('2d');
  if (!ctx) return;
  const recData = filteredData.filter(r => r.situacao === 'REC');
  const counts = countBy(recData, r => r.distrito);
  const entries = sortedEntries(counts);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  destroyChart(chartRecepcionadosDistrito);
  chartRecepcionadosDistrito = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Recepcionados', data, backgroundColor: PALETTE_RECEPCIONADOS.map(c => c + 'bb'), borderColor: PALETTE_RECEPCIONADOS, borderWidth: 2, borderRadius: 8, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: TOOLTIP_BASE, datalabels: { anchor: 'center', align: 'center', color: '#fff', font: { family: 'Inter', size: 13, weight: 'bold' }, formatter: val => val > 0 ? fmt(val) : '' } },
      scales: { x: { ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 30 }, grid: { display: false } }, y: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

function renderChartRecepcionadosEspecialidade() {
  const ctx = document.getElementById('chartRecepcionadosEspecialidade')?.getContext('2d');
  if (!ctx) return;
  const recData = filteredData.filter(r => r.situacao === 'REC');
  const counts = countBy(recData, r => r.cbo);
  const entries = sortedEntries(counts, 15);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  destroyChart(chartRecepcionadosEspecialidade);
  chartRecepcionadosEspecialidade = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Recepcionados', data, backgroundColor: PALETTE_RECEPCIONADOS.map(c => c + 'dd'), borderColor: PALETTE_RECEPCIONADOS, borderWidth: 2, borderRadius: 5, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: TOOLTIP_BASE, datalabels: { anchor: 'end', align: 'end', color: '#3d5166', font: { family: 'Inter', size: 13, weight: '800' }, formatter: val => fmt(val) } },
      layout: { padding: { right: 54 } },
      scales: { y: { ticks: { font: { family: 'Inter', size: 10 }, color: '#3d5166' }, grid: { display: false } }, x: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

function renderChartRecepcionadosPrestador() {
  const ctx = document.getElementById('chartRecepcionadosPrestador')?.getContext('2d');
  if (!ctx) return;
  const recData = filteredData.filter(r => r.situacao === 'REC');
  const counts = countBy(recData, r => r.unidadeExecutante);
  const entries = sortedEntries(counts, 10);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  destroyChart(chartRecepcionadosPrestador);
  chartRecepcionadosPrestador = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Recepcionados', data, backgroundColor: PALETTE_RECEPCIONADOS.map(c => c + 'dd'), borderColor: PALETTE_RECEPCIONADOS, borderWidth: 2, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: TOOLTIP_BASE, datalabels: { anchor: 'end', align: 'end', color: '#3d5166', font: { family: 'Inter', size: 13, weight: '800' }, formatter: val => fmt(val) } },
      layout: { padding: { right: 54 } },
      scales: { y: { ticks: { font: { family: 'Inter', size: 9 }, color: '#3d5166' }, grid: { display: false } }, x: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

function renderChartRecepcionadosMensal() {
  const ctx = document.getElementById('chartRecepcionadosMensal')?.getContext('2d');
  if (!ctx) return;
  const map = {};
  filteredData.forEach(r => {
    const key = r.mesAgendamento || '–';
    if (!map[key]) map[key] = 0;
    if (r.situacao === 'REC') map[key]++;
  });
  const mesesOrdenados = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const sortedKeys = Object.keys(map).sort((a, b) => {
    const [mesA, anoA] = a.split('/');
    const [mesB, anoB] = b.split('/');
    const indexMesA = mesesOrdenados.indexOf(mesA);
    const indexMesB = mesesOrdenados.indexOf(mesB);
    if (anoA !== anoB) return parseInt(anoA) - parseInt(anoB);
    return indexMesA - indexMesB;
  });
  const labels = sortedKeys;
  const data = sortedKeys.map(k => map[k]);
  destroyChart(chartRecepcionadosMensal);
  chartRecepcionadosMensal = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets: [{ label: 'Recepcionados', data, borderColor: '#27ae60', backgroundColor: 'rgba(39,174,96,0.1)', borderWidth: 3, pointRadius: 6, pointBackgroundColor: '#27ae60', tension: 0.4, fill: true }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { ...TOOLTIP_BASE, callbacks: { label: ctx => ` ${fmt(ctx.raw)} recepcionados` } }, datalabels: { anchor: 'end', align: 'end', color: '#27ae60', font: { family: 'Inter', size: 11, weight: 'bold' }, formatter: val => fmt(val) } },
      scales: { x: { ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 30 }, grid: { display: false } }, y: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

// ============================================================
// GRÁFICOS - ATENDIDOS
// ============================================================

function renderChartAtendidosDistrito() {
  const ctx = document.getElementById('chartAtendidosDistrito')?.getContext('2d');
  if (!ctx) return;
  const atdData = filteredData.filter(r => r.situacao === 'ATD');
  const counts = countBy(atdData, r => r.distrito);
  const entries = sortedEntries(counts);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  destroyChart(chartAtendidosDistrito);
  chartAtendidosDistrito = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Atendidos', data, backgroundColor: PALETTE_ATENDIDOS.map(c => c + 'bb'), borderColor: PALETTE_ATENDIDOS, borderWidth: 2, borderRadius: 8, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: TOOLTIP_BASE, datalabels: { anchor: 'center', align: 'center', color: '#fff', font: { family: 'Inter', size: 13, weight: 'bold' }, formatter: val => val > 0 ? fmt(val) : '' } },
      scales: { x: { ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 30 }, grid: { display: false } }, y: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

function renderChartAtendidosEspecialidade() {
  const ctx = document.getElementById('chartAtendidosEspecialidade')?.getContext('2d');
  if (!ctx) return;
  const atdData = filteredData.filter(r => r.situacao === 'ATD');
  const counts = countBy(atdData, r => r.cbo);
  const entries = sortedEntries(counts, 15);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  destroyChart(chartAtendidosEspecialidade);
  chartAtendidosEspecialidade = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Atendidos', data, backgroundColor: PALETTE_ATENDIDOS.map(c => c + 'dd'), borderColor: PALETTE_ATENDIDOS, borderWidth: 2, borderRadius: 5, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: TOOLTIP_BASE, datalabels: { anchor: 'end', align: 'end', color: '#3d5166', font: { family: 'Inter', size: 13, weight: '800' }, formatter: val => fmt(val) } },
      layout: { padding: { right: 54 } },
      scales: { y: { ticks: { font: { family: 'Inter', size: 10 }, color: '#3d5166' }, grid: { display: false } }, x: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

function renderChartAtendidosPrestador() {
  const ctx = document.getElementById('chartAtendidosPrestador')?.getContext('2d');
  if (!ctx) return;
  const atdData = filteredData.filter(r => r.situacao === 'ATD');
  const counts = countBy(atdData, r => r.unidadeExecutante);
  const entries = sortedEntries(counts, 10);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  destroyChart(chartAtendidosPrestador);
  chartAtendidosPrestador = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Atendidos', data, backgroundColor: PALETTE_ATENDIDOS.map(c => c + 'dd'), borderColor: PALETTE_ATENDIDOS, borderWidth: 2, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: TOOLTIP_BASE, datalabels: { anchor: 'end', align: 'end', color: '#3d5166', font: { family: 'Inter', size: 13, weight: '800' }, formatter: val => fmt(val) } },
      layout: { padding: { right: 54 } },
      scales: { y: { ticks: { font: { family: 'Inter', size: 9 }, color: '#3d5166' }, grid: { display: false } }, x: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

function renderChartAtendidosMensal() {
  const ctx = document.getElementById('chartAtendidosMensal')?.getContext('2d');
  if (!ctx) return;
  const map = {};
  filteredData.forEach(r => {
    const key = r.mesAgendamento || '–';
    if (!map[key]) map[key] = 0;
    if (r.situacao === 'ATD') map[key]++;
  });
  const mesesOrdenados = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const sortedKeys = Object.keys(map).sort((a, b) => {
    const [mesA, anoA] = a.split('/');
    const [mesB, anoB] = b.split('/');
    const indexMesA = mesesOrdenados.indexOf(mesA);
    const indexMesB = mesesOrdenados.indexOf(mesB);
    if (anoA !== anoB) return parseInt(anoA) - parseInt(anoB);
    return indexMesA - indexMesB;
  });
  const labels = sortedKeys;
  const data = sortedKeys.map(k => map[k]);
  destroyChart(chartAtendidosMensal);
  chartAtendidosMensal = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets: [{ label: 'Atendidos', data, borderColor: '#2980b9', backgroundColor: 'rgba(41,128,185,0.1)', borderWidth: 3, pointRadius: 6, pointBackgroundColor: '#2980b9', tension: 0.4, fill: true }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { ...TOOLTIP_BASE, callbacks: { label: ctx => ` ${fmt(ctx.raw)} atendidos` } }, datalabels: { anchor: 'end', align: 'end', color: '#2980b9', font: { family: 'Inter', size: 11, weight: 'bold' }, formatter: val => fmt(val) } },
      scales: { x: { ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 30 }, grid: { display: false } }, y: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

// ============================================================
// GRÁFICOS - CANCELADOS
// ============================================================

function renderChartCancelamentosDist() {
  const ctx = document.getElementById('chartCancelamentosDist')?.getContext('2d');
  if (!ctx) return;
  const map = {};
  filteredData.forEach(r => {
    const key = r.distrito || 'OUTROS';
    if (!map[key]) map[key] = { can: 0, total: 0 };
    map[key].total++;
    if (r.situacao === 'CAN') map[key].can++;
  });
  const entries = Object.entries(map).filter(([,v]) => v.total > 0).map(([k,v]) => ({ label: k, pct: parseFloat((v.can / v.total * 100).toFixed(1)), can: v.can, total: v.total })).sort((a,b) => b.pct - a.pct);
  const labels = entries.map(e => e.label);
  const data = entries.map(e => e.pct);
  destroyChart(chartCancelamentosDist);
  chartCancelamentosDist = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: '% Cancelamentos', data, backgroundColor: PALETTE_CANCELADOS.map(c => c + 'bb'), borderColor: PALETTE_CANCELADOS, borderWidth: 2, borderRadius: 7, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { ...TOOLTIP_BASE, callbacks: { label: (ctx) => { const e = entries[ctx.dataIndex]; return [` ${ctx.raw}% cancelados`, ` ${fmt(e.can)} cancelamentos de ${fmt(e.total)} registros`]; } } }, datalabels: { anchor: 'end', align: 'end', clamp: true, color: '#1a2a3a', font: { family: 'Inter', size: 11, weight: 'bold' }, formatter: val => val + '%' } },
      layout: { padding: { top: 28 } },
      scales: { x: { ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 35 }, grid: { display: false } }, y: { beginAtZero: true, max: 100, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6', callback: v => v + '%' }, grid: { display: false } } }
    }
  });
}

function renderChartCancelamentosEsp() {
  const ctx = document.getElementById('chartCancelamentosEsp')?.getContext('2d');
  if (!ctx) return;
  const map = {};
  filteredData.forEach(r => {
    const key = r.cbo || '–';
    if (!map[key]) map[key] = { can: 0, total: 0 };
    map[key].total++;
    if (r.situacao === 'CAN') map[key].can++;
  });
  const entries = Object.entries(map).filter(([,v]) => v.total > 0).map(([k,v]) => ({ label: k, pct: parseFloat((v.can / v.total * 100).toFixed(1)), can: v.can, total: v.total })).sort((a,b) => b.pct - a.pct).slice(0, 15);
  const labels = entries.map(e => e.label);
  const data = entries.map(e => e.pct);
  destroyChart(chartCancelamentosEsp);
  chartCancelamentosEsp = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: '% Cancelamentos', data, backgroundColor: PALETTE_CANCELADOS.map(c => c + 'bb'), borderColor: PALETTE_CANCELADOS, borderWidth: 2, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: { ...TOOLTIP_BASE, callbacks: { label: (ctx) => { const e = entries[ctx.dataIndex]; return [` ${ctx.raw}% cancelados`, ` ${fmt(e.can)} cancelamentos de ${fmt(e.total)} registros`]; } } }, datalabels: { anchor: 'end', align: 'end', clamp: true, color: '#3d5166', font: { family: 'Inter', size: 13, weight: '800' }, formatter: val => val + '%' } },
      layout: { padding: { right: 54 } },
      scales: { y: { ticks: { font: { family: 'Inter', size: 10 }, color: '#3d5166', callback(val) { const label = this.getLabelForValue(val); return label && label.length > 22 ? label.substring(0,20)+'…' : label; } }, grid: { display: false } }, x: { beginAtZero: true, max: 100, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6', callback: v => v + '%' }, grid: { display: false } } }
    }
  });
}

function renderChartCancelamentosPrestador() {
  const ctx = document.getElementById('chartCancelamentosPrestador')?.getContext('2d');
  if (!ctx) return;
  const map = {};
  filteredData.forEach(r => {
    const key = r.unidadeExecutante || '–';
    if (!map[key]) map[key] = { can: 0, total: 0 };
    map[key].total++;
    if (r.situacao === 'CAN') map[key].can++;
  });
  const entries = Object.entries(map).filter(([,v]) => v.total > 0).map(([k,v]) => ({ label: k, pct: parseFloat((v.can / v.total * 100).toFixed(1)), can: v.can, total: v.total })).sort((a,b) => b.pct - a.pct).slice(0, 10);
  const labels = entries.map(e => e.label);
  const data = entries.map(e => e.pct);
  destroyChart(chartCancelamentosPrestador);
  chartCancelamentosPrestador = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: '% Cancelamentos', data, backgroundColor: PALETTE_CANCELADOS.map(c => c + 'cc'), borderColor: PALETTE_CANCELADOS, borderWidth: 2, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: TOOLTIP_BASE, datalabels: { anchor: 'end', align: 'end', color: '#3d5166', font: { family: 'Inter', size: 10, weight: 'bold' }, formatter: val => val + '%' } },
      layout: { padding: { right: 44 } },
      scales: { y: { ticks: { font: { family: 'Inter', size: 10 }, color: '#3d5166' }, grid: { display: false } }, x: { beginAtZero: true, max: 100, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6', callback: v => v + '%' }, grid: { display: false } } }
    }
  });
}

function renderChartCancelamentosMensal() {
  const ctx = document.getElementById('chartCancelamentosMensal')?.getContext('2d');
  if (!ctx) return;
  const map = {};
  filteredData.forEach(r => {
    const key = r.mesAgendamento || '–';
    if (!map[key]) map[key] = { can: 0, total: 0 };
    map[key].total++;
    if (r.situacao === 'CAN') map[key].can++;
  });
  const mesesOrdenados = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const sortedKeys = Object.keys(map).sort((a, b) => {
    const [mesA, anoA] = a.split('/');
    const [mesB, anoB] = b.split('/');
    const indexMesA = mesesOrdenados.indexOf(mesA);
    const indexMesB = mesesOrdenados.indexOf(mesB);
    if (anoA !== anoB) return parseInt(anoA) - parseInt(anoB);
    return indexMesA - indexMesB;
  });
  const labels = sortedKeys;
  const data = sortedKeys.map(k => map[k].total > 0 ? map[k].can : 0);
  destroyChart(chartCancelamentosMensal);
  chartCancelamentosMensal = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets: [{ label: 'Cancelamentos', data, borderColor: '#c0392b', backgroundColor: 'rgba(192,57,43,0.1)', borderWidth: 3, pointRadius: 6, pointBackgroundColor: '#c0392b', tension: 0.4, fill: true }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { ...TOOLTIP_BASE, callbacks: { label: ctx => ` ${fmt(ctx.raw)} cancelamentos` } }, datalabels: { anchor: 'end', align: 'end', color: '#c0392b', font: { family: 'Inter', size: 11, weight: 'bold' }, formatter: val => fmt(val) } },
      scales: { x: { ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 30 }, grid: { display: false } }, y: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

// ============================================================
// GRÁFICOS - TRANSFERIDOS
// ============================================================

function renderChartTransferidosDistrito() {
  const ctx = document.getElementById('chartTransferidosDistrito')?.getContext('2d');
  if (!ctx) return;
  const traData = filteredData.filter(r => r.situacao === 'TRA');
  const counts = countBy(traData, r => r.distrito);
  const entries = sortedEntries(counts);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  destroyChart(chartTransferidosDistrito);
  chartTransferidosDistrito = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Transferidos', data, backgroundColor: PALETTE_TRANSFERIDOS.map(c => c + 'bb'), borderColor: PALETTE_TRANSFERIDOS, borderWidth: 2, borderRadius: 8, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: TOOLTIP_BASE, datalabels: { anchor: 'center', align: 'center', color: '#fff', font: { family: 'Inter', size: 13, weight: 'bold' }, formatter: val => val > 0 ? fmt(val) : '' } },
      scales: { x: { ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 30 }, grid: { display: false } }, y: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

function renderChartTransferidosEspecialidade() {
  const ctx = document.getElementById('chartTransferidosEspecialidade')?.getContext('2d');
  if (!ctx) return;
  const traData = filteredData.filter(r => r.situacao === 'TRA');
  const counts = countBy(traData, r => r.cbo);
  const entries = sortedEntries(counts, 15);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  destroyChart(chartTransferidosEspecialidade);
  chartTransferidosEspecialidade = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Transferidos', data, backgroundColor: PALETTE_TRANSFERIDOS.map(c => c + 'dd'), borderColor: PALETTE_TRANSFERIDOS, borderWidth: 2, borderRadius: 5, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: TOOLTIP_BASE, datalabels: { anchor: 'end', align: 'end', color: '#3d5166', font: { family: 'Inter', size: 13, weight: '800' }, formatter: val => fmt(val) } },
      layout: { padding: { right: 54 } },
      scales: { y: { ticks: { font: { family: 'Inter', size: 10 }, color: '#3d5166' }, grid: { display: false } }, x: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

function renderChartTransferidosPrestador() {
  const ctx = document.getElementById('chartTransferidosPrestador')?.getContext('2d');
  if (!ctx) return;
  const traData = filteredData.filter(r => r.situacao === 'TRA');
  const counts = countBy(traData, r => r.unidadeExecutante);
  const entries = sortedEntries(counts, 10);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  destroyChart(chartTransferidosPrestador);
  chartTransferidosPrestador = new Chart(ctx, {
    type: 'bar',
    data: { labels, datasets: [{ label: 'Transferidos', data, backgroundColor: PALETTE_TRANSFERIDOS.map(c => c + 'dd'), borderColor: PALETTE_TRANSFERIDOS, borderWidth: 2, borderRadius: 6, borderSkipped: false }] },
    options: {
      responsive: true, maintainAspectRatio: false, indexAxis: 'y',
      plugins: { legend: { display: false }, tooltip: TOOLTIP_BASE, datalabels: { anchor: 'end', align: 'end', color: '#3d5166', font: { family: 'Inter', size: 13, weight: '800' }, formatter: val => fmt(val) } },
      layout: { padding: { right: 54 } },
      scales: { y: { ticks: { font: { family: 'Inter', size: 9 }, color: '#3d5166' }, grid: { display: false } }, x: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

function renderChartTransferidosMensal() {
  const ctx = document.getElementById('chartTransferidosMensal')?.getContext('2d');
  if (!ctx) return;
  const map = {};
  filteredData.forEach(r => {
    const key = r.mesAgendamento || '–';
    if (!map[key]) map[key] = 0;
    if (r.situacao === 'TRA') map[key]++;
  });
  const mesesOrdenados = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho','Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];
  const sortedKeys = Object.keys(map).sort((a, b) => {
    const [mesA, anoA] = a.split('/');
    const [mesB, anoB] = b.split('/');
    const indexMesA = mesesOrdenados.indexOf(mesA);
    const indexMesB = mesesOrdenados.indexOf(mesB);
    if (anoA !== anoB) return parseInt(anoA) - parseInt(anoB);
    return indexMesA - indexMesB;
  });
  const labels = sortedKeys;
  const data = sortedKeys.map(k => map[k]);
  destroyChart(chartTransferidosMensal);
  chartTransferidosMensal = new Chart(ctx, {
    type: 'line',
    data: { labels, datasets: [{ label: 'Transferidos', data, borderColor: '#7d3c98', backgroundColor: 'rgba(125,60,152,0.1)', borderWidth: 3, pointRadius: 6, pointBackgroundColor: '#7d3c98', tension: 0.4, fill: true }] },
    options: {
      responsive: true, maintainAspectRatio: false,
      plugins: { legend: { display: false }, tooltip: { ...TOOLTIP_BASE, callbacks: { label: ctx => ` ${fmt(ctx.raw)} transferidos` } }, datalabels: { anchor: 'end', align: 'end', color: '#7d3c98', font: { family: 'Inter', size: 11, weight: 'bold' }, formatter: val => fmt(val) } },
      scales: { x: { ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 30 }, grid: { display: false } }, y: { beginAtZero: true, ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' }, grid: { display: false } } }
    }
  });
}

// ============================================================
// TABELA CONSOLIDADA (COM ATD)
// ============================================================
function buildTableData() {
  const map = {};
  filteredData.forEach(r => {
    const key = `${r.distrito}|||${r.unidadeSolicitante}|||${r.tipoAtendimento}|||${r.cbo}|||${r.profissional}|||${r.unidadeExecutante}`;
    if (!map[key]) {
      map[key] = {
        distrito: r.distrito, unidadeSolicitante: r.unidadeSolicitante,
        tipoServico: r.tipoAtendimento, cbo: r.cbo,
        profissional: r.profissional, unidadeExecutante: r.unidadeExecutante,
        age: 0, rec: 0, atd: 0, fal: 0, can: 0, tra: 0,
      };
    }
    const sit = r.situacao;
    if (sit === 'AGE') map[key].age++;
    else if (sit === 'REC') map[key].rec++;
    else if (sit === 'ATD') map[key].atd++;
    else if (sit === 'FAL') map[key].fal++;
    else if (sit === 'CAN') map[key].can++;
    else if (sit === 'TRA') map[key].tra++;
  });
  tableData = Object.values(map).map(r => {
    const totalAgendamentos = r.rec + r.fal + r.atd;
    const pctAbsenteismo = totalAgendamentos > 0 ? parseFloat((r.fal / totalAgendamentos * 100).toFixed(1)) : 0;
    return { ...r, totalAgendamentos, pctAbsenteismo };
  }).sort((a,b) => b.totalAgendamentos - a.totalAgendamentos);
  tableSearched = [...tableData];
}

function filterTable() {
  const q = (document.getElementById('tableSearch')?.value || '').toLowerCase();
  tableSearched = !q ? [...tableData] : tableData.filter(r =>
    (r.distrito||'').toLowerCase().includes(q) ||
    (r.unidadeSolicitante||'').toLowerCase().includes(q) ||
    (r.tipoServico||'').toLowerCase().includes(q) ||
    (r.cbo||'').toLowerCase().includes(q) ||
    (r.profissional||'').toLowerCase().includes(q) ||
    (r.unidadeExecutante||'').toLowerCase().includes(q)
  );
  currentPage = 1;
  renderTable();
}

function sortTable(col) {
  if (sortColIdx === col) sortAscFlag = !sortAscFlag;
  else { sortColIdx = col; sortAscFlag = true; }
  const keys = ['distrito','unidadeSolicitante','tipoServico','cbo','profissional','unidadeExecutante','age','rec','atd','fal','can','tra','totalAgendamentos','pctAbsenteismo'];
  const key = keys[col];
  tableSearched.sort((a,b) => {
    const va = a[key] ?? '';
    const vb = b[key] ?? '';
    const cmp = typeof va === 'number' ? va - vb : va.toString().localeCompare(vb.toString(), 'pt-BR');
    return sortAscFlag ? cmp : -cmp;
  });
  renderTable();
}

function absentClass(pct) {
  if (pct >= 30) return 'badge-absent badge-absent-high';
  if (pct >= 15) return 'badge-absent badge-absent-medium';
  return 'badge-absent badge-absent-low';
}

function renderTable() {
  const pageSize = parseInt(document.getElementById('tablePageSize')?.value || 15);
  const total = tableSearched.length;
  const pages = Math.max(1, Math.ceil(total / pageSize));
  if (currentPage > pages) currentPage = pages;
  const start = (currentPage - 1) * pageSize;
  const slice = tableSearched.slice(start, start + pageSize);
  const tbody = document.getElementById('tableBody');
  const tfoot = document.getElementById('tableFoot');
  if (!tbody) return;
  if (slice.length === 0) {
    tbody.innerHTML = '<tr><td colspan="14" class="empty-msg">Nenhum registro encontrado.</td></tr>';
    tfoot.innerHTML = '';
  } else {
    tbody.innerHTML = slice.map(r => `
      <tr>
        <td><span style="font-size:0.78rem;font-weight:700;color:#1e3a5f;background:rgba(30,58,95,0.08);border-radius:6px;padding:2px 8px;">${r.distrito || '–'}</span></td>
        <td style="font-size:0.78rem;color:#3d5166;">${r.unidadeSolicitante || '–'}</td>
        <td><span style="font-size:0.76rem;font-weight:600;color:${r.tipoServico === 'Primeira Consulta' ? '#1a7a3f' : '#7d3c98'}">${r.tipoServico || '–'}</span></td>
        <td><div style="font-weight:600;color:#1e3a5f;font-size:0.8rem;">${r.cbo || '–'}</div></td>
        <td style="font-size:0.78rem;color:#3d5166;">${r.profissional || '–'}</td>
        <td style="font-size:0.78rem;color:#3d5166;">${r.unidadeExecutante || '–'}</td>
        <td class="text-center"><span class="badge-num badge-age">${fmt(r.age)}</span></td>
        <td class="text-center"><span class="badge-num badge-rec">${fmt(r.rec)}</span></td>
        <td class="text-center"><span class="badge-num badge-atd">${fmt(r.atd)}</span></td>
        <td class="text-center"><span class="badge-num badge-fal">${fmt(r.fal)}</span></td>
        <td class="text-center"><span class="badge-num badge-can">${fmt(r.can)}</span></td>
        <td class="text-center"><span class="badge-num badge-tra">${fmt(r.tra)}</span></td>
        <td class="text-center"><span class="badge-num badge-total">${fmt(r.totalAgendamentos)}</span></td>
        <td class="text-center"><span class="${absentClass(r.pctAbsenteismo)}">${r.pctAbsenteismo.toFixed(1)}%</span></td>
      </tr>
    `).join('');
    const sAge = tableSearched.reduce((s,r) => s + r.age, 0);
    const sRec = tableSearched.reduce((s,r) => s + r.rec, 0);
    const sAtd = tableSearched.reduce((s,r) => s + r.atd, 0);
    const sFal = tableSearched.reduce((s,r) => s + r.fal, 0);
    const sCan = tableSearched.reduce((s,r) => s + r.can, 0);
    const sTra = tableSearched.reduce((s,r) => s + r.tra, 0);
    const sTotal = tableSearched.reduce((s,r) => s + r.totalAgendamentos, 0);
    const pctGeral = sTotal > 0 ? parseFloat((sFal / sTotal * 100).toFixed(1)) : 0;
    tfoot.innerHTML = `
      <tr>
        <td colspan="6"><i class="fas fa-calculator" style="margin-right:6px;"></i>TOTAL GERAL (${fmt(tableSearched.length)} linhas)</td>
        <td class="text-center">${fmt(sAge)}</td>
        <td class="text-center">${fmt(sRec)}</td>
        <td class="text-center">${fmt(sAtd)}</td>
        <td class="text-center">${fmt(sFal)}</td>
        <td class="text-center">${fmt(sCan)}</td>
        <td class="text-center">${fmt(sTra)}</td>
        <td class="text-center">${fmt(sTotal)}</td>
        <td class="text-center">${pctGeral.toFixed(1)}%</td>
      </tr>
    `;
  }
  document.getElementById('tablePaginationInfo').textContent = `Mostrando ${total === 0 ? 0 : start+1} a ${Math.min(start+pageSize, total)} de ${fmt(total)} registros`;
  renderPagination(currentPage, pages);
}

function renderPagination(cur, total) {
  const container = document.getElementById('pagination');
  if (!container) return;
  let html = `<button class="page-btn" onclick="goPage(${cur-1})" ${cur===1?'disabled':''}>‹</button>`;
  let pages = [];
  if (total <= 7) {
    for (let i=1; i<=total; i++) pages.push(i);
  } else {
    pages = [1];
    if (cur > 3) pages.push('...');
    for (let i=Math.max(2,cur-1); i<=Math.min(total-1,cur+1); i++) pages.push(i);
    if (cur < total-2) pages.push('...');
    pages.push(total);
  }
  pages.forEach(p => {
    if (p === '...') html += `<button class="page-btn" disabled>…</button>`;
    else html += `<button class="page-btn ${p===cur?'active':''}" onclick="goPage(${p})">${p}</button>`;
  });
  html += `<button class="page-btn" onclick="goPage(${cur+1})" ${cur===total?'disabled':''}>›</button>`;
  container.innerHTML = html;
}

function goPage(p) {
  const pageSize = parseInt(document.getElementById('tablePageSize')?.value || 15);
  const pages = Math.max(1, Math.ceil(tableSearched.length / pageSize));
  if (p < 1 || p > pages) return;
  currentPage = p;
  renderTable();
}

// ============================================================
// TABELA CONSOLIDADO PROFISSIONAIS POR ATENDIMENTO (SEM DISTRITO - COM ATD)
// ============================================================

function buildTableDataResumido() {
  console.log('🔄 Construindo tabela consolidado profissionais...');
  const map = {};

  if (!filteredData || filteredData.length === 0) {
    console.warn('⚠️ Nenhum dado filtrado para construir a tabela.');
    tableDataResumido = [];
    tableSearchedResumido = [];
    return;
  }

  filteredData.forEach(r => {
    const key = `${r.unidadeExecutante}|||${r.tipoAtendimento}|||${r.cbo}|||${r.profissional}`;
    if (!map[key]) {
      map[key] = {
        unidadeExecutante: r.unidadeExecutante || 'NÃO INFORMADO',
        tipoServico: r.tipoAtendimento || 'NÃO INFORMADO',
        cbo: r.cbo || 'NÃO INFORMADO',
        profissional: r.profissional || 'NÃO INFORMADO',
        age: 0, rec: 0, atd: 0, fal: 0, can: 0, tra: 0,
      };
    }
    const sit = r.situacao;
    if (sit === 'AGE') map[key].age++;
    else if (sit === 'REC') map[key].rec++;
    else if (sit === 'ATD') map[key].atd++;
    else if (sit === 'FAL') map[key].fal++;
    else if (sit === 'CAN') map[key].can++;
    else if (sit === 'TRA') map[key].tra++;
  });

  tableDataResumido = Object.values(map).map(r => {
    const totalAgendamentos = r.rec + r.fal + r.atd;
    const pctAbsenteismo = totalAgendamentos > 0 ? parseFloat((r.fal / totalAgendamentos * 100).toFixed(1)) : 0;
    return { ...r, totalAgendamentos, pctAbsenteismo };
  }).sort((a,b) => b.totalAgendamentos - a.totalAgendamentos);

  tableSearchedResumido = [...tableDataResumido];
  console.log(`✅ ${tableDataResumido.length} registros na tabela consolidado profissionais`);
}

function renderTableResumido() {
  console.log('🔄 Renderizando tabela consolidado profissionais...');

  const tbody = document.getElementById('tableBodyResumido');
  const tfoot = document.getElementById('tableFootResumido');

  if (!tbody) {
    console.error('❌ Elemento tableBodyResumido não encontrado!');
    return;
  }

  if (!tableSearchedResumido || tableSearchedResumido.length === 0) {
    tbody.innerHTML = '<tr><td colspan="12" class="empty-msg">Nenhum registro encontrado.</td></tr>';
    if (tfoot) tfoot.innerHTML = '';
    const infoEl = document.getElementById('tablePaginationInfoResumido');
    if (infoEl) infoEl.textContent = 'Mostrando 0 registros';
    return;
  }

  const pageSize = parseInt(document.getElementById('tablePageSizeResumido')?.value || 15);
  const total = tableSearchedResumido.length;
  const pages = Math.max(1, Math.ceil(total / pageSize));
  if (currentPageResumido > pages) currentPageResumido = pages;
  const start = (currentPageResumido - 1) * pageSize;
  const slice = tableSearchedResumido.slice(start, start + pageSize);

  tbody.innerHTML = slice.map(r => `
    <tr>
      <td style="font-size:0.78rem;color:#3d5166;">${r.unidadeExecutante || '–'}</td>
      <td><span style="font-size:0.76rem;font-weight:600;color:${r.tipoServico === 'Primeira Consulta' ? '#1a7a3f' : '#7d3c98'}">${r.tipoServico || '–'}</span></td>
      <td><div style="font-weight:600;color:#1e3a5f;font-size:0.8rem;">${r.cbo || '–'}</div></td>
      <td style="font-size:0.78rem;color:#3d5166;">${r.profissional || '–'}</td>
      <td class="text-center"><span class="badge-num badge-age">${fmt(r.age)}</span></td>
      <td class="text-center"><span class="badge-num badge-rec">${fmt(r.rec)}</span></td>
      <td class="text-center"><span class="badge-num badge-atd">${fmt(r.atd)}</span></td>
      <td class="text-center"><span class="badge-num badge-fal">${fmt(r.fal)}</span></td>
      <td class="text-center"><span class="badge-num badge-can">${fmt(r.can)}</span></td>
      <td class="text-center"><span class="badge-num badge-tra">${fmt(r.tra)}</span></td>
      <td class="text-center"><span class="badge-num badge-total">${fmt(r.totalAgendamentos)}</span></td>
      <td class="text-center"><span class="${absentClass(r.pctAbsenteismo)}">${r.pctAbsenteismo.toFixed(1)}%</span></td>
    </tr>
  `).join('');

  if (tfoot) {
    const sAge = tableSearchedResumido.reduce((s,r) => s + r.age, 0);
    const sRec = tableSearchedResumido.reduce((s,r) => s + r.rec, 0);
    const sAtd = tableSearchedResumido.reduce((s,r) => s + r.atd, 0);
    const sFal = tableSearchedResumido.reduce((s,r) => s + r.fal, 0);
    const sCan = tableSearchedResumido.reduce((s,r) => s + r.can, 0);
    const sTra = tableSearchedResumido.reduce((s,r) => s + r.tra, 0);
    const sTotal = tableSearchedResumido.reduce((s,r) => s + r.totalAgendamentos, 0);
    const pctGeral = sTotal > 0 ? parseFloat((sFal / sTotal * 100).toFixed(1)) : 0;

    tfoot.innerHTML = `
      <tr>
        <td colspan="4"><i class="fas fa-calculator" style="margin-right:6px;"></i>TOTAL GERAL (${fmt(tableSearchedResumido.length)} linhas)</td>
        <td class="text-center">${fmt(sAge)}</td>
        <td class="text-center">${fmt(sRec)}</td>
        <td class="text-center">${fmt(sAtd)}</td>
        <td class="text-center">${fmt(sFal)}</td>
        <td class="text-center">${fmt(sCan)}</td>
        <td class="text-center">${fmt(sTra)}</td>
        <td class="text-center">${fmt(sTotal)}</td>
        <td class="text-center">${pctGeral.toFixed(1)}%</td>
      </tr>
    `;
  }

  const infoEl = document.getElementById('tablePaginationInfoResumido');
  if (infoEl) {
    infoEl.textContent = `Mostrando ${total === 0 ? 0 : start+1} a ${Math.min(start+pageSize, total)} de ${fmt(total)} registros`;
  }
  renderPaginationResumido(currentPageResumido, pages);
}

function filterTableResumido() {
  const q = (document.getElementById('tableSearchResumido')?.value || '').toLowerCase();
  tableSearchedResumido = !q ? [...tableDataResumido] : tableDataResumido.filter(r =>
    (r.unidadeExecutante||'').toLowerCase().includes(q) ||
    (r.tipoServico||'').toLowerCase().includes(q) ||
    (r.cbo||'').toLowerCase().includes(q) ||
    (r.profissional||'').toLowerCase().includes(q)
  );
  currentPageResumido = 1;
  renderTableResumido();
}

function sortTableResumido(col) {
  if (sortColIdxResumido === col) sortAscFlagResumido = !sortAscFlagResumido;
  else { sortColIdxResumido = col; sortAscFlagResumido = true; }
  const keys = ['unidadeExecutante','tipoServico','cbo','profissional','age','rec','atd','fal','can','tra','totalAgendamentos','pctAbsenteismo'];
  const key = keys[col - 1];
  if (!key) return;
  tableSearchedResumido.sort((a,b) => {
    const va = a[key] ?? '';
    const vb = b[key] ?? '';
    const cmp = typeof va === 'number' ? va - vb : va.toString().localeCompare(vb.toString(), 'pt-BR');
    return sortAscFlagResumido ? cmp : -cmp;
  });
  renderTableResumido();
}

function renderPaginationResumido(cur, total) {
  const container = document.getElementById('paginationResumido');
  if (!container) return;
  let html = `<button class="page-btn" onclick="goPageResumido(${cur-1})" ${cur===1?'disabled':''}>‹</button>`;
  let pages = [];
  if (total <= 7) {
    for (let i=1; i<=total; i++) pages.push(i);
  } else {
    pages = [1];
    if (cur > 3) pages.push('...');
    for (let i=Math.max(2,cur-1); i<=Math.min(total-1,cur+1); i++) pages.push(i);
    if (cur < total-2) pages.push('...');
    pages.push(total);
  }
  pages.forEach(p => {
    if (p === '...') html += `<button class="page-btn" disabled>…</button>`;
    else html += `<button class="page-btn ${p===cur?'active':''}" onclick="goPageResumido(${p})">${p}</button>`;
  });
  html += `<button class="page-btn" onclick="goPageResumido(${cur+1})" ${cur===total?'disabled':''}>›</button>`;
  container.innerHTML = html;
}

function goPageResumido(p) {
  const pageSize = parseInt(document.getElementById('tablePageSizeResumido')?.value || 15);
  const pages = Math.max(1, Math.ceil(tableSearchedResumido.length / pageSize));
  if (p < 1 || p > pages) return;
  currentPageResumido = p;
  renderTableResumido();
}

// ============================================================
// EXPORTAR EXCEL
// ============================================================
function exportExcel() {
  if (!filteredData.length) { alert('Nenhum dado para exportar.'); return; }
  const btn = document.getElementById('btnExcel');
  btn.disabled = true;
  btn.innerHTML = '<i class="fas fa-spinner fa-spin"></i> Gerando...';
  setTimeout(() => {
    try {
      const wsData = filteredData.map(r => ({
        'Unidade Executante': r.unidadeExecutante,
        'Unidade Solicitante': r.unidadeSolicitante,
        'Distrito': r.distrito,
        'Especialidade (CBO)': r.cbo,
        'Tipo Especialidade': r.especialidade,
        'Profissional': r.profissional,
        'Tipo Atendimento': r.tipoAtendimento,
        'Situação': r.situacaoLabel,
        'Operador': r.operador,
        'Data Agenda': r.dataAgenda,
        'Data Criação': r.dataCriacao,
        'Mês Agendamento': r.mesAgendamento,
      }));
      const wsSummary = tableData.map(r => ({
        'Distrito': r.distrito,
        'Unidade Solicitante': r.unidadeSolicitante,
        'Tipo de Serviço': r.tipoServico,
        'CBO / Especialidade': r.cbo,
        'Profissional Médico': r.profissional,
        'Unidade Executante': r.unidadeExecutante,
        'AGE': r.age,
        'REC (Recepcionados)': r.rec,
        'ATD (Atendidos)': r.atd,
        'FAL (Faltosos)': r.fal,
        'CAN (Cancelados)': r.can,
        'TRA (Transferidos)': r.tra,
        'Total Agendamentos (REC+FAL+ATD)': r.totalAgendamentos,
        '% Absenteísmo': r.pctAbsenteismo + '%',
      }));
      const wb = XLSX.utils.book_new();
      const ws1 = XLSX.utils.json_to_sheet(wsData);
      autoSizeColumns(ws1, wsData);
      XLSX.utils.book_append_sheet(wb, ws1, 'Dados Filtrados');
      const ws2 = XLSX.utils.json_to_sheet(wsSummary);
      autoSizeColumns(ws2, wsSummary);
      XLSX.utils.book_append_sheet(wb, ws2, 'Consolidado Agendamentos');
      const now = new Date();
      const fname = `CMC_Consolidado_${now.getFullYear()}${String(now.getMonth()+1).padStart(2,'0')}${String(now.getDate()).padStart(2,'0')}.xlsx`;
      XLSX.writeFile(wb, fname);
    } catch(e) {
      console.error(e);
      alert('Erro ao gerar o arquivo Excel.');
    }
    btn.disabled = false;
    btn.innerHTML = '<i class="fas fa-file-excel"></i> Excel';
  }, 100);
}

function autoSizeColumns(ws, data) {
  if (!data.length) return;
  const cols = Object.keys(data[0]);
  ws['!cols'] = cols.map(col => ({
    wch: Math.min(data.reduce((max, row) => Math.max(max, (row[col]||'').toString().length), col.length) + 2, 60)
  }));
}

// ============================================================
// UTILITÁRIOS UI
// ============================================================
function showLoading(show) {
  document.getElementById('loadingOverlay')?.classList.toggle('hidden', !show);
}

function setStatus(msg, ok) {
  const el  = document.getElementById('statusText');
  const dot = document.querySelector('.status-dot');
  if (el)  el.textContent = msg;
  if (dot) dot.className  = 'status-dot ' + (ok ? 'connected' : 'error');
}

function showError(msg) {
  setStatus('Erro', false);
  showLoading(false);
  const toast = document.createElement('div');
  toast.style.cssText = 'position:fixed;bottom:24px;right:24px;z-index:9998;background:#c0392b;color:#fff;border-radius:12px;padding:14px 22px;font-family:Inter,sans-serif;font-size:.85rem;font-weight:600;box-shadow:0 6px 24px rgba(0,0,0,.3);display:flex;align-items:center;gap:10px;';
  toast.innerHTML = `<i class="fas fa-exclamation-circle"></i> ${msg}`;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 7000);
}

function updateLastUpdate(date) {
  const el = document.getElementById('lastUpdate');
  if (el) {
    const d = date || new Date();
    el.textContent = `Última atualização: ${d.toLocaleString('pt-BR')}`;
  }
}

// ============================================================
// DATE PICKER
// ============================================================
function initDatePickers() {
  window._fpInicio = flatpickr('#filterDataInicio', {
    locale: 'pt',
    dateFormat: 'd/m/Y',
    allowInput: false,
    disableMobile: false,
    onChange: () => applyFilters()
  });
}

// ============================================================
// FORÇAR RENDERIZAÇÃO DA TABELA CONSOLIDADO PROFISSIONAIS
// ============================================================
function fixResumidoTab() {
  console.log('🔧 Aplicando correção da aba resumida...');

  const resumidoPanel = document.getElementById('tab-profissionais-atendimento');
  if (resumidoPanel && resumidoPanel.classList.contains('active')) {
    console.log('✅ Aba profissionais-atendimento está ativa, renderizando...');
    if (tableDataResumido && tableDataResumido.length > 0) {
      tableSearchedResumido = [...tableDataResumido];
      currentPageResumido = 1;
      renderTableResumido();
    } else {
      console.log('🔄 Construindo dados da tabela resumida...');
      buildTableDataResumido();
      renderTableResumido();
    }
  } else {
    console.log('ℹ️ Aba profissionais-atendimento não está ativa no momento.');
  }
}

const originalOnDataLoaded = onDataLoaded;
onDataLoaded = function() {
  originalOnDataLoaded();
  setTimeout(fixResumidoTab, 500);
};

document.addEventListener('DOMContentLoaded', function() {
  const resumidoBtn = document.querySelector('[data-tab="profissionais-atendimento"]');
  if (resumidoBtn) {
    resumidoBtn.addEventListener('click', function() {
      setTimeout(function() {
        console.log('🔄 Aba profissionais-atendimento clicada, renderizando...');
        if (tableDataResumido && tableDataResumido.length > 0) {
          tableSearchedResumido = [...tableDataResumido];
          currentPageResumido = 1;
          renderTableResumido();
        } else {
          buildTableDataResumido();
          renderTableResumido();
        }
      }, 200);
    });
  }

  setTimeout(fixResumidoTab, 1500);
});

// ============================================================
// INICIALIZAÇÃO
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  initDatePickers();
  initTabs();
  loadData();
});
