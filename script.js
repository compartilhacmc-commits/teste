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

// ============================================================
// MAPEAMENTO DE DISTRITOS
// ============================================================
const DISTRITO_MAP = {
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
  'UNIDADE BASICA DE SAUDE INDUSTRIAL III SECAO': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE JARDIM INDUSTRIAL': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE VILA SAO PAULO': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE SANDOVAL DE AZEVEDO': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE JOAO EVANGELISTA': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE BANDEIRANTES': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE AMAZONAS': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE VILA DINIZ': 'INDUSTRIAL',
  'UNIDADE BASICA DE SAUDE NACIONAL': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE ILDA EFIGENIA': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE JOAQUIM MURTINHO': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE XANGRILA': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE AMENDOEIRAS': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE ESTRELA DALVA': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE TIJUCA': 'NACIONAL',
  'UNIDADE BASICA DE SAUDE PETROLANDIA': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE TROPICAL II': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE SAPUCAIAS': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE DUQUE DE CAXIAS': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE SAO LUIZ II': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE SAO LUIZ I': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE CAMPO ALTO': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE TROPICAL I': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE ESTANCIAS IMPERIAIS': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE NASCENTES IMPERIAIS': 'PETROLANDIA',
  'UNIDADE BASICA DE SAUDE CAMPINA VERDE': 'RESSACA',
  'UNIDADE BASICA DE SAUDE LAGUNA': 'RESSACA',
  'UNIDADE BASICA DE SAUDE ARPOADOR': 'RESSACA',
  'UNIDADE BASICA DE SAUDE SAO JOAQUIM': 'RESSACA',
  'UNIDADE BASICA DE SAUDE PARQUE TURISTA': 'RESSACA',
  'UNIDADE BASICA DE SAUDE VILA PEROLA': 'RESSACA',
  'UNIDADE BASICA DE SAUDE NOVO PROGRESSO II': 'RESSACA',
  'UNIDADE BASICA DE SAUDE COLORADO': 'RESSACA',
  'UNIDADE BASICA DE SAUDE CANDIDA FERREIRA': 'RESSACA',
  'UNIDADE DE SAUDE DA FAMILIA VILA PEROLA II USF 84': 'RESSACA',
  'UNIDADE BASICA DE SAUDE PRESIDENTE KENNEDY': 'RESSACA',
  'UNIDADE BASICA DE SAUDE OITIS': 'RESSACA',
  'UNIDADE BASICA DE SAUDE MORADA NOVA': 'RESSACA',
  'UNIDADE BASICA DE SAUDE INCONFIDENTES': 'RIACHO',
  'UNIDADE BASICA DE SAUDE RIACHO': 'RIACHO',
  'UNIDADE BASICA DE SAUDE FLAMENGO': 'RIACHO',
  'UNIDADE BASICA DE SAUDE NOVO RIACHO': 'RIACHO',
  'UNIDADE BASICA DE SAUDE DURVAL DE BARROS': 'RIACHO',
  'UNIDADE BASICA DE SAUDE MONTE CASTELO': 'RIACHO',
  'UNIDADE BASICA DE SAUDE CHACARAS': 'SEDE',
  'UNIDADE BASICA DE SAUDE CANADA': 'SEDE',
  'UNIDADE BASICA DE SAUDE CENTRO (CAD)': 'SEDE',
  'UBS BERNARDO MONTEIRO/MOACIR PINTO MOREIRA': 'SEDE',
  'UNIDADE BASICA DE SAUDE LINDA VISTA': 'SEDE',
  'UNIDADE BASICA DE SAUDE SANTA HELENA': 'SEDE',
  'UNIDADE BASICA DE SAUDE VILA ITALIA': 'SEDE',
  'UNIDADE BASICA DE SAUDE BEATRIZ USF 72': 'SEDE',
  'UNIDADE BASICA DE SAUDE MARIA DA CONCEICAO': 'SEDE',
  'UNIDADE BASICA DE SAUDE FUNCIONARIOS': 'SEDE',
  'UNIDADE BASICA DE SAUDE PRAIA': 'SEDE',
  'UBS UNIDADE XVI (SEDE)': 'SEDE',
  'UNIDADE BASICA DE SAUDE VILA RENASCER': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE NOVA CONTAGEM': 'VARGEM DAS FLORES',
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
  'UNIDADE BASICA DE SAUDE NOVA CONTAGEM I': 'VARGEM DAS FLORES',
  'CONTAGEM PENITENCIARIA NELSON HUNGRIA': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE TUPA': 'VARGEM DAS FLORES',
  'UNIDADE BASICA DE SAUDE LIBERDADE II': 'VARGEM DAS FLORES'
};

const CAE_UNITS = ['CAE IRIA DINIZ', 'CAE RESSACA', 'CEAPS'];
const MESES_PT  = ['Janeiro','Fevereiro','Março','Abril','Maio','Junho',
                   'Julho','Agosto','Setembro','Outubro','Novembro','Dezembro'];

// ============================================================
// PALETAS DE CORES
// ============================================================
const PALETTE_BLUE    = ['#1e3a5f','#2d5494','#4a90d9','#74b3e8','#a8d1f5','#c8e4fb'];
const PALETTE_PURPLE  = ['#6c3483','#7d3c98','#9b59b6','#af7ac5','#c39bd3','#d7bde2'];
const PALETTE_TEAL    = ['#0e6655','#148069','#1abc9c','#45d1a0','#76ddb5','#a8e8ce','#c0f0df','#d4f7ec','#e0fbf3','#e8fdf7','#f0fefb','#f5fffd'];
const PALETTE_ORANGE  = ['#d35400','#e67e22','#f39c12','#f5b041','#f8c471','#fad7a0'];
const PALETTE_ROSE    = ['#922b21','#c0392b','#e74c3c','#ec7063','#f1948a','#f5b7b1'];
const PALETTE_GRAPE   = ['#512e5f','#7d3c98','#9b59b6','#bb8fce','#d2b4de','#e8daef'];

// ============================================================
// PALETA MULTICOLOR – gráfico Rosca por Mês
// ============================================================
const PALETTE_MONTHS = [
  '#e74c3c', '#3498db', '#e67e22', '#2ecc71',
  '#9b59b6', '#1abc9c', '#f39c12', '#e91e63',
  '#00bcd4', '#ff5722', '#8bc34a', '#673ab7'
];

// ============================================================
// PALETA VERDE – gráfico de Situação (Barras Verdes)
// ============================================================
const PALETTE_GREEN = [
  '#1a7a3f', '#27ae60', '#2ecc71', '#52be80', '#76d7c4'
];

const SITUACAO_COLORS = {
  AGE: { bg: 'rgba(41,128,185,0.85)',  border: '#2980b9', label: 'Agendados (AGE)' },
  REC: { bg: 'rgba(39,174,96,0.85)',   border: '#27ae60', label: 'Recepcionados' },
  FAL: { bg: 'rgba(230,126,34,0.85)',  border: '#e67e22', label: 'Faltosos' },
  CAN: { bg: 'rgba(192,57,43,0.85)',   border: '#c0392b', label: 'Cancelados' },
  TRA: { bg: 'rgba(125,60,152,0.85)',  border: '#7d3c98', label: 'Transferidos' }
};

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

// Instâncias dos gráficos
let chartDistrito            = null;
let chartTipoAtendimento     = null;
let chartEspecialidade       = null;
let chartPrestador           = null;
let chartSituacao            = null;
let chartMeses               = null;

let chartAbsenteismoEsp      = null;
let chartAbsenteismoDist     = null;
let chartCancelamentosDist   = null;
let chartCancelamentosEsp    = null;

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
  const map = { AGE:'Agendados', CAN:'Cancelados', FAL:'Faltosos', REC:'Recepcionados', TRA:'Transferidos' };
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

// ============================================================
// CARREGAR DADOS
// ============================================================
async function loadData() {
  showLoading(true);
  setStatus('Carregando...', false);
  const icon = document.getElementById('refreshIcon');
  icon.classList.add('spinning');

  try {
    const response = await fetch(CSV_URL + '&t=' + Date.now(), { cache: 'no-store', mode: 'cors' });
    if (!response.ok) throw new Error(`HTTP ${response.status}`);
    const text = await response.text();

    Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      transformHeader: h => h.trim(),
      complete(results) {
        allData = normalizeData(results.data);
        populateFilterOptions();
        applyFilters();
        setStatus('Conectado', true);
        updateLastUpdate();
        showLoading(false);
        icon.classList.remove('spinning');
      },
      error(err) {
        console.error('Parse error:', err);
        showError('Erro ao processar os dados.');
        icon.classList.remove('spinning');
        showLoading(false);
      }
    });
  } catch (err) {
    console.error('Fetch error:', err);
    showError('Não foi possível carregar os dados. Verifique as permissões da planilha.');
    icon.classList.remove('spinning');
    showLoading(false);
  }
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

    const dataCriacao       = get('DATA CRIAÇÃO DO AGENDAMENTO','DATA CRIACAO DO AGENDAMENTO','DATA CRIAÇÃO','DATA_CRIACAO');
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
// POPULAR FILTROS
// ============================================================
function populateFilterOptions() {
  populateSelect('filterPrestador',    [...new Set(allData.map(r => r.unidadeExecutante).filter(Boolean))].sort());
  populateSelect('filterEspecialidade',[...new Set(allData.map(r => r.cbo).filter(Boolean))].sort());
  populateSelect('filterProfissional', [...new Set(allData.map(r => r.profissional).filter(Boolean))].sort());
  populateSelect('filterUnidade',      [...new Set(allData.map(r => r.unidadeSolicitante).filter(Boolean))].sort());
  populateSelect('filterDistrito',     [...new Set(allData.map(r => r.distrito).filter(Boolean))].sort());

  const mesesSet = {};
  allData.forEach(r => {
    if (r.dataAgendaParsed && r.mesAgendamento) {
      const d   = r.dataAgendaParsed;
      const key = d.getFullYear() * 100 + d.getMonth();
      mesesSet[key] = r.mesAgendamento;
    }
  });
  const mesesOrdenados = Object.entries(mesesSet).sort((a,b) => +a[0] - +b[0]).map(e => e[1]);
  populateSelect('filterMes', mesesOrdenados);
}

function populateSelect(id, values) {
  const sel = document.getElementById(id);
  if (!sel) return;
  const current = sel.value;
  while (sel.options.length > 1) sel.remove(1);
  values.forEach(v => {
    if (!v) return;
    const opt = document.createElement('option');
    opt.value = v; opt.textContent = v;
    sel.appendChild(opt);
  });
  if (current) sel.value = current;
}

// ============================================================
// APLICAR FILTROS (suporte a múltiplos valores em situação)
// ============================================================
function applyFilters() {
  const prestador    = document.getElementById('filterPrestador')?.value    || '';
  const especialidade= document.getElementById('filterEspecialidade')?.value|| '';
  const tipoServico  = document.getElementById('filterTipoServico')?.value  || '';
  const profissional = document.getElementById('filterProfissional')?.value || '';
  const mes          = document.getElementById('filterMes')?.value          || '';
  const unidade      = document.getElementById('filterUnidade')?.value      || '';
  const distrito     = document.getElementById('filterDistrito')?.value     || '';
  const situacao     = document.getElementById('filterSituacao')?.value     || '';
  const dataCriacaoSelecionada = window._fpInicio ? window._fpInicio.selectedDates[0] : null;

  // Converter situação em array se tiver múltiplos valores separados por vírgula
  const situacoesPermitidas = situacao ? situacao.split(',') : [];

  filteredData = allData.filter(r => {
    if (prestador    && r.unidadeExecutante   !== prestador)    return false;
    if (especialidade && r.cbo               !== especialidade) return false;
    if (tipoServico  && r.tipoAtendimento    !== tipoServico)   return false;
    if (profissional && r.profissional       !== profissional)  return false;
    if (mes          && r.mesAgendamento     !== mes)           return false;
    if (unidade      && r.unidadeSolicitante !== unidade)       return false;
    if (distrito     && r.distrito           !== distrito)      return false;
    
    // Filtro de situação com suporte a múltiplos valores
    if (situacoesPermitidas.length > 0 && !situacoesPermitidas.includes(r.situacao)) return false;
    
    if (dataCriacaoSelecionada) {
      if (!r.dataCriacaoParsed) return false;
      if (!isSameDay(r.dataCriacaoParsed, dataCriacaoSelecionada)) return false;
    }
    return true;
  });

  updateKPIs();
  renderAllCharts();
  buildTableData();
  currentPage = 1;
  renderTable();
}

function clearFilters() {
  ['filterPrestador','filterEspecialidade','filterTipoServico','filterProfissional',
   'filterMes','filterUnidade','filterDistrito','filterSituacao'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  if (window._fpInicio) window._fpInicio.clear();
  applyFilters();
}

// ============================================================
// KPIS (cards respeitam os filtros atuais)
// ============================================================
function updateKPIs() {
  const rec = filteredData.filter(r => r.situacao === 'REC').length;
  const fal = filteredData.filter(r => r.situacao === 'FAL').length;
  const can = filteredData.filter(r => r.situacao === 'CAN').length;
  const tra = filteredData.filter(r => r.situacao === 'TRA').length;
  
  // Agendados = REC + FAL (dos dados filtrados atualmente)
  const agendados = rec + fal;

  animateCount('kpiAgendados', agendados);
  animateCount('kpiRecepcionados', rec);
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
// RENDER TODOS OS GRÁFICOS
// ============================================================
function renderAllCharts() {
  renderChartDistrito();
  renderChartTipoAtendimento();
  renderChartEspecialidade();
  renderChartPrestador();
  renderChartSituacao();
  renderChartMeses();

  renderChartAbsenteismoEsp();
  renderChartAbsenteismoDist();
  renderChartCancelamentosDist();
  renderChartCancelamentosEsp();
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
// GRÁFICO 1: Agendamentos por Distrito
// ============================================================
function renderChartDistrito() {
  const ctx = document.getElementById('chartDistrito')?.getContext('2d');
  if (!ctx) return;

  const counts  = countBy(filteredData, r => r.distrito);
  const entries = sortedEntries(counts);
  const labels  = entries.map(e => e[0]);
  const data    = entries.map(e => e[1]);
  const total   = data.reduce((a,b) => a+b, 0);

  destroyChart(chartDistrito);
  chartDistrito = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Agendamentos',
        data,
        backgroundColor: labels.map(() => 'rgba(192,57,43,0.85)'),
        borderColor:     labels.map(() => '#c0392b'),
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
          anchor: 'center', align: 'center',
          color: '#fff',
          textStrokeColor: 'rgba(0,0,0,0.30)', textStrokeWidth: 2,
          font: { family: 'Inter', size: 13, weight: 'bold' },
          formatter: val => val > 0 ? fmt(val) : ''
        }
      },
      scales: {
        x: {
          ticks: {
            font: { family: 'Inter', size: 10, weight: '600' },
            color: '#3d5166', maxRotation: 30
          },
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

// ============================================================
// GRÁFICO 2: 1ª Consulta vs Retorno
// ============================================================
function renderChartTipoAtendimento() {
  const ctx = document.getElementById('chartTipoAtendimento')?.getContext('2d');
  if (!ctx) return;

  const pc  = filteredData.filter(r => r.tipoAtendimento === 'Primeira Consulta').length;
  const ret = filteredData.filter(r => r.tipoAtendimento === 'Retorno').length;

  destroyChart(chartTipoAtendimento);
  chartTipoAtendimento = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['1ª Consulta', 'Retorno'],
      datasets: [{
        label: 'Quantidade',
        data: [pc, ret],
        backgroundColor: ['rgba(30,58,95,0.88)','rgba(39,174,96,0.88)'],
        borderColor:     ['#1e3a5f','#27ae60'],
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
          anchor: 'center', align: 'center',
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

// ============================================================
// GRÁFICO 3: Agendamentos por Especialidade (Barras Horizontais)
// ============================================================
function renderChartEspecialidade() {
  const ctx = document.getElementById('chartEspecialidade')?.getContext('2d');
  if (!ctx) return;

  const counts  = countBy(filteredData, r => r.cbo);
  const entries = sortedEntries(counts, 15);
  const labels  = entries.map(e => e[0]);
  const data    = entries.map(e => e[1]);

  destroyChart(chartEspecialidade);
  chartEspecialidade = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Agendamentos',
        data,
        backgroundColor: labels.map((_,i) => PALETTE_PURPLE[i % PALETTE_PURPLE.length] + 'dd'),
        borderColor:     labels.map((_,i) => PALETTE_PURPLE[i % PALETTE_PURPLE.length]),
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
          anchor: 'end', align: 'end',
          color: '#3d5166',
          font: { family: 'Inter', size: 13, weight: '800' },
          formatter: val => fmt(val)
        }
      },
      layout: { padding: { right: 54 } },
      scales: {
        y: {
          ticks: {
            font: { family: 'Inter', size: 10 }, color: '#3d5166',
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

// ============================================================
// GRÁFICO 4: Agendamentos por Prestador (Barras Horizontais)
// ============================================================
function renderChartPrestador() {
  const ctx = document.getElementById('chartPrestador')?.getContext('2d');
  if (!ctx) return;

  const counts  = countBy(filteredData, r => r.unidadeExecutante);
  const entries = sortedEntries(counts, 10);
  const labels  = entries.map(e => e[0]);
  const data    = entries.map(e => e[1]);

  destroyChart(chartPrestador);
  chartPrestador = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Agendamentos',
        data,
        backgroundColor: labels.map((_,i) => PALETTE_BLUE[i % PALETTE_BLUE.length] + 'dd'),
        borderColor:     labels.map((_,i) => PALETTE_BLUE[i % PALETTE_BLUE.length]),
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
          anchor: 'end', align: 'end',
          color: '#3d5166',
          font: { family: 'Inter', size: 13, weight: '800' },
          formatter: val => fmt(val)
        }
      },
      layout: { padding: { right: 54 } },
      scales: {
        y: {
          ticks: {
            font: { family: 'Inter', size: 9 }, color: '#3d5166',
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

// ============================================================
// GRÁFICO 5: Quantidade de Consultas por Situação (BARRAS VERDES)
// ============================================================
function renderChartSituacao() {
  const ctx = document.getElementById('chartSituacao')?.getContext('2d');
  if (!ctx) return;

  const age = filteredData.filter(r => r.situacao === 'AGE').length;
  const rec = filteredData.filter(r => r.situacao === 'REC').length;
  const fal = filteredData.filter(r => r.situacao === 'FAL').length;
  const can = filteredData.filter(r => r.situacao === 'CAN').length;
  const tra = filteredData.filter(r => r.situacao === 'TRA').length;

  const labels = ['Agendados (AGE)', 'Recepcionados (REC)', 'Faltosos (FAL)', 'Cancelados (CAN)', 'Transferidos (TRA)'];
  const data = [age, rec, fal, can, tra];
  const total = data.reduce((a,b) => a+b, 0);

  destroyChart(chartSituacao);
  chartSituacao = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Quantidade',
        data,
        backgroundColor: [
          'rgba(13,71,36,0.88)',    // Verde escuro muito escuro para AGE
          'rgba(26,122,63,0.88)',   // Verde escuro para REC
          'rgba(39,174,96,0.88)',   // Verde escuro médio para FAL
          'rgba(52,152,85,0.88)',   // Verde escuro claro para CAN
          'rgba(65,130,74,0.88)'    // Verde escuro teal para TRA
        ],
        borderColor: [
          '#0d4724',
          '#1a7a3f',
          '#27ae60',
          '#349855',
          '#41824a'
        ],
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
          anchor: 'center', align: 'center',
          color: '#fff',
          textStrokeColor: 'rgba(0,0,0,0.30)', textStrokeWidth: 2,
          font: { family: 'Inter', size: 13, weight: 'bold' },
          formatter: val => val > 0 ? fmt(val) : ''
        }
      },
      scales: {
        x: {
          ticks: {
            font: { family: 'Inter', size: 10, weight: '600' },
            color: '#3d5166', maxRotation: 35
          },
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

// ============================================================
// GRÁFICO 6: Distribuição por Mês – ROSCA (MULTICOLOR)
// ============================================================
function renderChartMeses() {
  const ctx = document.getElementById('chartMeses')?.getContext('2d');
  if (!ctx) return;

  const counts = countBy(filteredData, r => r.mesAgendamento);
  const entries = sortedEntries(counts);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);
  const total = data.reduce((a,b) => a+b, 0);

  destroyChart(chartMeses);
  chartMeses = new Chart(ctx, {
    type: 'doughnut',
    data: {
      labels,
      datasets: [{
        data,
        backgroundColor: labels.map((_,i) => PALETTE_MONTHS[i % PALETTE_MONTHS.length]),
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
            generateLabels: (chart) => {
              const data = chart.data;
              return data.labels.map((label, i) => ({
                text: `${label}: ${fmt(data.datasets[0].data[i])}`,
                fillStyle: data.datasets[0].backgroundColor[i],
                hidden: false,
                index: i
              }));
            }
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
          formatter: (val, ctx) => {
            const pct = total > 0 ? (val/total*100).toFixed(0) : 0;
            return pct + '%';
          }
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
// GRÁFICO 7: % Absenteísmo por Especialidade
// ============================================================
function renderChartAbsenteismoEsp() {
  const ctx = document.getElementById('chartAbsenteismoEsp')?.getContext('2d');
  if (!ctx) return;

  const map = {};
  filteredData.forEach(r => {
    const key = r.cbo || '–';
    if (!map[key]) map[key] = { fal: 0, rec: 0, can: 0 };
    
    if (r.situacao === 'FAL') map[key].fal++;
    else if (r.situacao === 'REC') map[key].rec++;
    else if (r.situacao === 'CAN') map[key].can++;
  });

  const entries = Object.entries(map)
    .filter(([, v]) => (v.rec + v.fal + v.can) > 0)
    .map(([k, v]) => {
      const totalAgendamentos = v.rec + v.fal + v.can;
      return {
        label: k,
        pct:   parseFloat((v.fal / totalAgendamentos * 100).toFixed(1)),
        fal:   v.fal,
        total: totalAgendamentos,
      };
    })
    .sort((a,b) => b.pct - a.pct)
    .slice(0, 15);

  const labels  = entries.map(e => e.label);
  const data    = entries.map(e => e.pct);
  const bgs     = data.map(v => v >= 30 ? 'rgba(192,57,43,0.80)' : v >= 15 ? 'rgba(230,126,34,0.80)' : 'rgba(243,156,18,0.80)');
  const borders = data.map(v => v >= 30 ? '#c0392b' : v >= 15 ? '#e67e22' : '#f39c12');

  destroyChart(chartAbsenteismoEsp);
  chartAbsenteismoEsp = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '% Absenteísmo',
        data,
        backgroundColor: bgs,
        borderColor:     borders,
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
        tooltip: {
          ...TOOLTIP_BASE,
          callbacks: {
            label: (ctx) => {
              const e = entries[ctx.dataIndex];
              return [` ${ctx.raw}% de absenteísmo`, ` ${fmt(e.fal)} faltosos de ${fmt(e.total)} registros`];
            }
          }
        },
        datalabels: {
          anchor: 'end', align: 'end', clamp: true,
          color: '#3d5166',
          font: { family: 'Inter', size: 10, weight: 'bold' },
          formatter: val => val + '%'
        }
      },
      layout: { padding: { right: 44 } },
      scales: {
        y: {
          ticks: {
            font: { family: 'Inter', size: 10 }, color: '#3d5166',
            callback(val) {
              const label = this.getLabelForValue(val);
              return label && label.length > 22 ? label.substring(0,20)+'…' : label;
            }
          },
          grid: { display: false }
        },
        x: {
          beginAtZero: true, max: 100,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6', callback: v => v + '%' },
          grid: { display: false }
        }
      }
    }
  });
}

// ============================================================
// GRÁFICO 8: % Absenteísmo por Distrito
// ============================================================
function renderChartAbsenteismoDist() {
  const ctx = document.getElementById('chartAbsenteismoDist')?.getContext('2d');
  if (!ctx) return;

  const map = {};
  filteredData.forEach(r => {
    const key = r.distrito || 'OUTROS';
    if (!map[key]) map[key] = { fal: 0, rec: 0, can: 0 };
    
    if (r.situacao === 'FAL') map[key].fal++;
    else if (r.situacao === 'REC') map[key].rec++;
    else if (r.situacao === 'CAN') map[key].can++;
  });

  const entries = Object.entries(map)
    .filter(([, v]) => (v.rec + v.fal + v.can) > 0)
    .map(([k, v]) => {
      const totalAgendamentos = v.rec + v.fal + v.can;
      return {
        label: k,
        pct:   parseFloat((v.fal / totalAgendamentos * 100).toFixed(1)),
        fal:   v.fal,
        total: totalAgendamentos,
      };
    })
    .sort((a,b) => b.pct - a.pct);

  const labels  = entries.map(e => e.label);
  const data    = entries.map(e => e.pct);
  const bgs     = data.map(v => v >= 30 ? 'rgba(192,57,43,0.80)' : v >= 15 ? 'rgba(230,126,34,0.80)' : 'rgba(243,156,18,0.80)');
  const borders = data.map(v => v >= 30 ? '#c0392b' : v >= 15 ? '#e67e22' : '#f39c12');

  destroyChart(chartAbsenteismoDist);
  chartAbsenteismoDist = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '% Absenteísmo',
        data,
        backgroundColor: bgs,
        borderColor:     borders,
        borderWidth: 2,
        borderRadius: 7,
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
            label: (ctx) => {
              const e = entries[ctx.dataIndex];
              return [` ${ctx.raw}% de absenteísmo`, ` ${fmt(e.fal)} faltosos de ${fmt(e.total)} registros`];
            }
          }
        },
        datalabels: {
          anchor: 'end', align: 'end', clamp: true,
          color: '#1a2a3a',
          font: { family: 'Inter', size: 11, weight: 'bold' },
          formatter: val => val + '%'
        }
      },
      layout: { padding: { top: 28 } },
      scales: {
        x: {
          ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 35 },
          grid: { display: false }
        },
        y: {
          beginAtZero: true, max: 100,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6', callback: v => v + '%' },
          grid: { display: false }
        }
      }
    }
  });
}

// ============================================================
// GRÁFICO 9: % Cancelamentos por Distrito
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

  const entries = Object.entries(map)
    .filter(([, v]) => v.total > 0)
    .map(([k, v]) => ({
      label: k,
      pct:   parseFloat((v.can / v.total * 100).toFixed(1)),
      can:   v.can,
      total: v.total,
    }))
    .sort((a,b) => b.pct - a.pct);

  const labels  = entries.map(e => e.label);
  const data    = entries.map(e => e.pct);
  const bgs     = data.map((_,i) => PALETTE_ROSE[i % PALETTE_ROSE.length] + 'bb');
  const borders = data.map((_,i) => PALETTE_ROSE[i % PALETTE_ROSE.length]);

  destroyChart(chartCancelamentosDist);
  chartCancelamentosDist = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '% Cancelamentos',
        data,
        backgroundColor: bgs,
        borderColor:     borders,
        borderWidth: 2,
        borderRadius: 7,
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
            label: (ctx) => {
              const e = entries[ctx.dataIndex];
              return [` ${ctx.raw}% cancelados`, ` ${fmt(e.can)} cancelamentos de ${fmt(e.total)} registros`];
            }
          }
        },
        datalabels: {
          anchor: 'end', align: 'end', clamp: true,
          color: '#1a2a3a',
          font: { family: 'Inter', size: 11, weight: 'bold' },
          formatter: val => val + '%'
        }
      },
      layout: { padding: { top: 28 } },
      scales: {
        x: {
          ticks: { font: { family: 'Inter', size: 10, weight: '600' }, color: '#3d5166', maxRotation: 35 },
          grid: { display: false }
        },
        y: {
          beginAtZero: true, max: 100,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6', callback: v => v + '%' },
          grid: { display: false }
        }
      }
    }
  });
}

// ============================================================
// GRÁFICO 10: % Cancelamentos por Especialidade (Barras Horizontais)
// ============================================================
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

  const entries = Object.entries(map)
    .filter(([, v]) => v.total > 0)
    .map(([k, v]) => ({
      label: k,
      pct:   parseFloat((v.can / v.total * 100).toFixed(1)),
      can:   v.can,
      total: v.total,
    }))
    .sort((a,b) => b.pct - a.pct)
    .slice(0, 15);

  const labels  = entries.map(e => e.label);
  const data    = entries.map(e => e.pct);
  const bgs     = data.map((_,i) => PALETTE_GRAPE[i % PALETTE_GRAPE.length] + 'bb');
  const borders = data.map((_,i) => PALETTE_GRAPE[i % PALETTE_GRAPE.length]);

  destroyChart(chartCancelamentosEsp);
  chartCancelamentosEsp = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: '% Cancelamentos',
        data,
        backgroundColor: bgs,
        borderColor:     borders,
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
        tooltip: {
          ...TOOLTIP_BASE,
          callbacks: {
            label: (ctx) => {
              const e = entries[ctx.dataIndex];
              return [` ${ctx.raw}% cancelados`, ` ${fmt(e.can)} cancelamentos de ${fmt(e.total)} registros`];
            }
          }
        },
        datalabels: {
          anchor: 'end', align: 'end', clamp: true,
          color: '#3d5166',
          font: { family: 'Inter', size: 13, weight: '800' },
          formatter: val => val + '%'
        }
      },
      layout: { padding: { right: 54 } },
      scales: {
        y: {
          ticks: {
            font: { family: 'Inter', size: 10 }, color: '#3d5166',
            callback(val) {
              const label = this.getLabelForValue(val);
              return label && label.length > 22 ? label.substring(0,20)+'…' : label;
            }
          },
          grid: { display: false }
        },
        x: {
          beginAtZero: true, max: 100,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6', callback: v => v + '%' },
          grid: { display: false }
        }
      }
    }
  });
}

// ============================================================
// TABELA CONSOLIDADA
// ============================================================
function buildTableData() {
  const map = {};
  filteredData.forEach(r => {
    const key = `${r.distrito}|||${r.unidadeSolicitante}|||${r.tipoAtendimento}|||${r.cbo}|||${r.unidadeExecutante}`;
    if (!map[key]) {
      map[key] = {
        distrito:          r.distrito,
        unidadeSolicitante:r.unidadeSolicitante,
        tipoServico:       r.tipoAtendimento,
        cbo:               r.cbo,
        unidadeExecutante: r.unidadeExecutante,
        age: 0, rec: 0, fal: 0, can: 0, tra: 0,
      };
    }
    const sit = r.situacao;
    if      (sit === 'AGE') map[key].age++;
    else if (sit === 'REC') map[key].rec++;
    else if (sit === 'FAL') map[key].fal++;
    else if (sit === 'CAN') map[key].can++;
    else if (sit === 'TRA') map[key].tra++;
  });

  tableData = Object.values(map).map(r => {
    const totalAgendamentos = r.rec + r.fal + r.can;
    const pctAbsenteismo    = totalAgendamentos > 0
      ? parseFloat((r.fal / totalAgendamentos * 100).toFixed(1))
      : 0;
    return { ...r, totalAgendamentos, pctAbsenteismo };
  }).sort((a,b) => b.totalAgendamentos - a.totalAgendamentos);

  tableSearched = [...tableData];
}

function filterTable() {
  const q = (document.getElementById('tableSearch')?.value || '').toLowerCase();
  tableSearched = !q
    ? [...tableData]
    : tableData.filter(r =>
        (r.distrito          ||'').toLowerCase().includes(q) ||
        (r.unidadeSolicitante||'').toLowerCase().includes(q) ||
        (r.tipoServico       ||'').toLowerCase().includes(q) ||
        (r.cbo               ||'').toLowerCase().includes(q) ||
        (r.unidadeExecutante ||'').toLowerCase().includes(q)
      );
  currentPage = 1;
  renderTable();
}

function sortTable(col) {
  if (sortColIdx === col) sortAscFlag = !sortAscFlag;
  else { sortColIdx = col; sortAscFlag = true; }

  const keys = [
    'distrito','unidadeSolicitante','tipoServico','cbo','unidadeExecutante',
    'age','rec','fal','can','tra','totalAgendamentos','pctAbsenteismo'
  ];
  const key = keys[col];
  tableSearched.sort((a,b) => {
    const va = a[key] ?? '';
    const vb = b[key] ?? '';
    const cmp = typeof va === 'number'
      ? va - vb
      : va.toString().localeCompare(vb.toString(), 'pt-BR');
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
  const total    = tableSearched.length;
  const pages    = Math.max(1, Math.ceil(total / pageSize));
  if (currentPage > pages) currentPage = pages;

  const start = (currentPage - 1) * pageSize;
  const slice = tableSearched.slice(start, start + pageSize);

  const tbody = document.getElementById('tableBody');
  const tfoot = document.getElementById('tableFoot');
  if (!tbody) return;

  if (slice.length === 0) {
    tbody.innerHTML = '<tr><td colspan="12" class="empty-msg">Nenhum registro encontrado.</td></tr>';
    tfoot.innerHTML = '';
  } else {
    tbody.innerHTML = slice.map(r => `
      <tr>
        <td><span style="font-size:0.78rem;font-weight:700;color:#1e3a5f;background:rgba(30,58,95,0.08);border-radius:6px;padding:2px 8px;">${r.distrito || '–'}</span></td>
        <td style="font-size:0.78rem;color:#3d5166;">${r.unidadeSolicitante || '–'}</td>
        <td><span style="font-size:0.76rem;font-weight:600;color:${r.tipoServico === 'Primeira Consulta' ? '#1a7a3f' : '#7d3c98'}">${r.tipoServico || '–'}</span></td>
        <td><div style="font-weight:600;color:#1e3a5f;font-size:0.8rem;">${r.cbo || '–'}</div></td>
        <td style="font-size:0.78rem;color:#3d5166;">${r.unidadeExecutante || '–'}</td>
        <td class="text-center"><span class="badge-num badge-age">${fmt(r.age)}</span></td>
        <td class="text-center"><span class="badge-num badge-rec">${fmt(r.rec)}</span></td>
        <td class="text-center"><span class="badge-num badge-fal">${fmt(r.fal)}</span></td>
        <td class="text-center"><span class="badge-num badge-can">${fmt(r.can)}</span></td>
        <td class="text-center"><span class="badge-num badge-tra">${fmt(r.tra)}</span></td>
        <td class="text-center"><span class="badge-num badge-total">${fmt(r.totalAgendamentos)}</span></td>
        <td class="text-center"><span class="${absentClass(r.pctAbsenteismo)}">${r.pctAbsenteismo.toFixed(1)}%</span></td>
      </tr>
    `).join('');

    const sAge   = tableSearched.reduce((s,r) => s + r.age,   0);
    const sRec   = tableSearched.reduce((s,r) => s + r.rec,   0);
    const sFal   = tableSearched.reduce((s,r) => s + r.fal,   0);
    const sCan   = tableSearched.reduce((s,r) => s + r.can,   0);
    const sTra   = tableSearched.reduce((s,r) => s + r.tra,   0);
    const sTotal = tableSearched.reduce((s,r) => s + r.totalAgendamentos, 0);
    const pctGeral = sTotal > 0 ? parseFloat((sFal / sTotal * 100).toFixed(1)) : 0;

    tfoot.innerHTML = `
      <tr>
        <td colspan="5"><i class="fas fa-calculator" style="margin-right:6px;"></i>TOTAL GERAL (${fmt(tableSearched.length)} linhas)</td>
        <td class="text-center">${fmt(sAge)}</td>
        <td class="text-center">${fmt(sRec)}</td>
        <td class="text-center">${fmt(sFal)}</td>
        <td class="text-center">${fmt(sCan)}</td>
        <td class="text-center">${fmt(sTra)}</td>
        <td class="text-center">${fmt(sTotal)}</td>
        <td class="text-center">${pctGeral.toFixed(1)}%</td>
      </tr>
    `;
  }

  document.getElementById('tablePaginationInfo').textContent =
    `Mostrando ${total === 0 ? 0 : start+1} a ${Math.min(start+pageSize, total)} de ${fmt(total)} registros`;

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
  const pages    = Math.max(1, Math.ceil(tableSearched.length / pageSize));
  if (p < 1 || p > pages) return;
  currentPage = p;
  renderTable();
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
        'Unidade Executante':  r.unidadeExecutante,
        'Unidade Solicitante': r.unidadeSolicitante,
        'Distrito':            r.distrito,
        'Especialidade (CBO)': r.cbo,
        'Tipo Especialidade':  r.especialidade,
        'Profissional':        r.profissional,
        'Tipo Atendimento':    r.tipoAtendimento,
        'Situação':            r.situacaoLabel,
        'Operador':            r.operador,
        'Data Agenda':         r.dataAgenda,
        'Data Criação':        r.dataCriacao,
        'Mês Agendamento':     r.mesAgendamento,
      }));

      const wsSummary = tableData.map(r => ({
        'Distrito':               r.distrito,
        'Unidade Solicitante':    r.unidadeSolicitante,
        'Tipo de Serviço':        r.tipoServico,
        'CBO / Especialidade':    r.cbo,
        'Unidade Executante':     r.unidadeExecutante,
        'AGE':                    r.age,
        'REC (Recepcionados)':    r.rec,
        'FAL (Faltosos)':         r.fal,
        'CAN (Cancelados)':       r.can,
        'TRA (Transferidos)':     r.tra,
        'Total Agendamentos (REC+FAL+CAN)': r.totalAgendamentos,
        '% Absenteísmo':          r.pctAbsenteismo + '%',
      }));

      const wb  = XLSX.utils.book_new();
      const ws1 = XLSX.utils.json_to_sheet(wsData);
      autoSizeColumns(ws1, wsData);
      XLSX.utils.book_append_sheet(wb, ws1, 'Dados Filtrados');

      const ws2 = XLSX.utils.json_to_sheet(wsSummary);
      autoSizeColumns(ws2, wsSummary);
      XLSX.utils.book_append_sheet(wb, ws2, 'Consolidado Agendamentos');

      const now   = new Date();
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
    wch: Math.min(
      data.reduce((max, row) => Math.max(max, (row[col]||'').toString().length), col.length) + 2,
      60
    )
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
  toast.style.cssText = `
    position:fixed;bottom:24px;right:24px;z-index:9998;
    background:#c0392b;color:#fff;border-radius:12px;
    padding:14px 22px;font-family:Inter,sans-serif;font-size:.85rem;
    font-weight:600;box-shadow:0 6px 24px rgba(0,0,0,.3);
    display:flex;align-items:center;gap:10px;
  `;
  toast.innerHTML = `<i class="fas fa-exclamation-circle"></i> ${msg}`;
  document.body.appendChild(toast);
  setTimeout(() => toast.remove(), 7000);
}

function updateLastUpdate() {
  const el = document.getElementById('lastUpdate');
  if (el) el.textContent = `Última atualização: ${new Date().toLocaleString('pt-BR')}`;
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
// INICIALIZAÇÃO
// ============================================================
document.addEventListener('DOMContentLoaded', () => {
  initDatePickers();
  loadData();
});


