/* ============================================================
   DIRETORIA DE REGULAÇÃO DO ACESSO – DASHBOARD
   script.js – v3.1 (COMPLETO)
   ============================================================ */

'use strict';

Chart.register(ChartDataLabels);

// ============================================================
// CONFIGURAÇÕES
// ============================================================
const SHEET_ID = '1puNbYysRBj-5CY6fhnHNnYd9OH96cl7guMFBOLeYZV4';
const SHEET_GID = '1698493941';

const CSV_URL_GVIZ = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&gid=${SHEET_GID}`;
const CSV_URL_EXPORT = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${SHEET_GID}`;

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
const MESES_PT = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                  'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'];

// ============================================================
// ESTADO GLOBAL
// ============================================================
let allData = [];
let filteredData = [];
let tableData = [];
let tableSearched = [];
let currentPage = 1;
let sortColIdx = -1;
let sortAscFlag = true;
let chartPrestador = null;
let chartTipoAtendimento = null;

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
  const artigos = ['DE', 'DA', 'DO', 'DAS', 'DOS', 'E', 'A', 'O', 'EM', 'NO', 'NA', 'NOS', 'NAS', 'POR', 'COM', 'PARA'];
  return str.toString().toLowerCase().split(' ').map((w, i) => {
    if (i === 0 || !artigos.includes(w.toUpperCase())) {
      return w.charAt(0).toUpperCase() + w.slice(1);
    }
    return w;
  }).join(' ');
}

function formatCBO(nomeCBO, especialidade) {
  if (!nomeCBO) return '';
  let name = nomeCBO.toString().trim();
  name = name.replace(/^M[eé]dico[:\-\s]*/i, '').trim();
  name = titleCase(name);
  if (especialidade) {
    const esp = especialidade.toString().trim();
    const addons = ['Adulto', 'Infantil', 'Geral', 'Pediátrico', 'Pediátrica'];
    for (const addon of addons) {
      if (norm(esp).includes(norm(addon)) && !norm(name).includes(norm(addon))) {
        name += ' ' + addon;
        break;
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
  const solNorm = norm(solicitante);
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
  const map = { AGE: 'Agendados', CAN: 'Cancelados', FAL: 'Faltosos', REC: 'Recepcionados', TRA: 'Transferidos' };
  return map[(val || '').toUpperCase()] || val || '–';
}

function getOperador(codigo) {
  const key = (codigo || '').toString().trim();
  return OPERADORES[key] || 'Outros';
}

function fmt(n) {
  return (n || 0).toLocaleString('pt-BR');
}

function parseDate(str) {
  if (!str) return null;
  str = str.toString().trim();
  let m = str.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return new Date(+m[3], +m[2] - 1, +m[1]);
  m = str.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return new Date(+m[1], +m[2] - 1, +m[3]);
  return null;
}

function isSameDay(d1, d2) {
  if (!d1 || !d2) return false;
  return d1.getFullYear() === d2.getFullYear() &&
         d1.getMonth() === d2.getMonth() &&
         d1.getDate() === d2.getDate();
}

// ============================================================
// CARREGAR DADOS
// ============================================================
function looksLikeHtml(text) {
  if (!text) return false;
  const t = text.trim().slice(0, 200).toLowerCase();
  return t.startsWith('<!doctype') || t.startsWith('<html') || t.includes('<head') || t.includes('<body');
}

async function fetchCsvText(urlBase) {
  const url = urlBase + (urlBase.includes('?') ? '&' : '?') + 't=' + Date.now();
  const response = await fetch(url, { cache: 'no-store', mode: 'cors' });
  if (!response.ok) throw new Error(`HTTP ${response.status}`);
  return await response.text();
}

async function loadData() {
  showLoading(true);
  setStatus('Carregando...', false);
  const icon = document.getElementById('refreshIcon');
  if (icon) icon.classList.add('spinning');

  try {
    let text = await fetchCsvText(CSV_URL_GVIZ);

    if (looksLikeHtml(text)) {
      text = await fetchCsvText(CSV_URL_EXPORT);
    }

    if (looksLikeHtml(text)) {
      throw new Error('Resposta HTML (provável permissão/restrição).');
    }

    Papa.parse(text, {
      header: true,
      skipEmptyLines: true,
      dynamicTyping: false,
      transformHeader: h => h.trim(),
      complete: function(results) {
        allData = normalizeData(results.data);
        populateFilterOptions();
        applyFilters();
        setStatus('Conectado', true);
        updateLastUpdate();
        showLoading(false);
        if (icon) icon.classList.remove('spinning');
      },
      error: function(err) {
        console.error('Parse error:', err);
        showError('Erro ao processar os dados.');
        if (icon) icon.classList.remove('spinning');
        showLoading(false);
      }
    });
  } catch (err) {
    console.error('Fetch error:', err);
    showError('Não foi possível carregar os dados. Verifique as permissões da planilha.');
    if (icon) icon.classList.remove('spinning');
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
      'UNIDADE SOLICITANTE': get('UNIDADE SOLICITANTE'),
      'UNIDADE DE REFERÊNCIA': get('UNIDADE DE REFERÊNCIA', 'UNIDADE DE REFERENCIA'),
    });

    const nomeCBO = get('NOME CBO');
    const especialidade = get('ESPECIALIDADE');
    const cbof = formatCBO(nomeCBO, especialidade);
    const profissional = formatProfissional(get('NOME PROFISSIONAL'));
    const tipoAtend = getTipoAtendimento(get('TIPO DE ATENDIMENTO'));
    const situacao = (get('SITUAÇÃO', 'SITUACAO') || '').toUpperCase().trim();
    const operCod = get('OPERADOR AGENDAMENTO');

    const dataCriacao = get('DATA CRIAÇÃO DO AGENDAMENTO', 'DATA CRIACAO DO AGENDAMENTO', 'DATA CRIAÇÃO', 'DATA_CRIACAO');
    const dataCriacaoParsed = parseDate(dataCriacao);

    const dataAgenda = get('DATA AGENDA', 'DATA_AGENDA');
    const dataAgendaParsed = parseDate(dataAgenda);

    const unidadeExec = get('UNIDADE EXECUTANTE');
    const distrito = getDistrito(unidadeSolicitante);

    let mesLabel = '';
    if (dataAgendaParsed) {
      mesLabel = MESES_PT[dataAgendaParsed.getMonth()] + '/' + dataAgendaParsed.getFullYear();
    }

    return {
      _raw: row,
      unidadeExecutante: unidadeExec,
      unidadeSolicitante: unidadeSolicitante,
      cbo: cbof,
      especialidade: especialidade ? titleCase(especialidade.toString()) : '',
      nomeCBO,
      profissional,
      tipoAtendimento: tipoAtend,
      situacao,
      situacaoLabel: getSituacaoLabel(situacao),
      operadorCod: operCod,
      operador: getOperador(operCod),
      dataCriacao,
      dataCriacaoParsed,
      dataAgenda,
      dataAgendaParsed,
      mesAgendamento: mesLabel,
      distrito,
      valor: parseFloat((get('VALOR') || '0').toString().replace(',', '.')) || 0,
      quantidade: parseInt(get('QUANTIDADE') || '1') || 1,
    };
  }).filter(r => r.unidadeExecutante || r.profissional);
}

// ============================================================
// POPULAR FILTROS
// ============================================================
function populateFilterOptions() {
  populateSelect('filterPrestador', [...new Set(allData.map(r => r.unidadeExecutante).filter(Boolean))].sort());
  populateSelect('filterUnidadeSolicitante', [...new Set(allData.map(r => r.unidadeSolicitante).filter(Boolean))].sort());
  populateSelect('filterEspecialidade', [...new Set(allData.map(r => r.cbo).filter(Boolean))].sort());
  populateSelect('filterProfissional', [...new Set(allData.map(r => r.profissional).filter(Boolean))].sort());

  const mesesSet = {};
  allData.forEach(r => {
    if (r.dataAgendaParsed && r.mesAgendamento) {
      const d = r.dataAgendaParsed;
      const key = d.getFullYear() * 100 + d.getMonth();
      mesesSet[key] = r.mesAgendamento;
    }
  });
  const mesesOrdenados = Object.entries(mesesSet).sort((a, b) => +a[0] - +b[0]).map(e => e[1]);
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
    opt.value = v;
    opt.textContent = v;
    sel.appendChild(opt);
  });
  if (current && [...sel.options].some(o => o.value === current)) sel.value = current;
}

// ============================================================
// APLICAR FILTROS
// ============================================================
function applyFilters() {
  const prestador = document.getElementById('filterPrestador')?.value || '';
  const unidadeSol = document.getElementById('filterUnidadeSolicitante')?.value || '';
  const especialidade = document.getElementById('filterEspecialidade')?.value || '';
  const tipoServico = document.getElementById('filterTipoServico')?.value || '';
  const profissional = document.getElementById('filterProfissional')?.value || '';
  const mes = document.getElementById('filterMes')?.value || '';
  const dataCriacaoSelecionada = window._fpInicio ? window._fpInicio.selectedDates[0] : null;

  filteredData = allData.filter(r => {
    if (prestador && r.unidadeExecutante !== prestador) return false;
    if (unidadeSol && r.unidadeSolicitante !== unidadeSol) return false;
    if (especialidade && r.cbo !== especialidade) return false;
    if (tipoServico && r.tipoAtendimento !== tipoServico) return false;
    if (profissional && r.profissional !== profissional) return false;
    if (mes && r.mesAgendamento !== mes) return false;
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
  ['filterPrestador', 'filterUnidadeSolicitante', 'filterEspecialidade',
   'filterTipoServico', 'filterProfissional', 'filterMes'].forEach(id => {
    const el = document.getElementById(id);
    if (el) el.value = '';
  });
  if (window._fpInicio) window._fpInicio.clear();
  applyFilters();
}

// ============================================================
// KPIS
// ============================================================
function updateKPIs() {
  const total = filteredData.length;

  const hoje = new Date();
  const mesAtual = hoje.getMonth();
  const anoAtual = hoje.getFullYear();
  const mesAtualStr = MESES_PT[mesAtual] + '/' + anoAtual;

  let proximoMes = mesAtual + 1;
  let proximoAno = anoAtual;
  if (proximoMes > 11) {
    proximoMes = 0;
    proximoAno = anoAtual + 1;
  }
  const proximoMesStr = MESES_PT[proximoMes] + '/' + proximoAno;

  const ofertaMesAtual = filteredData.filter(r => r.mesAgendamento === mesAtualStr).length;
  const ofertaProximoMes = filteredData.filter(r => r.mesAgendamento === proximoMesStr).length;

  animateCount('kpiTotalGeral', total);
  animateCount('kpiMesAtual', ofertaMesAtual);
  animateCount('kpiProximoMes', ofertaProximoMes);
}

function animateCount(id, target) {
  const el = document.getElementById(id);
  if (!el) return;
  const current = parseInt(el.textContent.replace(/\D/g, '')) || 0;
  if (current === target) {
    el.textContent = fmt(target);
    return;
  }
  const step = Math.max(1, Math.round(Math.abs(target - current) / 20));
  let val = current;
  const inc = target > current ? step : -step;
  const timer = setInterval(() => {
    val += inc;
    if ((inc > 0 && val >= target) || (inc < 0 && val <= target)) {
      val = target;
      clearInterval(timer);
    }
    el.textContent = fmt(val);
  }, 16);
}

// ============================================================
// GRÁFICOS
// ============================================================
const PALETTE_BLUE = ['#1e3a5f', '#2d5494', '#4a90d9', '#74b3e8', '#a8d1f5', '#c8e4fb'];
const DARK_BLUE = '#1e3a5f';

const TOOLTIP_BASE = {
  backgroundColor: 'rgba(20,40,68,0.92)',
  titleFont: { family: 'Inter', size: 12, weight: '700' },
  bodyFont: { family: 'Inter', size: 11 },
  padding: 12,
  cornerRadius: 10,
  callbacks: { label: ctx => ` ${fmt(ctx.raw)}` }
};

function countBy(data, keyFn) {
  const map = {};
  data.forEach(r => {
    const k = keyFn(r);
    if (!k) return;
    map[k] = (map[k] || 0) + 1;
  });
  return map;
}

function sortedEntries(obj, limit = 0) {
  let entries = Object.entries(obj).sort((a, b) => b[1] - a[1]);
  if (limit > 0) entries = entries.slice(0, limit);
  return entries;
}

function renderAllCharts() {
  renderChartPrestador();
  renderChartTipoAtendimento();
}

function renderChartPrestador() {
  const ctx = document.getElementById('chartPrestador')?.getContext('2d');
  if (!ctx) return;

  const counts = countBy(filteredData, r => r.unidadeExecutante);
  const entries = sortedEntries(counts, 12);
  const labels = entries.map(e => e[0]);
  const data = entries.map(e => e[1]);

  if (chartPrestador) chartPrestador.destroy();
  chartPrestador = new Chart(ctx, {
    type: 'bar',
    data: {
      labels,
      datasets: [{
        label: 'Agendamentos',
        data,
        backgroundColor: labels.map((_, i) => PALETTE_BLUE[i % PALETTE_BLUE.length] + 'dd'),
        borderColor: labels.map((_, i) => PALETTE_BLUE[i % PALETTE_BLUE.length]),
        borderWidth: 2,
        borderRadius: 6,
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
          anchor: 'end',
          align: 'end',
          clamp: true,
          offset: 2,
          color: '#1a2a3a',
          font: { family: 'Inter', size: 14, weight: 'bold' },
          formatter: val => fmt(val)
        }
      },
      layout: { padding: { top: 30 } },
      scales: {
        x: {
          ticks: {
            font: { family: 'Inter', size: 10 },
            color: '#3d5166',
            maxRotation: 35,
            minRotation: 20,
            callback: function(val) {
              const label = this.getLabelForValue(val);
              return label && label.length > 18 ? label.substring(0, 16) + '…' : label;
            }
          },
          grid: { display: false }
        },
        y: {
          beginAtZero: true,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' },
          grid: { color: 'rgba(0,0,0,0.05)' }
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

  if (chartTipoAtendimento) chartTipoAtendimento.destroy();
  chartTipoAtendimento = new Chart(ctx, {
    type: 'bar',
    data: {
      labels: ['1ª Consulta', 'Retorno'],
      datasets: [{
        label: 'Quantidade',
        data: [pc, ret],
        backgroundColor: ['rgba(30,58,95,0.88)', 'rgba(39,174,96,0.88)'],
        borderColor: ['#1e3a5f', '#27ae60'],
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
          callbacks: { label: ctx => ` ${fmt(ctx.raw)} agendamentos` }
        },
        datalabels: {
          anchor: 'center',
          align: 'center',
          color: '#ffffff',
          font: { family: 'Inter', size: 18, weight: 'bold' },
          formatter: val => val > 0 ? fmt(val) : ''
        }
      },
      scales: {
        x: {
          ticks: {
            font: { family: 'Inter', size: 15, weight: '700' },
            color: '#1e3a5f'
          },
          grid: { display: false }
        },
        y: {
          beginAtZero: true,
          ticks: { font: { family: 'Inter', size: 10 }, color: '#7a8fa6' },
          grid: { color: 'rgba(0,0,0,0.05)' }
        }
      }
    }
  });
}

// ============================================================
// TABELA
// ============================================================
function buildTableData() {
  const map = {};
  filteredData.forEach(r => {
    const key = `${r.unidadeExecutante}|||${r.unidadeSolicitante}|||${r.cbo}|||${r.especialidade}|||${r.profissional}|||${r.operador}`;
    if (!map[key]) {
      map[key] = {
        unidadeExecutante: r.unidadeExecutante,
        unidadeSolicitante: r.unidadeSolicitante,
        cbo: r.cbo,
        especialidade: r.especialidade,
        profissional: r.profissional,
        operador: r.operador,
        primeira: 0,
        retorno: 0,
        total: 0
      };
    }
    if (r.tipoAtendimento === 'Primeira Consulta') map[key].primeira++;
    else if (r.tipoAtendimento === 'Retorno') map[key].retorno++;
    map[key].total++;
  });

  tableData = Object.values(map).sort((a, b) => b.total - a.total);
  tableSearched = [...tableData];
}

function filterTable() {
  const q = (document.getElementById('tableSearch')?.value || '').toLowerCase();
  tableSearched = !q
    ? [...tableData]
    : tableData.filter(r =>
        (r.unidadeExecutante || '').toLowerCase().includes(q) ||
        (r.unidadeSolicitante || '').toLowerCase().includes(q) ||
        (r.cbo || '').toLowerCase().includes(q) ||
        (r.profissional || '').toLowerCase().includes(q) ||
        (r.especialidade || '').toLowerCase().includes(q) ||
        (r.operador || '').toLowerCase().includes(q)
      );
  currentPage = 1;
  renderTable();
}

function sortTable(col) {
  if (sortColIdx === col) sortAscFlag = !sortAscFlag;
  else {
    sortColIdx = col;
    sortAscFlag = true;
  }

  const keys = ['unidadeExecutante', 'unidadeSolicitante', 'cbo', 'profissional', 'primeira', 'retorno', 'total', 'operador'];
  const key = keys[col];
  tableSearched.sort((a, b) => {
    const va = a[key] ?? '';
    const vb = b[key] ?? '';
    const cmp = typeof va === 'number' ? va - vb : va.toString().localeCompare(vb.toString(), 'pt-BR');
    return sortAscFlag ? cmp : -cmp;
  });
  renderTable();
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

  if (slice.length === 0) {
    tbody.innerHTML = '<tr><td colspan="8" class="empty-msg">Nenhum registro encontrado.</td></tr>';
    tfoot.innerHTML = '';
  } else {
    tbody.innerHTML = slice.map(r => `
      <tr>
        <td>${r.unidadeExecutante || '–'}</td>
        <td>${r.unidadeSolicitante || '–'}</td>
        <td>
          <div style="font-weight:600;color:#1e3a5f;">${r.cbo || '–'}</div>
          ${r.especialidade ? `<div style="font-size:0.74rem;color:#7a8fa6;margin-top:2px;">${r.especialidade}</div>` : ''}
        </td>
        <td>${r.profissional || '–'}</td>
        <td class="text-center"><span class="badge-num">${fmt(r.primeira)}</span></td>
        <td class="text-center"><span class="badge-num" style="background:rgba(39,174,96,0.15);color:#1e7a3f;">${fmt(r.retorno)}</span></td>
        <td class="text-center"><span class="badge-num badge-total">${fmt(r.total)}</span></td>
        <td style="font-size:0.8rem;color:#3d5166;">${r.operador || '–'}</td>
      </tr>
    `).join('');

    const totPrimeira = tableSearched.reduce((s, r) => s + r.primeira, 0);
    const totRetorno = tableSearched.reduce((s, r) => s + r.retorno, 0);
    const totGeral = tableSearched.reduce((s, r) => s + r.total, 0);

    tfoot.innerHTML = `
      <tr>
        <td colspan="2"><i class="fas fa-calculator" style="margin-right:6px;"></i>TOTAL GERAL (${fmt(tableSearched.length)} registros)</td>
        <td colspan="2">–</td>
        <td class="text-center"><strong>${fmt(totPrimeira)}</strong></td>
        <td class="text-center"><strong>${fmt(totRetorno)}</strong></td>
        <td class="text-center"><strong>${fmt(totGeral)}</strong></td>
        <td>–</td>
      </tr>
    `;
  }

  document.getElementById('tablePaginationInfo').textContent =
    `Mostrando ${total === 0 ? 0 : start + 1} a ${Math.min(start + pageSize, total)} de ${fmt(total)} registros`;

  renderPagination(currentPage, pages);
}

function renderPagination(cur, total) {
  const container = document.getElementById('pagination');
  if (!container) return;

  let html = `<button class="page-btn" onclick="goPage(${cur - 1})" ${cur === 1 ? 'disabled' : ''}>‹</button>`;

  let pages = [];
  if (total <= 7) {
    for (let i = 1; i <= total; i++) pages.push(i);
  } else {
    pages = [1];
    if (cur > 3) pages.push('...');
    for (let i = Math.max(2, cur - 1); i <= Math.min(total - 1, cur + 1); i++) pages.push(i);
    if (cur < total - 2) pages.push('...');
    pages.push(total);
  }

  pages.forEach(p => {
    if (p === '...') {
      html += `<button class="page-btn" disabled>…</button>`;
    } else {
      html += `<button class="page-btn ${p === cur ? 'active' : ''}" onclick="goPage(${p})">${p}</button>`;
    }
  });

  html += `<button class="page-btn" onclick="goPage(${cur + 1})" ${cur === total ? 'disabled' : ''}>›</button>`;
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
// EXPORTAR EXCEL
// ============================================================
function exportExcel() {
  if (!filteredData.length) {
    alert('Nenhum dado para exportar.');
    return;
  }
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
        'Mês Agendamento': r.mesAgendamento
      }));

      const wsSummary = tableData.map(r => ({
        'Unidade Executante': r.unidadeExecutante,
        'Unidade Solicitante': r.unidadeSolicitante,
        'CBO / Especialidade': r.cbo,
        'Tipo de Especialidade': r.especialidade,
        'Profissional': r.profissional,
        'Total 1ª Consulta': r.primeira,
        'Total Retorno': r.retorno,
        'Total Geral': r.total,
        'Operador': r.operador
      }));

      const wb = XLSX.utils.book_new();
      const ws1 = XLSX.utils.json_to_sheet(wsData);
      autoSizeColumns(ws1, wsData);
      XLSX.utils.book_append_sheet(wb, ws1, 'Dados Filtrados');

      const ws2 = XLSX.utils.json_to_sheet(wsSummary);
      autoSizeColumns(ws2, wsSummary);
      XLSX.utils.book_append_sheet(wb, ws2, 'Resumo por Profissional');

      const now = new Date();
      const fname = `Agendamentos_CMC_${now.getFullYear()}${String(now.getMonth() + 1).padStart(2, '0')}${String(now.getDate()).padStart(2, '0')}.xlsx`;
      XLSX.writeFile(wb, fname);
    } catch (e) {
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
      data.reduce((max, row) => Math.max(max, (row[col] || '').toString().length), col.length) + 2,
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
  const el = document.getElementById('statusText');
  const dot = document.querySelector('.status-dot');
  if (el) el.textContent = msg;
  if (dot) dot.className = 'status-dot ' + (ok ? 'connected' : 'error');
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
