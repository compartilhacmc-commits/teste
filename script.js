// Configuração
const SHEET_ID = '1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4';
const SHEET_NAME = 'pendencias eldorado';
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(SHEET_NAME)}`;

// Dados globais
let dadosOriginais = [];
let dadosFiltrados = [];
let chartUnidades = null;
let chartEspecialidades = null;
let chartVenc15 = null;
let chartVenc30 = null;
let currentPage = 1;
let pageSize = 25;

// Inicialização
document.addEventListener('DOMContentLoaded', () => {
    carregarDados();
    configurarEventos();
});

// Configurar eventos dos botões
function configurarEventos() {
    document.getElementById('btnAtualizar').addEventListener('click', () => {
        carregarDados();
    });
    
    document.getElementById('btnLimparFiltros').addEventListener('click', () => {
        limparFiltros();
    });
    
    document.getElementById('btnExportar').addEventListener('click', () => {
        exportarParaExcel();
    });
    
    document.getElementById('pageSize').addEventListener('change', (e) => {
        pageSize = parseInt(e.target.value);
        currentPage = 1;
        renderizarTabela();
    });
    
    // Eventos de filtros
    ['filtroUnidade', 'filtroPrestador', 'filtroEspecialidade', 'filtroStatus'].forEach(id => {
        document.getElementById(id).addEventListener('change', aplicarFiltros);
    });
}

// Carregar dados do Google Sheets
async function carregarDados() {
    try {
        mostrarCarregamento(true);
        
        const response = await fetch(SHEET_URL);
        const text = await response.text();
        
        // Remover prefixo do Google e parsear JSON
        const json = JSON.parse(text.substring(47).slice(0, -2));
        
        if (json.table && json.table.rows) {
            dadosOriginais = processarDados(json.table);
            dadosFiltrados = [...dadosOriginais];
            
            popularFiltros();
            aplicarFiltros();
            atualizarCards();
            renderizarGraficos();
            renderizarTabela();
        }
        
        mostrarCarregamento(false);
        mostrarNotificacao('Dados atualizados com sucesso!', 'success');
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
        mostrarCarregamento(false);
        mostrarNotificacao('Erro ao carregar dados. Verifique a conexão.', 'error');
    }
}

// Processar dados do Google Sheets
function processarDados(table) {
    const rows = table.rows;
    const dados = [];
    
    // Pular cabeçalho (primeira linha)
    for (let i = 1; i < rows.length; i++) {
        const row = rows[i];
        if (!row.c) continue;
        
        const registro = {
            numeroSolicitacao: getCellValue(row.c[0]),
            dataSolicitacao: getCellValue(row.c[1]),
            prontuario: getCellValue(row.c[2]),
            telefone: getCellValue(row.c[3]),
            unidadeSolicitante: getCellValue(row.c[9]),
            cboEspecialidade: getCellValue(row.c[10]),
            dataInicioPendencia: getCellValue(row.c[12]),
            status: getCellValue(row.c[19]),
            prestador: getCellValue(row.c[14]) || 'Não informado'
        };
        
        dados.push(registro);
    }
    
    return dados;
}

// Obter valor da célula
function getCellValue(cell) {
    if (!cell || cell.v === null || cell.v === undefined) return '';
    return cell.v.toString();
}

// Popular filtros
function popularFiltros() {
    const unidades = [...new Set(dadosOriginais.map(d => d.unidadeSolicitante).filter(v => v))].sort();
    const prestadores = [...new Set(dadosOriginais.map(d => d.prestador).filter(v => v))].sort();
    const especialidades = [...new Set(dadosOriginais.map(d => d.cboEspecialidade).filter(v => v))].sort();
    const status = [...new Set(dadosOriginais.map(d => d.status).filter(v => v))].sort();
    
    popularSelect('filtroUnidade', unidades);
    popularSelect('filtroPrestador', prestadores);
    popularSelect('filtroEspecialidade', especialidades);
    popularSelect('filtroStatus', status);
}

// Popular select
function popularSelect(id, valores) {
    const select = document.getElementById(id);
    select.innerHTML = valores.map(v => `<option value="${v}">${v}</option>`).join('');
}

// Aplicar filtros
function aplicarFiltros() {
    const unidadesSelecionadas = Array.from(document.getElementById('filtroUnidade').selectedOptions).map(o => o.value);
    const prestadoresSelecionados = Array.from(document.getElementById('filtroPrestador').selectedOptions).map(o => o.value);
    const especialidadesSelecionadas = Array.from(document.getElementById('filtroEspecialidade').selectedOptions).map(o => o.value);
    const statusSelecionados = Array.from(document.getElementById('filtroStatus').selectedOptions).map(o => o.value);
    
    dadosFiltrados = dadosOriginais.filter(d => {
        const passaUnidade = unidadesSelecionadas.length === 0 || unidadesSelecionadas.includes(d.unidadeSolicitante);
        const passaPrestador = prestadoresSelecionados.length === 0 || prestadoresSelecionados.includes(d.prestador);
        const passaEspecialidade = especialidadesSelecionadas.length === 0 || especialidadesSelecionadas.includes(d.cboEspecialidade);
        const passaStatus = statusSelecionados.length === 0 || statusSelecionados.includes(d.status);
        
        return passaUnidade && passaPrestador && passaEspecialidade && passaStatus;
    });
    
    currentPage = 1;
    atualizarCards();
    renderizarGraficos();
    renderizarTabela();
}

// Limpar filtros
function limparFiltros() {
    ['filtroUnidade', 'filtroPrestador', 'filtroEspecialidade', 'filtroStatus'].forEach(id => {
        const select = document.getElementById(id);
        Array.from(select.options).forEach(option => option.selected = false);
    });
    
    aplicarFiltros();
    mostrarNotificacao('Filtros limpos com sucesso!', 'success');
}

// Atualizar cards
function atualizarCards() {
    const total = dadosFiltrados.length;
    const totalGeral = dadosOriginais.length;
    
    // Calcular vencimentos
    const hoje = new Date();
    const venc15 = dadosFiltrados.filter(d => {
        if (!d.dataInicioPendencia) return false;
        const dataInicio = parseData(d.dataInicioPendencia);
        const diffDias = Math.floor((hoje - dataInicio) / (1000 * 60 * 60 * 24));
        return diffDias >= 15;
    }).length;
    
    const venc30 = dadosFiltrados.filter(d => {
        if (!d.dataInicioPendencia) return false;
        const dataInicio = parseData(d.dataInicioPendencia);
        const diffDias = Math.floor((hoje - dataInicio) / (1000 * 60 * 60 * 24));
        return diffDias >= 30;
    }).length;
    
    const porcentagem = totalGeral > 0 ? ((total / totalGeral) * 100).toFixed(1) : 0;
    
    document.getElementById('cardTotal').textContent = total;
    document.getElementById('cardVenc15').textContent = venc15;
    document.getElementById('cardVenc30').textContent = venc30;
    document.getElementById('cardPorcentagem').textContent = porcentagem + '%';
}

// Parse de data
function parseData(dataStr) {
    if (!dataStr) return null;
    
    // Tentar diferentes formatos
    const formatos = [
        // dd/mm/yyyy
        /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/,
        // yyyy-mm-dd
        /^(\d{4})-(\d{1,2})-(\d{1,2})$/,
        // Date(valor)
        /^Date\((\d+)\)$/
    ];
    
    for (let formato of formatos) {
        const match = dataStr.match(formato);
        if (match) {
            if (formato.source.includes('Date')) {
                // Formato Date(valor)
                return new Date(parseInt(match[1]));
            } else if (formato.source.startsWith('^\\(\\d{4}')) {
                // Formato yyyy-mm-dd
                return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
            } else {
                // Formato dd/mm/yyyy
                return new Date(parseInt(match[3]), parseInt(match[2]) - 1, parseInt(match[1]));
            }
        }
    }
    
    // Tentar Date.parse como último recurso
    const date = new Date(dataStr);
    return isNaN(date.getTime()) ? null : date;
}

// Renderizar gráficos
function renderizarGraficos() {
    renderizarGraficoUnidades();
    renderizarGraficoEspecialidades();
    renderizarGraficosVencimento();
}

// Gráfico de Unidades
function renderizarGraficoUnidades() {
    const contagem = {};
    dadosFiltrados.forEach(d => {
        if (d.unidadeSolicitante) {
            contagem[d.unidadeSolicitante] = (contagem[d.unidadeSolicitante] || 0) + 1;
        }
    });
    
    const dados = Object.entries(contagem)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15);
    
    const ctx = document.getElementById('chartUnidades');
    
    if (chartUnidades) {
        chartUnidades.destroy();
    }
    
    chartUnidades = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: dados.map(d => d[0]),
            datasets: [{
                label: 'Pendências',
                data: dados.map(d => d[1]),
                backgroundColor: '#4a90e2',
                borderRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    display: false
                },
                datalabels: {
                    display: true,
                    color: '#fff',
                    font: {
                        weight: 'bold',
                        size: 14
                    },
                    anchor: 'center',
                    align: 'center'
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        stepSize: 1
                    }
                },
                x: {
                    ticks: {
                        autoSkip: false,
                        maxRotation: 45,
                        minRotation: 45
                    }
                }
            }
        },
        plugins: [{
            afterDatasetsDraw: function(chart) {
                const ctx = chart.ctx;
                chart.data.datasets.forEach((dataset, i) => {
                    const meta = chart.getDatasetMeta(i);
                    meta.data.forEach((bar, index) => {
                        const data = dataset.data[index];
                        ctx.fillStyle = '#fff';
                        ctx.font = 'bold 14px Arial';
                        ctx.textAlign = 'center';
                        ctx.textBaseline = 'middle';
                        ctx.fillText(data, bar.x, bar.y);
                    });
                });
            }
        }]
    });
}

// Gráfico de Especialidades
function renderizarGraficoEspecialidades() {
    const contagem = {};
    dadosFiltrados.forEach(d => {
        if (d.cboEspecialidade) {
            contagem[d.cboEspecialidade] = (contagem[d.cboEspecialidade] || 0) + 1;
        }
    });
    
    const dados = Object.entries(contagem)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15);
    
    const ctx = document.getElementById('chartEspecialidades');
    
    if (chartEspecialidades) {
        chartEspecialidades.destroy();
    }
    
    chartEspecialidades = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: dados.map(d => d[0]),
            datasets: [{
                label: 'Pendências',
                data: dados.map(d => d[1]),
                backgroundColor: '#28a745',
                borderRadius: 8
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        stepSize: 1
                    }
                },
                x: {
                    ticks: {
                        autoSkip: false,
                        maxRotation: 45,
                        minRotation: 45
                    }
                }
            }
        },
        plugins: [{
            afterDatasetsDraw: function(chart) {
                const ctx = chart.ctx;
                chart.data.datasets.forEach((dataset, i) => {
                    const meta = chart.getDatasetMeta(i);
                    meta.data.forEach((bar, index) => {
                        const data = dataset.data[index];
                        ctx.fillStyle = '#fff';
                        ctx.font = 'bold 14px Arial';
                        ctx.textAlign = 'center';
                        ctx.textBaseline = 'middle';
                        ctx.fillText(data, bar.x, bar.y);
                    });
                });
            }
        }]
    });
}

// Gráficos de vencimento
function renderizarGraficosVencimento() {
    const hoje = new Date();
    
    // Vencendo em 15 dias
    const dados15 = dadosFiltrados.filter(d => {
        if (!d.dataInicioPendencia) return false;
        const dataInicio = parseData(d.dataInicioPendencia);
        if (!dataInicio) return false;
        const diffDias = Math.floor((hoje - dataInicio) / (1000 * 60 * 60 * 24));
        return diffDias >= 15;
    });
    
    // Vencendo em 30 dias
    const dados30 = dadosFiltrados.filter(d => {
        if (!d.dataInicioPendencia) return false;
        const dataInicio = parseData(d.dataInicioPendencia);
        if (!dataInicio) return false;
        const diffDias = Math.floor((hoje - dataInicio) / (1000 * 60 * 60 * 24));
        return diffDias >= 30;
    });
    
    // Gráfico 15 dias
    const ctx15 = document.getElementById('chartVenc15');
    if (chartVenc15) chartVenc15.destroy();
    
    chartVenc15 = new Chart(ctx15, {
        type: 'doughnut',
        data: {
            labels: ['Vencendo em 15 dias', 'Outros'],
            datasets: [{
                data: [dados15.length, dadosFiltrados.length - dados15.length],
                backgroundColor: ['#ffc107', '#e0e0e0']
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    position: 'bottom'
                }
            }
        }
    });
    
    // Gráfico 30 dias
    const ctx30 = document.getElementById('chartVenc30');
    if (chartVenc30) chartVenc30.destroy();
    
    chartVenc30 = new Chart(ctx30, {
        type: 'doughnut',
        data: {
            labels: ['Vencendo em 30 dias', 'Outros'],
            datasets: [{
                data: [dados30.length, dadosFiltrados.length - dados30.length],
                backgroundColor: ['#dc3545', '#e0e0e0']
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    position: 'bottom'
                }
            }
        }
    });
}

// Renderizar tabela
function renderizarTabela() {
    const tbody = document.getElementById('tabelaBody');
    const inicio = (currentPage - 1) * pageSize;
    const fim = inicio + pageSize;
    const dadosPagina = dadosFiltrados.slice(inicio, fim);
    
    if (dadosPagina.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="loading-cell">Nenhum registro encontrado</td></tr>';
        document.getElementById('paginationInfo').textContent = 'Mostrando 0 de 0 registros';
        document.getElementById('paginationButtons').innerHTML = '';
        return;
    }
    
    tbody.innerHTML = dadosPagina.map(d => `
        <tr>
            <td>${d.numeroSolicitacao || '-'}</td>
            <td>${d.dataSolicitacao || '-'}</td>
            <td>${d.prontuario || '-'}</td>
            <td>${d.telefone || '-'}</td>
            <td>${d.unidadeSolicitante || '-'}</td>
            <td>${d.cboEspecialidade || '-'}</td>
            <td>${d.dataInicioPendencia || '-'}</td>
            <td>${d.status || '-'}</td>
        </tr>
    `).join('');
    
    // Atualizar paginação
    const totalPaginas = Math.ceil(dadosFiltrados.length / pageSize);
    document.getElementById('paginationInfo').textContent = 
        `Mostrando ${inicio + 1} até ${Math.min(fim, dadosFiltrados.length)} de ${dadosFiltrados.length} registros`;
    
    renderizarBotoesPaginacao(totalPaginas);
}

// Renderizar botões de paginação
function renderizarBotoesPaginacao(totalPaginas) {
    const container = document.getElementById('paginationButtons');
    let html = '';
    
    // Botão anterior
    html += `<button ${currentPage === 1 ? 'disabled' : ''} onclick="mudarPagina(${currentPage - 1})">Anterior</button>`;
    
    // Botões de páginas
    const maxBotoes = 5;
    let inicioPag = Math.max(1, currentPage - Math.floor(maxBotoes / 2));
    let fimPag = Math.min(totalPaginas, inicioPag + maxBotoes - 1);
    
    if (fimPag - inicioPag < maxBotoes - 1) {
        inicioPag = Math.max(1, fimPag - maxBotoes + 1);
    }
    
    if (inicioPag > 1) {
        html += `<button onclick="mudarPagina(1)">1</button>`;
        if (inicioPag > 2) html += '<button disabled>...</button>';
    }
    
    for (let i = inicioPag; i <= fimPag; i++) {
        html += `<button class="${i === currentPage ? 'active' : ''}" onclick="mudarPagina(${i})">${i}</button>`;
    }
    
    if (fimPag < totalPaginas) {
        if (fimPag < totalPaginas - 1) html += '<button disabled>...</button>';
        html += `<button onclick="mudarPagina(${totalPaginas})">${totalPaginas}</button>`;
    }
    
    // Botão próximo
    html += `<button ${currentPage === totalPaginas ? 'disabled' : ''} onclick="mudarPagina(${currentPage + 1})">Próximo</button>`;
    
    container.innerHTML = html;
}

// Mudar página
function mudarPagina(pagina) {
    currentPage = pagina;
    renderizarTabela();
}

// Exportar para Excel
function exportarParaExcel() {
    if (dadosFiltrados.length === 0) {
        mostrarNotificacao('Não há dados para exportar', 'warning');
        return;
    }
    
    const dadosExportar = dadosFiltrados.map(d => ({
        'Nº Solicitação': d.numeroSolicitacao,
        'Data Solicitação': d.dataSolicitacao,
        'Nº Prontuário': d.prontuario,
        'Telefone': d.telefone,
        'Unidade Solicitante': d.unidadeSolicitante,
        'CBO Especialidade': d.cboEspecialidade,
        'Data Início Pendência': d.dataInicioPendencia,
        'Status': d.status
    }));
    
    const ws = XLSX.utils.json_to_sheet(dadosExportar);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pendências');
    
    const dataAtual = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Pendencias_Eldorado_${dataAtual}.xlsx`);
    
    mostrarNotificacao('Arquivo Excel exportado com sucesso!', 'success');
}

// Mostrar carregamento
function mostrarCarregamento(mostrar) {
    const tbody = document.getElementById('tabelaBody');
    if (mostrar) {
        tbody.innerHTML = '<tr><td colspan="8" class="loading-cell">Carregando dados...</td></tr>';
    }
}

// Mostrar notificação
function mostrarNotificacao(mensagem, tipo = 'info') {
    // Criar elemento de notificação
    const notif = document.createElement('div');
    notif.style.cssText = `
        position: fixed;
        top: 20px;
        right: 20px;
        padding: 1rem 1.5rem;
        background: ${tipo === 'success' ? '#28a745' : tipo === 'error' ? '#dc3545' : '#4a90e2'};
        color: white;
        border-radius: 8px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        z-index: 10000;
        animation: slideIn 0.3s ease;
        font-weight: 600;
    `;
    notif.textContent = mensagem;
    
    document.body.appendChild(notif);
    
    setTimeout(() => {
        notif.style.animation = 'slideOut 0.3s ease';
        setTimeout(() => notif.remove(), 300);
    }, 3000);
}

// Adicionar animações CSS
const style = document.createElement('style');
style.textContent = `
    @keyframes slideIn {
        from {
            transform: translateX(100%);
            opacity: 0;
        }
        to {
            transform: translateX(0);
            opacity: 1;
        }
    }
    
    @keyframes slideOut {
        from {
            transform: translateX(0);
            opacity: 1;
        }
        to {
            transform: translateX(100%);
            opacity: 0;
        }
    }
`;
document.head.appendChild(style);
