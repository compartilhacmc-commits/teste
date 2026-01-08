// Configuração do Google Sheets
const SHEET_ID = '1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4';
const SHEET_GID = '278071504';
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/export?format=csv&gid=${SHEET_GID}`;

// Variáveis globais
let allData = [];
let filteredData = [];
let chartUnidades = null;
let chartEspecialidades = null;
let chartVencimento15 = null;
let chartVencimento30 = null;

// Carregar dados ao iniciar
document.addEventListener('DOMContentLoaded', () => {
    loadData();
});

// Função para carregar dados do Google Sheets
async function loadData() {
    try {
        const response = await fetch(SHEET_URL);
        const csvText = await response.text();
        
        // Parse CSV
        const lines = csvText.split('\n');
        const headers = lines[0].split(',');
        
        allData = [];
        
        // Processar linhas (começar em 1 para pular cabeçalho)
        for (let i = 1; i < lines.length; i++) {
            const line = lines[i];
            if (!line.trim()) continue;
            
            const values = parseCSVLine(line);
            
            // Verificar se tem dados válidos
            if (values[0] && values[0].trim()) {
                const row = {
                    numeroSolicitacao: values[0] || '',
                    dataSolicitacao: values[1] || '',
                    prontuario: values[2] || '',
                    usuario: values[3] || '',
                    cns: values[4] || '',
                    nascimento: values[5] || '',
                    telefone: values[6] || '',
                    urgencia: values[7] || '',
                    tipoServico: values[8] || '',
                    unidadeSolicitante: values[9] || '',
                    prestador: values[10] || '',
                    cboEspecialidade: values[11] || '',
                    dataInicioPendencia: values[12] || '',
                    prazo15: values[13] || '',
                    dataFinalPrazo15: values[14] || '',
                    dataEnvioEmail15: values[15] || '',
                    prazo30: values[16] || '',
                    dataFinalPrazo30: values[17] || '',
                    dataEnvioEmail30: values[18] || '',
                    status: values[19] || '',
                    motivoPendencia: values[20] || '',
                    respostaPendencia: values[21] || '',
                    observacao: values[22] || ''
                };
                allData.push(row);
            }
        }
        
        filteredData = [...allData];
        
        // Inicializar interface
        populateFilters();
        updateDashboard();
        
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
        alert('Erro ao carregar dados da planilha. Verifique a conexão e as permissões.');
    }
}

// Função para fazer parse de linha CSV (considerando vírgulas dentro de aspas)
function parseCSVLine(line) {
    const result = [];
    let current = '';
    let inQuotes = false;
    
    for (let i = 0; i < line.length; i++) {
        const char = line[i];
        
        if (char === '"') {
            inQuotes = !inQuotes;
        } else if (char === ',' && !inQuotes) {
            result.push(current.trim());
            current = '';
        } else {
            current += char;
        }
    }
    
    result.push(current.trim());
    return result;
}

// Popular filtros com dados únicos
function populateFilters() {
    // Unidades
    const unidades = [...new Set(allData.map(d => d.unidadeSolicitante).filter(u => u))].sort();
    const selectUnidade = document.getElementById('filter-unidade');
    selectUnidade.innerHTML = '<option value="">Todas</option>';
    unidades.forEach(u => {
        const option = document.createElement('option');
        option.value = u;
        option.textContent = u;
        selectUnidade.appendChild(option);
    });
    
    // Especialidades
    const especialidades = [...new Set(allData.map(d => d.cboEspecialidade).filter(e => e))].sort();
    const selectEspecialidade = document.getElementById('filter-especialidade');
    selectEspecialidade.innerHTML = '<option value="">Todas</option>';
    especialidades.forEach(e => {
        const option = document.createElement('option');
        option.value = e;
        option.textContent = e;
        selectEspecialidade.appendChild(option);
    });
    
    // Status
    const statusList = [...new Set(allData.map(d => d.status).filter(s => s))].sort();
    const selectStatus = document.getElementById('filter-status');
    selectStatus.innerHTML = '<option value="">Todos</option>';
    statusList.forEach(s => {
        const option = document.createElement('option');
        option.value = s;
        option.textContent = s;
        selectStatus.appendChild(option);
    });
}

// Aplicar filtros
function applyFilters() {
    const selectedUnidades = Array.from(document.getElementById('filter-unidade').selectedOptions).map(o => o.value).filter(v => v);
    const selectedEspecialidades = Array.from(document.getElementById('filter-especialidade').selectedOptions).map(o => o.value).filter(v => v);
    const selectedStatus = Array.from(document.getElementById('filter-status').selectedOptions).map(o => o.value).filter(v => v);
    
    filteredData = allData.filter(row => {
        const unidadeMatch = selectedUnidades.length === 0 || selectedUnidades.includes(row.unidadeSolicitante);
        const especialidadeMatch = selectedEspecialidades.length === 0 || selectedEspecialidades.includes(row.cboEspecialidade);
        const statusMatch = selectedStatus.length === 0 || selectedStatus.includes(row.status);
        
        return unidadeMatch && especialidadeMatch && statusMatch;
    });
    
    updateDashboard();
}

// Limpar filtros
function clearFilters() {
    document.getElementById('filter-unidade').selectedIndex = -1;
    document.getElementById('filter-especialidade').selectedIndex = -1;
    document.getElementById('filter-status').selectedIndex = -1;
    
    filteredData = [...allData];
    updateDashboard();
}

// Atualizar dashboard completo
function updateDashboard() {
    updateCards();
    updateCharts();
    updateTable();
}

// Atualizar cards
function updateCards() {
    const total = filteredData.length;
    document.getElementById('total-pendencias').textContent = total;
    
    // Calcular vencimentos
    const hoje = new Date();
    const vencendo15 = filteredData.filter(row => {
        if (!row.dataInicioPendencia) return false;
        const dataInicio = parseDate(row.dataInicioPendencia);
        const diff = Math.ceil((hoje - dataInicio) / (1000 * 60 * 60 * 24));
        return diff >= 0 && diff <= 15;
    }).length;
    
    const vencendo30 = filteredData.filter(row => {
        if (!row.dataInicioPendencia) return false;
        const dataInicio = parseDate(row.dataInicioPendencia);
        const diff = Math.ceil((hoje - dataInicio) / (1000 * 60 * 60 * 24));
        return diff >= 0 && diff <= 30;
    }).length;
    
    document.getElementById('vencendo-15').textContent = vencendo15;
    document.getElementById('vencendo-30').textContent = vencendo30;
    
    // Porcentagem
    const porcentagem = allData.length > 0 ? ((total / allData.length) * 100).toFixed(1) : 0;
    document.getElementById('porcentagem-filtro').textContent = `${porcentagem}%`;
}

// Atualizar gráficos
function updateCharts() {
    updateChartUnidades();
    updateChartEspecialidades();
    updateChartVencimentos();
}

// Gráfico de Unidades
function updateChartUnidades() {
    const unidadeCounts = {};
    filteredData.forEach(row => {
        const unidade = row.unidadeSolicitante || 'Não informado';
        unidadeCounts[unidade] = (unidadeCounts[unidade] || 0) + 1;
    });
    
    const sortedUnidades = Object.entries(unidadeCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15);
    
    const labels = sortedUnidades.map(u => u[0]);
    const data = sortedUnidades.map(u => u[1]);
    
    const ctx = document.getElementById('chartUnidades').getContext('2d');
    
    if (chartUnidades) {
        chartUnidades.destroy();
    }
    
    chartUnidades = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Pendências',
                data: data,
                backgroundColor: 'rgba(42, 82, 152, 0.8)',
                borderColor: 'rgba(42, 82, 152, 1)',
                borderWidth: 2,
                borderRadius: 8
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                datalabels: {
                    anchor: 'center',
                    align: 'center',
                    color: 'white',
                    font: {
                        weight: 'bold',
                        size: 14
                    }
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    grid: {
                        display: true,
                        color: 'rgba(0, 0, 0, 0.05)'
                    }
                },
                y: {
                    grid: {
                        display: false
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
                        ctx.fillStyle = 'white';
                        ctx.font = 'bold 14px Arial';
                        ctx.textAlign = 'center';
                        ctx.textBaseline = 'middle';
                        ctx.fillText(data, bar.x / 2, bar.y);
                    });
                });
            }
        }]
    });
}

// Gráfico de Especialidades
function updateChartEspecialidades() {
    const especialidadeCounts = {};
    filteredData.forEach(row => {
        const especialidade = row.cboEspecialidade || 'Não informado';
        especialidadeCounts[especialidade] = (especialidadeCounts[especialidade] || 0) + 1;
    });
    
    const sortedEspecialidades = Object.entries(especialidadeCounts)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 15);
    
    const labels = sortedEspecialidades.map(e => e[0]);
    const data = sortedEspecialidades.map(e => e[1]);
    
    const ctx = document.getElementById('chartEspecialidades').getContext('2d');
    
    if (chartEspecialidades) {
        chartEspecialidades.destroy();
    }
    
    chartEspecialidades = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Pendências',
                data: data,
                backgroundColor: 'rgba(102, 187, 106, 0.8)',
                borderColor: 'rgba(102, 187, 106, 1)',
                borderWidth: 2,
                borderRadius: 8
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    grid: {
                        display: true,
                        color: 'rgba(0, 0, 0, 0.05)'
                    }
                },
                y: {
                    grid: {
                        display: false
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
                        ctx.fillStyle = 'white';
                        ctx.font = 'bold 14px Arial';
                        ctx.textAlign = 'center';
                        ctx.textBaseline = 'middle';
                        ctx.fillText(data, bar.x / 2, bar.y);
                    });
                });
            }
        }]
    });
}

// Gráficos de Vencimento
function updateChartVencimentos() {
    const hoje = new Date();
    
    // Vencimento 15 dias
    const vencimento15Counts = {
        'Vencidas': 0,
        '1-5 dias': 0,
        '6-10 dias': 0,
        '11-15 dias': 0
    };
    
    filteredData.forEach(row => {
        if (!row.dataInicioPendencia) return;
        const dataInicio = parseDate(row.dataInicioPendencia);
        const diff = Math.ceil((hoje - dataInicio) / (1000 * 60 * 60 * 24));
        
        if (diff < 0) vencimento15Counts['Vencidas']++;
        else if (diff <= 5) vencimento15Counts['1-5 dias']++;
        else if (diff <= 10) vencimento15Counts['6-10 dias']++;
        else if (diff <= 15) vencimento15Counts['11-15 dias']++;
    });
    
    const ctx15 = document.getElementById('chartVencimento15').getContext('2d');
    
    if (chartVencimento15) {
        chartVencimento15.destroy();
    }
    
    chartVencimento15 = new Chart(ctx15, {
        type: 'doughnut',
        data: {
            labels: Object.keys(vencimento15Counts),
            datasets: [{
                data: Object.values(vencimento15Counts),
                backgroundColor: [
                    'rgba(239, 83, 80, 0.8)',
                    'rgba(255, 167, 38, 0.8)',
                    'rgba(255, 193, 7, 0.8)',
                    'rgba(102, 187, 106, 0.8)'
                ],
                borderWidth: 2,
                borderColor: 'white'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom'
                }
            }
        }
    });
    
    // Vencimento 30 dias
    const vencimento30Counts = {
        'Vencidas': 0,
        '1-10 dias': 0,
        '11-20 dias': 0,
        '21-30 dias': 0
    };
    
    filteredData.forEach(row => {
        if (!row.dataInicioPendencia) return;
        const dataInicio = parseDate(row.dataInicioPendencia);
        const diff = Math.ceil((hoje - dataInicio) / (1000 * 60 * 60 * 24));
        
        if (diff < 0) vencimento30Counts['Vencidas']++;
        else if (diff <= 10) vencimento30Counts['1-10 dias']++;
        else if (diff <= 20) vencimento30Counts['11-20 dias']++;
        else if (diff <= 30) vencimento30Counts['21-30 dias']++;
    });
    
    const ctx30 = document.getElementById('chartVencimento30').getContext('2d');
    
    if (chartVencimento30) {
        chartVencimento30.destroy();
    }
    
    chartVencimento30 = new Chart(ctx30, {
        type: 'doughnut',
        data: {
            labels: Object.keys(vencimento30Counts),
            datasets: [{
                data: Object.values(vencimento30Counts),
                backgroundColor: [
                    'rgba(239, 83, 80, 0.8)',
                    'rgba(255, 167, 38, 0.8)',
                    'rgba(255, 193, 7, 0.8)',
                    'rgba(102, 187, 106, 0.8)'
                ],
                borderWidth: 2,
                borderColor: 'white'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'bottom'
                }
            }
        }
    });
}

// Atualizar tabela
function updateTable() {
    const tbody = document.getElementById('tableBody');
    tbody.innerHTML = '';
    
    filteredData.forEach(row => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td>${row.numeroSolicitacao}</td>
            <td>${row.dataSolicitacao}</td>
            <td>${row.prontuario}</td>
            <td>${row.telefone}</td>
            <td>${row.unidadeSolicitante}</td>
            <td>${row.cboEspecialidade}</td>
            <td>${row.dataInicioPendencia}</td>
            <td>${row.status}</td>
        `;
        tbody.appendChild(tr);
    });
}

// Parse data (formato DD/MM/YYYY)
function parseDate(dateStr) {
    const parts = dateStr.split('/');
    if (parts.length === 3) {
        return new Date(parts[2], parts[1] - 1, parts[0]);
    }
    return new Date();
}

// Recarregar página
function reloadPage() {
    location.reload();
}

// Download Excel
function downloadExcel() {
    const ws_data = [
        ['Nº Solicitação', 'Data Solicitação', 'Nº Prontuário', 'Telefone', 'Unidade Solicitante', 'CBO Especialidade', 'Data Início Pendência', 'Status']
    ];
    
    filteredData.forEach(row => {
        ws_data.push([
            row.numeroSolicitacao,
            row.dataSolicitacao,
            row.prontuario,
            row.telefone,
            row.unidadeSolicitante,
            row.cboEspecialidade,
            row.dataInicioPendencia,
            row.status
        ]);
    });
    
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(ws_data);
    
    XLSX.utils.book_append_sheet(wb, ws, 'Pendências');
    XLSX.writeFile(wb, 'pendencias_filtradas.xlsx');
}
