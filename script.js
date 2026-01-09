// ===================================
// CONFIGURAÇÃO DA PLANILHA
// ===================================
const SHEET_ID = '1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4';
const SHEET_NAME = 'PENDÊNCIAS ELDORADO';
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(SHEET_NAME)}`;

// ===================================
// VARIÁVEIS GLOBAIS
// ===================================
let allData = [];
let filteredData = [];
let chartUnidades = null;
let chartEspecialidades = null;
let chartStatus = null;
let chartPizzaStatus = null;

// ===================================
// INICIALIZAÇÃO
// ===================================
document.addEventListener('DOMContentLoaded', function() {
    console.log('Iniciando carregamento de dados...');
    loadData();
});

// ===================================
// CARREGAR DADOS DA PLANILHA
// ===================================
async function loadData() {
    showLoading(true);
    try {
        console.log('Fazendo requisição para:', SHEET_URL);
        
        const response = await fetch(SHEET_URL);
        
        if (!response.ok) {
            throw new Error(`Erro HTTP: ${response.status}`);
        }
        
        const csvText = await response.text();
        console.log('Dados CSV recebidos, primeiros 500 caracteres:', csvText.substring(0, 500));
        
        const rows = parseCSV(csvText);
        
        if (rows.length < 2) {
            throw new Error('Planilha vazia ou sem dados');
        }
        
        const headers = rows[0];
        console.log('Cabeçalhos encontrados:', headers);
        
        allData = rows.slice(1)
            .filter(row => row.length > 1 && row[0])
            .map(row => {
                const obj = {};
                headers.forEach((header, index) => {
                    obj[header.trim()] = (row[index] || '').trim();
                });
                return obj;
            });
        
        console.log(`Total de registros carregados: ${allData.length}`);
        console.log('Primeiro registro:', allData[0]);
        
        filteredData = [...allData];
        
        populateFilters();
        updateDashboard();
        
        console.log('Dados carregados com sucesso!');
        
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
        alert(`Erro ao carregar dados da planilha: ${error.message}\n\nVerifique:\n1. A planilha está pública?\n2. O nome da aba está correto: "${SHEET_NAME}"?\n3. Há dados na planilha?`);
    } finally {
        showLoading(false);
    }
}

// ===================================
// PARSE CSV (COM SUPORTE A ASPAS)
// ===================================
function parseCSV(text) {
    const rows = [];
    let currentRow = [];
    let currentCell = '';
    let insideQuotes = false;
    
    for (let i = 0; i < text.length; i++) {
        const char = text[i];
        const nextChar = text[i + 1];
        
        if (char === '"') {
            if (insideQuotes && nextChar === '"') {
                currentCell += '"';
                i++;
            } else {
                insideQuotes = !insideQuotes;
            }
        } else if (char === ',' && !insideQuotes) {
            currentRow.push(currentCell.trim());
            currentCell = '';
        } else if ((char === '\n' || char === '\r') && !insideQuotes) {
            if (currentCell || currentRow.length > 0) {
                currentRow.push(currentCell.trim());
                rows.push(currentRow);
                currentRow = [];
                currentCell = '';
            }
            if (char === '\r' && nextChar === '\n') {
                i++;
            }
        } else {
            currentCell += char;
        }
    }
    
    if (currentCell || currentRow.length > 0) {
        currentRow.push(currentCell.trim());
        rows.push(currentRow);
    }
    
    return rows;
}

// ===================================
// MOSTRAR/OCULTAR LOADING
// ===================================
function showLoading(show) {
    const overlay = document.getElementById('loadingOverlay');
    if (show) {
        overlay.classList.add('active');
    } else {
        overlay.classList.remove('active');
    }
}

// ===================================
// POPULAR FILTROS
// ===================================
function populateFilters() {
    const statusList = [...new Set(allData.map(item => item['Status']))].filter(Boolean).sort();
    const selectStatus = document.getElementById('filterStatus');
    selectStatus.innerHTML = '<option value="">Todos</option>';
    statusList.forEach(status => {
        const option = document.createElement('option');
        option.value = status;
        option.textContent = status;
        selectStatus.appendChild(option);
    });
    
    const unidades = [...new Set(allData.map(item => item['Unidade Solicitante']))].filter(Boolean).sort();
    const selectUnidade = document.getElementById('filterUnidade');
    selectUnidade.innerHTML = '<option value="">Todas</option>';
    unidades.forEach(unidade => {
        const option = document.createElement('option');
        option.value = unidade;
        option.textContent = unidade;
        selectUnidade.appendChild(option);
    });
    
    const especialidades = [...new Set(allData.map(item => item['Cbo Especialidade']))].filter(Boolean).sort();
    const selectEspecialidade = document.getElementById('filterEspecialidade');
    selectEspecialidade.innerHTML = '<option value="">Todas</option>';
    especialidades.forEach(especialidade => {
        const option = document.createElement('option');
        option.value = especialidade;
        option.textContent = especialidade;
        selectEspecialidade.appendChild(option);
    });
    
    const prestadores = [...new Set(allData.map(item => item['Prestador']))].filter(Boolean).sort();
    const selectPrestador = document.getElementById('filterPrestador');
    selectPrestador.innerHTML = '<option value="">Todos</option>';
    prestadores.forEach(prestador => {
        const option = document.createElement('option');
        option.value = prestador;
        option.textContent = prestador;
        selectPrestador.appendChild(option);
    });
}

// ===================================
// APLICAR FILTROS
// ===================================
function applyFilters() {
    const status = document.getElementById('filterStatus').value;
    const unidade = document.getElementById('filterUnidade').value;
    const especialidade = document.getElementById('filterEspecialidade').value;
    const prestador = document.getElementById('filterPrestador').value;
    
    filteredData = allData.filter(item => {
        return (!status || item['Status'] === status) &&
               (!unidade || item['Unidade Solicitante'] === unidade) &&
               (!especialidade || item['Cbo Especialidade'] === especialidade) &&
               (!prestador || item['Prestador'] === prestador);
    });
    
    updateDashboard();
}

// ===================================
// LIMPAR FILTROS
// ===================================
function clearFilters() {
    document.getElementById('filterStatus').value = '';
    document.getElementById('filterUnidade').value = '';
    document.getElementById('filterEspecialidade').value = '';
    document.getElementById('filterPrestador').value = '';
    document.getElementById('searchInput').value = '';
    
    filteredData = [...allData];
    updateDashboard();
}

// ===================================
// PESQUISAR NA TABELA
// ===================================
function searchTable() {
    const searchValue = document.getElementById('searchInput').value.toLowerCase();
    const tbody = document.getElementById('tableBody');
    const rows = tbody.getElementsByTagName('tr');
    
    let visibleCount = 0;
    
    for (let i = 0; i < rows.length; i++) {
        const row = rows[i];
        const cells = row.getElementsByTagName('td');
        let found = false;
        
        for (let j = 0; j < cells.length; j++) {
            const cellText = cells[j].textContent.toLowerCase();
            if (cellText.includes(searchValue)) {
                found = true;
                break;
            }
        }
        
        if (found) {
            row.style.display = '';
            visibleCount++;
        } else {
            row.style.display = 'none';
        }
    }
    
    const footer = document.getElementById('tableFooter');
    footer.textContent = `Mostrando ${visibleCount} de ${filteredData.length} registros`;
}

// ===================================
// ATUALIZAR DASHBOARD
// ===================================
function updateDashboard() {
    updateCards();
    updateCharts();
    updateTable();
}

// ===================================
// ATUALIZAR CARDS
// ===================================
function updateCards() {
    const total = allData.length;
    const filtrado = filteredData.length;
    
    const hoje = new Date();
    let pendencias15 = 0;
    let pendencias30 = 0;
    
    filteredData.forEach(item => {
        const dataInicio = parseDate(item['Data Início da Pendência']);
        if (dataInicio) {
            const diasDecorridos = Math.floor((hoje - dataInicio) / (1000 * 60 * 60 * 24));
            
            if (diasDecorridos >= 15 && diasDecorridos < 30) {
                pendencias15++;
            }
            if (diasDecorridos >= 30) {
                pendencias30++;
            }
        }
    });
    
    document.getElementById('totalPendencias').textContent = total;
    document.getElementById('pendencias15').textContent = pendencias15;
    document.getElementById('pendencias30').textContent = pendencias30;
    
    const percentFiltrados = total > 0 ? ((filtrado / total) * 100).toFixed(1) : '100.0';
    document.getElementById('percentFiltrados').textContent = percentFiltrados + '%';
}

// ===================================
// ATUALIZAR GRÁFICOS
// ===================================
function updateCharts() {
    // Gráfico de Unidades (HORIZONTAL VERDE)
    const unidadesCount = {};
    filteredData.forEach(item => {
        const unidade = item['Unidade Solicitante'] || 'Não informado';
        unidadesCount[unidade] = (unidadesCount[unidade] || 0) + 1;
    });
    
    const unidadesLabels = Object.keys(unidadesCount)
        .sort((a, b) => unidadesCount[b] - unidadesCount[a])
        .slice(0, 10);
    const unidadesValues = unidadesLabels.map(label => unidadesCount[label]);
    
    createHorizontalBarChart('chartUnidades', unidadesLabels, unidadesValues, '#48bb78');
    
    // Gráfico de Especialidades (HORIZONTAL VERMELHO)
    const especialidadesCount = {};
    filteredData.forEach(item => {
        const especialidade = item['Cbo Especialidade'] || 'Não informado';
        especialidadesCount[especialidade] = (especialidadesCount[especialidade] || 0) + 1;
    });
    
    const especialidadesLabels = Object.keys(especialidadesCount)
        .sort((a, b) => especialidadesCount[b] - especialidadesCount[a])
        .slice(0, 10);
    const especialidadesValues = especialidadesLabels.map(label => especialidadesCount[label]);
    
    createHorizontalBarChart('chartEspecialidades', especialidadesLabels, especialidadesValues, '#ef4444');
    
    // Gráfico de Status (HORIZONTAL LARANJA)
    const statusCount = {};
    filteredData.forEach(item => {
        const status = item['Status'] || 'Não informado';
        statusCount[status] = (statusCount[status] || 0) + 1;
    });
    
    const statusLabels = Object.keys(statusCount)
        .sort((a, b) => statusCount[b] - statusCount[a]);
    const statusValues = statusLabels.map(label => statusCount[label]);
    
    createHorizontalBarChart('chartStatus', statusLabels, statusValues, '#f97316');
    
    // NOVO: Gráfico de Pizza por Status
    createPieChart('chartPizzaStatus', statusLabels, statusValues);
}

// ===================================
// CRIAR GRÁFICO DE BARRAS HORIZONTAIS - CORRIGIDO
// ===================================
function createHorizontalBarChart(canvasId, labels, data, color) {
    const ctx = document.getElementById(canvasId);
    
    if (canvasId === 'chartUnidades' && chartUnidades) {
        chartUnidades.destroy();
    }
    if (canvasId === 'chartEspecialidades' && chartEspecialidades) {
        chartEspecialidades.destroy();
    }
    if (canvasId === 'chartStatus' && chartStatus) {
        chartStatus.destroy();
    }
    
    const chart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Quantidade',
                data: data,
                backgroundColor: color,
                borderWidth: 0,
                borderRadius: 4
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
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleFont: {
                        size: 14,
                        weight: 'bold'
                    },
                    bodyFont: {
                        size: 13
                    },
                    padding: 12,
                    cornerRadius: 8
                }
            },
            scales: {
                x: {
                    display: false,
                    grid: {
                        display: false
                    }
                },
                y: {
                    ticks: {
                        font: {
                            size: 12,
                            weight: '500'
                        },
                        color: '#4a5568',
                        padding: 8
                    },
                    grid: {
                        display: false
                    }
                }
            },
            layout: {
                padding: {
                    right: 50
                }
            }
        },
        plugins: [{
            id: 'customLabels',
            afterDatasetsDraw: function(chart) {
                const ctx = chart.ctx;
                chart.data.datasets.forEach(function(dataset, i) {
                    const meta = chart.getDatasetMeta(i);
                    if (!meta.hidden) {
                        meta.data.forEach(function(element, index) {
                            // CORREÇÃO: LEGENDAS FORA DAS BARRAS, COR PRETA, NEGRITO
                            ctx.fillStyle = '#000000';
                            ctx.font = 'bold 14px Arial';
                            ctx.textAlign = 'left';
                            ctx.textBaseline = 'middle';
                            
                            const dataString = dataset.data[index].toString();
                            const xPos = element.x + 10; // FORA DA BARRA
                            const yPos = element.y;
                            
                            ctx.fillText(dataString, xPos, yPos);
                        });
                    }
                });
            }
        }]
    });
    
    if (canvasId === 'chartUnidades') {
        chartUnidades = chart;
    }
    if (canvasId === 'chartEspecialidades') {
        chartEspecialidades = chart;
    }
    if (canvasId === 'chartStatus') {
        chartStatus = chart;
    }
}

// ===================================
// CRIAR GRÁFICO DE PIZZA - CORRIGIDO
// ===================================
function createPieChart(canvasId, labels, data) {
    const ctx = document.getElementById(canvasId);
    
    if (chartPizzaStatus) {
        chartPizzaStatus.destroy();
    }
    
    // Paleta de cores vibrantes
    const colors = [
        '#3b82f6', '#ef4444', '#10b981', '#f59e0b', '#8b5cf6',
        '#ec4899', '#06b6d4', '#f97316', '#6366f1', '#84cc16'
    ];
    
    const total = data.reduce((sum, val) => sum + val, 0);
    const percentages = data.map(val => ((val / total) * 100).toFixed(1));
    
    chartPizzaStatus = new Chart(ctx, {
        type: 'pie',
        data: {
            labels: labels,
            datasets: [{
                data: data,
                backgroundColor: colors.slice(0, labels.length),
                borderWidth: 3,
                borderColor: '#ffffff'
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    position: 'right',
                    labels: {
                        font: {
                            size: 13,
                            weight: '600'
                        },
                        color: '#1f2937',
                        padding: 15,
                        usePointStyle: true,
                        pointStyle: 'circle',
                        generateLabels: function(chart) {
                            const datasets = chart.data.datasets;
                            return chart.data.labels.map((label, i) => {
                                const value = datasets[0].data[i];
                                const percentage = percentages[i];
                                return {
                                    text: `${label}: ${value} (${percentage}%)`,
                                    fillStyle: datasets[0].backgroundColor[i],
                                    hidden: false,
                                    index: i
                                };
                            });
                        }
                    }
                },
                tooltip: {
                    enabled: true,
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    titleFont: {
                        size: 14,
                        weight: 'bold'
                    },
                    bodyFont: {
                        size: 13
                    },
                    padding: 12,
                    cornerRadius: 8,
                    callbacks: {
                        label: function(context) {
                            const value = context.parsed;
                            const percentage = ((value / total) * 100).toFixed(1);
                            return `${context.label}: ${value} (${percentage}%)`;
                        }
                    }
                }
            }
        },
        plugins: [{
            id: 'customPieLabels',
            afterDatasetsDraw: function(chart) {
                const ctx = chart.ctx;
                chart.data.datasets.forEach(function(dataset, datasetIndex) {
                    const meta = chart.getDatasetMeta(datasetIndex);
                    if (!meta.hidden) {
                        meta.data.forEach(function(element, index) {
                            // CORREÇÃO: LEGENDAS DENTRO DO GRÁFICO, COR BRANCA
                            const data = dataset.data[index];
                            const percentage = percentages[index];
                            
                            // Só mostra se for maior que 5% para não poluir
                            if (parseFloat(percentage) > 5) {
                                ctx.fillStyle = '#ffffff';
                                ctx.font = 'bold 13px Arial';
                                ctx.textAlign = 'center';
                                ctx.textBaseline = 'middle';
                                
                                const position = element.tooltipPosition();
                                ctx.fillText(`${data}`, position.x, position.y - 8);
                                ctx.fillText(`${percentage}%`, position.x, position.y + 8);
                            }
                        });
                    }
                });
            }
        }]
    });
}

// ===================================
// ATUALIZAR TABELA - CORRIGIDA
// ===================================
function updateTable() {
    const tbody = document.getElementById('tableBody');
    const footer = document.getElementById('tableFooter');
    tbody.innerHTML = '';
    
    if (filteredData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="12" class="loading-message"><i class="fas fa-inbox"></i> Nenhum registro encontrado</td></tr>';
        footer.textContent = 'Mostrando 0 registros';
        return;
    }
    
    filteredData.forEach(item => {
        const row = document.createElement('tr');
        
        // CORREÇÃO: Busca múltiplas variações possíveis do nome da coluna
        const numeroSolicitacao = item['Numero da Solicitação'] || 
                                  item['Numero Solicitação'] || 
                                  item['Número da Solicitação'] || 
                                  item['Número Solicitação'] ||
                                  item['N° da Solicitação'] ||
                                  item['N° Solicitação'] || '-';
        
        row.innerHTML = `
            <td>${numeroSolicitacao}</td>
            <td>${formatDate(item['Data da Solicitação'])}</td>
            <td>${item['Nº Prontuário'] || item['N° Prontuário'] || '-'}</td>
            <td>${item['Telefone'] || '-'}</td>
            <td>${item['Unidade Solicitante'] || '-'}</td>
            <td>${item['Cbo Especialidade'] || '-'}</td>
            <td>${formatDate(item['Data Início da Pendência'])}</td>
            <td>${item['Status'] || '-'}</td>
            <td>${formatDate(item['Data Final do Prazo (Pendência com 15 dias)'])}</td>
            <td>${formatDate(item['Data do envio do Email (Prazo: Pendência com 15 dias)'])}</td>
            <td>${formatDate(item['Data Final do Prazo (Pendência com 30 dias)'])}</td>
            <td>${formatDate(item['Data do envio do Email (Prazo: Pendência com 30 dias)'])}</td>
        `;
        tbody.appendChild(row);
    });
    
    const total = allData.length;
    const showing = filteredData.length;
    footer.textContent = `Mostrando de 1 até ${showing} de ${total} registros`;
}

// ===================================
// FUNÇÕES AUXILIARES
// ===================================
function parseDate(dateString) {
    if (!dateString) return null;
    
    let match = dateString.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (match) {
        return new Date(match[3], match[2] - 1, match[1]);
    }
    
    match = dateString.match(/(\d{4})-(\d{2})-(\d{2})/);
    if (match) {
        return new Date(match[1], match[2] - 1, match[3]);
    }
    
    return null;
}

function formatDate(dateString) {
    if (!dateString) return '-';
    
    const date = parseDate(dateString);
    if (!date || isNaN(date.getTime())) return dateString;
    
    const day = String(date.getDate()).padStart(2, '0');
    const month = String(date.getMonth() + 1).padStart(2, '0');
    const year = date.getFullYear();
    
    return `${day}/${month}/${year}`;
}

// ===================================
// ATUALIZAR DADOS
// ===================================
function refreshData() {
    loadData();
}

// ===================================
// DOWNLOAD EXCEL - CORRIGIDO
// ===================================
function downloadExcel() {
    if (filteredData.length === 0) {
        alert('Não há dados para exportar.');
        return;
    }
    
    const exportData = filteredData.map(item => ({
        'Numero da Solicitação': item['Numero da Solicitação'] || item['Numero Solicitação'] || item['N° Solicitação'] || '',
        'Data Solicitação': item['Data da Solicitação'] || '',
        'Nº Prontuário': item['Nº Prontuário'] || item['N° Prontuário'] || '',
        'Telefone': item['Telefone'] || '',
        'Unidade Solicitante': item['Unidade Solicitante'] || '',
        'CBO Especialidade': item['Cbo Especialidade'] || '',
        'Data Início Pendência': item['Data Início da Pendência'] || '',
        'Status': item['Status'] || '',
        'Prestador': item['Prestador'] || '',
        'Data Final Prazo 15d': item['Data Final do Prazo (Pendência com 15 dias)'] || '',
        'Data Envio Email 15d': item['Data do envio do Email (Prazo: Pendência com 15 dias)'] || '',
        'Data Final Prazo 30d': item['Data Final do Prazo (Pendência com 30 dias)'] || '',
        'Data Envio Email 30d': item['Data do envio do Email (Prazo: Pendência com 30 dias)'] || ''
    }));
    
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pendências');
    
    ws['!cols'] = [
        { wch: 18 }, { wch: 15 }, { wch: 15 }, { wch: 15 },
        { wch: 30 }, { wch: 30 }, { wch: 18 }, { wch: 20 }, 
        { wch: 25 }, { wch: 18 }, { wch: 20 }, { wch: 18 }, { wch: 20 }
    ];
    
    const hoje = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Pendencias_Eldorado_${hoje}.xlsx`);
}
