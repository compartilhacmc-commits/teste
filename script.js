// ===================================
// CONFIGURAÇÃO DA PLANILHA
// ===================================
const SHEET_ID = '1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4';
const SHEET_NAME = 'PENDÊNCIAS ELDORADO';
// URL corrigida com encoding do nome da aba
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(SHEET_NAME)}`;

// ===================================
// VARIÁVEIS GLOBAIS
// ===================================
let allData = [];
let filteredData = [];
let chartUnidades = null;
let chartEspecialidades = null;

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
        
        // Parse CSV
        const rows = parseCSV(csvText);
        
        if (rows.length < 2) {
            throw new Error('Planilha vazia ou sem dados');
        }
        
        // Cabeçalhos (primeira linha)
        const headers = rows[0];
        console.log('Cabeçalhos encontrados:', headers);
        
        // Converter linhas em objetos
        allData = rows.slice(1)
            .filter(row => row.length > 1 && row[0]) // Filtrar linhas vazias
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
        
        // Inicializar interface
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
                // Aspas duplas dentro de aspas
                currentCell += '"';
                i++; // Pular próxima aspa
            } else {
                // Alternar estado de aspas
                insideQuotes = !insideQuotes;
            }
        } else if (char === ',' && !insideQuotes) {
            // Fim de célula
            currentRow.push(currentCell.trim());
            currentCell = '';
        } else if ((char === '\n' || char === '\r') && !insideQuotes) {
            // Fim de linha
            if (currentCell || currentRow.length > 0) {
                currentRow.push(currentCell.trim());
                rows.push(currentRow);
                currentRow = [];
                currentCell = '';
            }
            // Pular \r\n
            if (char === '\r' && nextChar === '\n') {
                i++;
            }
        } else {
            currentCell += char;
        }
    }
    
    // Adicionar última célula e linha
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
    // Status
    const statusList = [...new Set(allData.map(item => item['Status']))].filter(Boolean).sort();
    const selectStatus = document.getElementById('filterStatus');
    selectStatus.innerHTML = '<option value="">Todos</option>';
    statusList.forEach(status => {
        const option = document.createElement('option');
        option.value = status;
        option.textContent = status;
        selectStatus.appendChild(option);
    });
    
    // Unidades
    const unidades = [...new Set(allData.map(item => item['Unidade Solicitante']))].filter(Boolean).sort();
    const selectUnidade = document.getElementById('filterUnidade');
    selectUnidade.innerHTML = '<option value="">Todas</option>';
    unidades.forEach(unidade => {
        const option = document.createElement('option');
        option.value = unidade;
        option.textContent = unidade;
        selectUnidade.appendChild(option);
    });
    
    // Especialidades
    const especialidades = [...new Set(allData.map(item => item['Cbo Especialidade']))].filter(Boolean).sort();
    const selectEspecialidade = document.getElementById('filterEspecialidade');
    selectEspecialidade.innerHTML = '<option value="">Todas</option>';
    especialidades.forEach(especialidade => {
        const option = document.createElement('option');
        option.value = especialidade;
        option.textContent = especialidade;
        selectEspecialidade.appendChild(option);
    });
    
    // Prestador (NOVO FILTRO)
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
    
    filteredData = [...allData];
    updateDashboard();
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
    
    // Calcular pendências por prazo
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
    
    // Atualizar valores dos cards
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
    // Gráfico de Unidades
    const unidadesCount = {};
    filteredData.forEach(item => {
        const unidade = item['Unidade Solicitante'] || 'Não informado';
        unidadesCount[unidade] = (unidadesCount[unidade] || 0) + 1;
    });
    
    // Ordenar e pegar top 10
    const unidadesLabels = Object.keys(unidadesCount)
        .sort((a, b) => unidadesCount[b] - unidadesCount[a])
        .slice(0, 10);
    const unidadesValues = unidadesLabels.map(label => unidadesCount[label]);
    
    createBarChart('chartUnidades', unidadesLabels, unidadesValues, '#2563eb');
    
    // Gráfico de Especialidades
    const especialidadesCount = {};
    filteredData.forEach(item => {
        const especialidade = item['Cbo Especialidade'] || 'Não informado';
        especialidadesCount[especialidade] = (especialidadesCount[especialidade] || 0) + 1;
    });
    
    // Ordenar e pegar top 10
    const especialidadesLabels = Object.keys(especialidadesCount)
        .sort((a, b) => especialidadesCount[b] - especialidadesCount[a])
        .slice(0, 10);
    const especialidadesValues = especialidadesLabels.map(label => especialidadesCount[label]);
    
    createBarChart('chartEspecialidades', especialidadesLabels, especialidadesValues, '#10b981');
}

// ===================================
// CRIAR GRÁFICO DE BARRAS
// ===================================
function createBarChart(canvasId, labels, data, color) {
    const ctx = document.getElementById(canvasId);
    
    // Destruir gráfico anterior
    if (canvasId === 'chartUnidades' && chartUnidades) {
        chartUnidades.destroy();
    } else if (canvasId === 'chartEspecialidades' && chartEspecialidades) {
        chartEspecialidades.destroy();
    }
    
    // Criar novo gráfico
    const newChart = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Quantidade',
                data: data,
                backgroundColor: color,
                borderColor: color,
                borderWidth: 0,
                borderRadius: 8,
                barThickness: 'flex',
                maxBarThickness: 50
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            plugins: {
                legend: {
                    display: false
                },
                tooltip: {
                    backgroundColor: 'rgba(0, 0, 0, 0.8)',
                    padding: 12,
                    titleFont: { size: 14, weight: 'bold' },
                    bodyFont: { size: 13 },
                    cornerRadius: 8,
                    displayColors: false
                },
                datalabels: {
                    anchor: 'end',
                    align: 'end',
                    color: '#1f2937',
                    font: {
                        weight: 'bold',
                        size: 12
                    },
                    formatter: (value) => value
                }
            },
            scales: {
                y: {
                    beginAtZero: true,
                    ticks: {
                        font: { size: 12 },
                        color: '#6b7280',
                        precision: 0
                    },
                    grid: {
                        color: '#e5e7eb',
                        drawBorder: false
                    }
                },
                x: {
                    ticks: {
                        font: { size: 11 },
                        color: '#6b7280',
                        maxRotation: 45,
                        minRotation: 45
                    },
                    grid: {
                        display: false,
                        drawBorder: false
                    }
                }
            }
        },
        plugins: [{
            afterDatasetsDraw: function(chart) {
                const ctx = chart.ctx;
                chart.data.datasets.forEach(function(dataset, i) {
                    const meta = chart.getDatasetMeta(i);
                    if (!meta.hidden) {
                        meta.data.forEach(function(element, index) {
                            ctx.fillStyle = '#ffffff';
                            ctx.font = 'bold 14px Arial';
                            ctx.textAlign = 'center';
                            ctx.textBaseline = 'middle';
                            const dataString = dataset.data[index].toString();
                            const padding = 5;
                            ctx.fillText(dataString, element.x, element.y - padding);
                        });
                    }
                });
            }
        }]
    });
    
    // Salvar referência
    if (canvasId === 'chartUnidades') {
        chartUnidades = newChart;
    } else if (canvasId === 'chartEspecialidades') {
        chartEspecialidades = newChart;
    }
}

// ===================================
// ATUALIZAR TABELA
// ===================================
function updateTable() {
    const tbody = document.getElementById('tableBody');
    const footer = document.getElementById('tableFooter');
    tbody.innerHTML = '';
    
    if (filteredData.length === 0) {
        tbody.innerHTML = '<tr><td colspan="8" class="loading-message"><i class="fas fa-inbox"></i> Nenhum registro encontrado</td></tr>';
        footer.textContent = 'Mostrando 0 registros';
        return;
    }
    
    filteredData.forEach(item => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${item['N° Solicitação'] || item['Número da Solicitação'] || '-'}</td>
            <td>${formatDate(item['Data da Solicitação'])}</td>
            <td>${item['Nº Prontuário'] || item['N° Prontuário'] || '-'}</td>
            <td>${item['Telefone'] || '-'}</td>
            <td>${item['Unidade Solicitante'] || '-'}</td>
            <td>${item['Cbo Especialidade'] || '-'}</td>
            <td>${formatDate(item['Data Início da Pendência'])}</td>
            <td>${item['Status'] || '-'}</td>
        `;
        tbody.appendChild(row);
    });
    
    // Atualizar rodapé
    const total = allData.length;
    const showing = filteredData.length;
    footer.textContent = `Mostrando de 1 até ${showing} de ${total} registros`;
}

// ===================================
// FUNÇÕES AUXILIARES
// ===================================
function parseDate(dateString) {
    if (!dateString) return null;
    
    // Tentar DD/MM/YYYY
    let match = dateString.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (match) {
        return new Date(match[3], match[2] - 1, match[1]);
    }
    
    // Tentar YYYY-MM-DD
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
// DOWNLOAD EXCEL
// ===================================
function downloadExcel() {
    if (filteredData.length === 0) {
        alert('Não há dados para exportar.');
        return;
    }
    
    // Preparar dados
    const exportData = filteredData.map(item => ({
        'Nº Solicitação': item['N° Solicitação'] || item['Número da Solicitação'] || '',
        'Data Solicitação': item['Data da Solicitação'] || '',
        'Nº Prontuário': item['Nº Prontuário'] || item['N° Prontuário'] || '',
        'Telefone': item['Telefone'] || '',
        'Unidade Solicitante': item['Unidade Solicitante'] || '',
        'CBO Especialidade': item['Cbo Especialidade'] || '',
        'Data Início Pendência': item['Data Início da Pendência'] || '',
        'Status': item['Status'] || '',
        'Prestador': item['Prestador'] || ''
    }));
    
    // Criar workbook
    const ws = XLSX.utils.json_to_sheet(exportData);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pendências');
    
    // Larguras das colunas
    ws['!cols'] = [
        { wch: 18 }, { wch: 15 }, { wch: 15 }, { wch: 15 },
        { wch: 30 }, { wch: 30 }, { wch: 18 }, { wch: 20 }, { wch: 25 }
    ];
    
    // Download
    const hoje = new Date().toISOString().split('T')[0];
    XLSX.writeFile(wb, `Pendencias_Eldorado_${hoje}.xlsx`);
}
