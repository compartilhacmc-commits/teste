// Configuração e Dados
const SHEET_ID = '1r6NLcVkVLD5vp4UxPEa7TcreBpOd0qeNt-QREOG4Xr4';
const SHEET_NAME = 'pendencias eldorado';
const SHEET_URL = `https://docs.google.com/spreadsheets/d/${SHEET_ID}/gviz/tq?tqx=out:json&sheet=${encodeURIComponent(SHEET_NAME)}`;

let dadosOriginais = [];
let dadosFiltrados = [];
let paginaAtual = 1;
const itensPorPagina = 15;

// Variáveis para gráficos
let chartUnidades, chartEspecialidades, chartVencimento15, chartVencimento30;

// Inicialização
document.addEventListener('DOMContentLoaded', async () => {
    await carregarDados();
    inicializarEventos();
    configurarFiltrosCustomizados();
});

// Carregar dados do Google Sheets
async function carregarDados() {
    try {
        const response = await fetch(SHEET_URL);
        const text = await response.text();
        const json = JSON.parse(text.substring(47).slice(0, -2));
        
        dadosOriginais = json.table.rows.slice(2).map(row => {
            const cells = row.c;
            return {
                numeroSolicitacao: cells[0]?.v || '',
                dataSolicitacao: cells[1]?.f || '',
                prontuario: cells[2]?.v || '',
                usuario: cells[3]?.v || '',
                cns: cells[4]?.v || '',
                nascimento: cells[5]?.f || '',
                telefone: cells[6]?.v || '',
                urgencia: cells[7]?.v || '',
                tipoServico: cells[8]?.v || '',
                unidadeSolicitante: cells[9]?.v || '',
                prestador: cells[10]?.v || '',
                cboEspecialidade: cells[11]?.v || '',
                dataInicioPendencia: cells[12]?.f || '',
                prazo15: cells[13]?.v || '',
                dataFinal15: cells[14]?.f || '',
                dataEnvioEmail15: cells[15]?.f || '',
                prazo30: cells[16]?.v || '',
                dataFinal30: cells[17]?.f || '',
                dataEnvioEmail30: cells[18]?.f || '',
                status: cells[19]?.v || '',
                motivoPendencia: cells[20]?.v || '',
                respostaPendencia: cells[21]?.v || '',
                observacao: cells[22]?.v || ''
            };
        }).filter(item => item.numeroSolicitacao);

        dadosFiltrados = [...dadosOriginais];
        
        // Preencher opções dos filtros
        preencherOpcoesSelect('ubsOptions', [...new Set(dadosOriginais.map(d => d.unidadeSolicitante))].filter(Boolean).sort());
        preencherOpcoesSelect('prestadorOptions', [...new Set(dadosOriginais.map(d => d.prestador))].filter(Boolean).sort());
        preencherOpcoesSelect('especialidadeOptions', [...new Set(dadosOriginais.map(d => d.cboEspecialidade))].filter(Boolean).sort());
        preencherOpcoesSelect('motivoOptions', [...new Set(dadosOriginais.map(d => d.motivoPendencia))].filter(Boolean).sort());
        
        atualizarDashboard();
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
        alert('Erro ao carregar dados da planilha. Verifique a conexão e permissões.');
    }
}

// Preencher opções de select
function preencherOpcoesSelect(containerId, opcoes) {
    const container = document.getElementById(containerId);
    opcoes.forEach(opcao => {
        const label = document.createElement('label');
        label.innerHTML = `<input type="checkbox" value="${opcao}"> ${opcao}`;
        container.appendChild(label);
    });
}

// Configurar filtros customizados
function configurarFiltrosCustomizados() {
    const selects = document.querySelectorAll('.custom-select');
    
    selects.forEach(select => {
        const trigger = select.querySelector('.select-trigger');
        const dropdown = select.querySelector('.select-dropdown');
        
        trigger.addEventListener('click', (e) => {
            e.stopPropagation();
            // Fechar outros dropdowns
            document.querySelectorAll('.select-dropdown').forEach(d => {
                if (d !== dropdown) d.classList.remove('active');
            });
            dropdown.classList.toggle('active');
        });
        
        // Atualizar texto do trigger
        const checkboxes = dropdown.querySelectorAll('input[type="checkbox"]');
        checkboxes.forEach(checkbox => {
            checkbox.addEventListener('change', () => {
                atualizarTextoSelect(select);
                aplicarFiltros();
            });
        });
    });
    
    // Fechar dropdowns ao clicar fora
    document.addEventListener('click', () => {
        document.querySelectorAll('.select-dropdown').forEach(d => d.classList.remove('active'));
    });
}

// Atualizar texto do select
function atualizarTextoSelect(select) {
    const trigger = select.querySelector('.select-trigger span');
    const checkboxes = select.querySelectorAll('input[type="checkbox"]:checked');
    const todosCheckbox = select.querySelector('input[value="Todos"], input[value="Todas"], input[value="Todas as demandas"]');
    
    if (checkboxes.length === 0 || (checkboxes.length === 1 && checkboxes[0] === todosCheckbox)) {
        trigger.textContent = todosCheckbox ? todosCheckbox.value : 'Selecione';
    } else {
        const valores = Array.from(checkboxes)
            .filter(cb => cb !== todosCheckbox)
            .map(cb => cb.value);
        trigger.textContent = valores.length > 2 
            ? `${valores.length} selecionados` 
            : valores.join(', ');
    }
}

// Aplicar filtros
function aplicarFiltros() {
    dadosFiltrados = dadosOriginais.filter(item => {
        // Filtro Status
        const statusSelecionados = obterValoresSelecionados('selectStatus');
        if (!statusSelecionados.includes('Todos') && !statusSelecionados.includes(item.status)) {
            return false;
        }
        
        // Filtro UBS
        const ubsSelecionadas = obterValoresSelecionados('selectUBS');
        if (!ubsSelecionadas.includes('Todas') && !ubsSelecionadas.includes(item.unidadeSolicitante)) {
            return false;
        }
        
        // Filtro Prestador
        const prestadoresSelecionados = obterValoresSelecionados('selectPrestador');
        if (!prestadoresSelecionados.includes('Todos') && !prestadoresSelecionados.includes(item.prestador)) {
            return false;
        }
        
        // Filtro Especialidade
        const especialidadesSelecionadas = obterValoresSelecionados('selectEspecialidade');
        if (!especialidadesSelecionadas.includes('Todas') && !especialidadesSelecionadas.includes(item.cboEspecialidade)) {
            return false;
        }
        
        // Filtro Motivo
        const motivosSelecionados = obterValoresSelecionados('selectMotivo');
        if (!motivosSelecionados.includes('Todas') && !motivosSelecionados.includes('Todas as demandas') && !motivosSelecionados.includes(item.motivoPendencia)) {
            return false;
        }
        
        return true;
    });
    
    paginaAtual = 1;
    atualizarDashboard();
}

// Obter valores selecionados de um select
function obterValoresSelecionados(selectId) {
    const select = document.getElementById(selectId);
    const checkboxes = select.querySelectorAll('input[type="checkbox"]:checked');
    return Array.from(checkboxes).map(cb => cb.value);
}

// Calcular pendências vencendo
function calcularVencimentos(dados, dias) {
    const hoje = new Date();
    const dataLimite = new Date();
    dataLimite.setDate(hoje.getDate() + dias);
    
    return dados.filter(item => {
        if (!item.dataInicioPendencia) return false;
        const dataInicio = parseDate(item.dataInicioPendencia);
        if (!dataInicio) return false;
        
        const dataVencimento = new Date(dataInicio);
        dataVencimento.setDate(dataInicio.getDate() + dias);
        
        return dataVencimento <= dataLimite && dataVencimento >= hoje;
    }).length;
}

// Parsear data
function parseDate(dateStr) {
    try {
        const parts = dateStr.split('/');
        if (parts.length === 3) {
            return new Date(parts[2], parts[1] - 1, parts[0]);
        }
        return new Date(dateStr);
    } catch {
        return null;
    }
}

// Atualizar Dashboard
function atualizarDashboard() {
    atualizarCards();
    atualizarGraficos();
    atualizarTabela();
}

// Atualizar Cards
function atualizarCards() {
    const total = dadosFiltrados.length;
    const vencendo15 = calcularVencimentos(dadosFiltrados, 15);
    const vencendo30 = calcularVencimentos(dadosFiltrados, 30);
    const porcentagem = dadosOriginais.length > 0 
        ? ((total / dadosOriginais.length) * 100).toFixed(1) 
        : 0;
    
    document.getElementById('totalPendencias').textContent = total;
    document.getElementById('vencendo15').textContent = vencendo15;
    document.getElementById('vencendo30').textContent = vencendo30;
    document.getElementById('porcentagem').textContent = porcentagem + '%';
}

// Atualizar Gráficos
function atualizarGraficos() {
    // Gráfico de Unidades
    const unidades = {};
    dadosFiltrados.forEach(item => {
        const unidade = item.unidadeSolicitante || 'Não informado';
        unidades[unidade] = (unidades[unidade] || 0) + 1;
    });
    
    const unidadesOrdenadas = Object.entries(unidades)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);
    
    criarGraficoBarras(
        'chartUnidades',
        unidadesOrdenadas.map(u => u[0]),
        unidadesOrdenadas.map(u => u[1]),
        'rgba(54, 162, 235, 0.8)',
        chartUnidades
    );
    
    // Gráfico de Especialidades
    const especialidades = {};
    dadosFiltrados.forEach(item => {
        const esp = item.cboEspecialidade || 'Não informado';
        especialidades[esp] = (especialidades[esp] || 0) + 1;
    });
    
    const especialidadesOrdenadas = Object.entries(especialidades)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);
    
    criarGraficoBarras(
        'chartEspecialidades',
        especialidadesOrdenadas.map(e => e[0]),
        especialidadesOrdenadas.map(e => e[1]),
        'rgba(75, 192, 192, 0.8)',
        chartEspecialidades
    );
    
    // Gráfico Vencimento 15 dias
    const vencimento15Unidades = {};
    dadosFiltrados.forEach(item => {
        if (calcularVencimentos([item], 15) > 0) {
            const unidade = item.unidadeSolicitante || 'Não informado';
            vencimento15Unidades[unidade] = (vencimento15Unidades[unidade] || 0) + 1;
        }
    });
    
    const venc15Ordenado = Object.entries(vencimento15Unidades)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);
    
    criarGraficoBarras(
        'chartVencimento15',
        venc15Ordenado.map(v => v[0]),
        venc15Ordenado.map(v => v[1]),
        'rgba(255, 159, 64, 0.8)',
        chartVencimento15
    );
    
    // Gráfico Vencimento 30 dias
    const vencimento30Unidades = {};
    dadosFiltrados.forEach(item => {
        if (calcularVencimentos([item], 30) > 0) {
            const unidade = item.unidadeSolicitante || 'Não informado';
            vencimento30Unidades[unidade] = (vencimento30Unidades[unidade] || 0) + 1;
        }
    });
    
    const venc30Ordenado = Object.entries(vencimento30Unidades)
        .sort((a, b) => b[1] - a[1])
        .slice(0, 10);
    
    criarGraficoBarras(
        'chartVencimento30',
        venc30Ordenado.map(v => v[0]),
        venc30Ordenado.map(v => v[1]),
        'rgba(255, 99, 132, 0.8)',
        chartVencimento30
    );
}

// Criar gráfico de barras
function criarGraficoBarras(canvasId, labels, data, backgroundColor, chartInstance) {
    const ctx = document.getElementById(canvasId);
    
    if (chartInstance) {
        chartInstance.destroy();
    }
    
    chartInstance = new Chart(ctx, {
        type: 'bar',
        data: {
            labels: labels,
            datasets: [{
                label: 'Quantidade',
                data: data,
                backgroundColor: backgroundColor,
                borderRadius: 8,
                barThickness: 40
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
                    color: 'white',
                    font: {
                        weight: 'bold',
                        size: 14
                    },
                    anchor: 'center',
                    align: 'center',
                    formatter: (value) => value
                }
            },
            scales: {
                x: {
                    beginAtZero: true,
                    ticks: {
                        stepSize: 1
                    }
                },
                y: {
                    ticks: {
                        autoSkip: false
                    }
                }
            }
        }
    });
    
    // Ajustar altura do canvas
    ctx.style.height = `${Math.max(300, labels.length * 50)}px`;
    
    return chartInstance;
}

// Atualizar Tabela
function atualizarTabela() {
    const tbody = document.getElementById('tabelaBody');
    const inicio = (paginaAtual - 1) * itensPorPagina;
    const fim = inicio + itensPorPagina;
    const dadosPagina = dadosFiltrados.slice(inicio, fim);
    
    tbody.innerHTML = '';
    
    dadosPagina.forEach(item => {
        const tr = document.createElement('tr');
        
        const statusClass = item.status.toLowerCase().replace(/\s+/g, '-');
        
        tr.innerHTML = `
            <td>${item.numeroSolicitacao}</td>
            <td>${item.dataSolicitacao}</td>
            <td>${item.prontuario}</td>
            <td>${item.telefone}</td>
            <td>${item.unidadeSolicitante}</td>
            <td>${item.cboEspecialidade}</td>
            <td>${item.dataInicioPendencia}</td>
            <td><span class="status-badge status-${statusClass}">${item.status}</span></td>
            <td>${item.observacao}</td>
        `;
        
        tbody.appendChild(tr);
    });
    
    // Atualizar informações de paginação
    document.getElementById('showingFrom').textContent = dadosFiltrados.length > 0 ? inicio + 1 : 0;
    document.getElementById('showingTo').textContent = Math.min(fim, dadosFiltrados.length);
    document.getElementById('totalRegistros').textContent = dadosFiltrados.length;
    document.getElementById('paginaAtual').textContent = paginaAtual;
    
    // Atualizar botões de paginação
    const totalPaginas = Math.ceil(dadosFiltrados.length / itensPorPagina);
    document.getElementById('btnAnterior').disabled = paginaAtual === 1;
    document.getElementById('btnProximo').disabled = paginaAtual >= totalPaginas;
    document.getElementById('btnProximo').textContent = paginaAtual < totalPaginas ? (paginaAtual + 1) : paginaAtual;
}

// Eventos
function inicializarEventos() {
    // Botão Atualizar
    document.getElementById('btnAtualizar').addEventListener('click', async () => {
        document.getElementById('btnAtualizar').disabled = true;
        document.getElementById('btnAtualizar').textContent = 'Atualizando...';
        await carregarDados();
        document.getElementById('btnAtualizar').disabled = false;
        document.getElementById('btnAtualizar').innerHTML = `
            <svg width="16" height="16" fill="currentColor" viewBox="0 0 16 16">
                <path fill-rule="evenodd" d="M8 3a5 5 0 1 0 4.546 2.914.5.5 0 0 1 .908-.417A6 6 0 1 1 8 2v1z"/>
                <path d="M8 4.466V.534a.25.25 0 0 1 .41-.192l2.36 1.966c.12.1.12.284 0 .384L8.41 4.658A.25.25 0 0 1 8 4.466z"/>
            </svg>
            Atualizar
        `;
    });
    
    // Botão Limpar Filtros
    document.getElementById('btnLimpar').addEventListener('click', () => {
        document.querySelectorAll('.select-dropdown input[type="checkbox"]').forEach(checkbox => {
            checkbox.checked = checkbox.value === 'Todos' || checkbox.value === 'Todas' || checkbox.value === 'Todas as demandas';
        });
        document.querySelectorAll('.custom-select').forEach(select => atualizarTextoSelect(select));
        aplicarFiltros();
    });
    
    // Botão Excel
    document.getElementById('btnExcel').addEventListener('click', exportarParaExcel);
    
    // Botões de Paginação
    document.getElementById('btnAnterior').addEventListener('click', () => {
        if (paginaAtual > 1) {
            paginaAtual--;
            atualizarTabela();
        }
    });
    
    document.getElementById('btnProximo').addEventListener('click', () => {
        const totalPaginas = Math.ceil(dadosFiltrados.length / itensPorPagina);
        if (paginaAtual < totalPaginas) {
            paginaAtual++;
            atualizarTabela();
        }
    });
}

// Exportar para Excel
function exportarParaExcel() {
    const dadosExport = dadosFiltrados.map(item => ({
        'N° Solicitação': item.numeroSolicitacao,
        'Data da Solicitação': item.dataSolicitacao,
        'Nº Prontuário': item.prontuario,
        'Telefone': item.telefone,
        'Unidade Solicitante': item.unidadeSolicitante,
        'CBO Especialidade': item.cboEspecialidade,
        'Data Início da Pendência': item.dataInicioPendencia,
        'Status': item.status,
        'Observação': item.observacao
    }));
    
    const ws = XLSX.utils.json_to_sheet(dadosExport);
    const wb = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'Pendências');
    
    const dataAtual = new Date().toLocaleDateString('pt-BR').replace(/\//g, '-');
    XLSX.writeFile(wb, `Pendencias_Eldorado_${dataAtual}.xlsx`);
}
