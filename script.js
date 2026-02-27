// Mapeamento de unidades para distritos
const unidadeDistritoMap = {
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
    'UNIDADE BASICA DE SAUDE SÃO JUDAS TADEU': 'VARGEM DAS FLORES',
    'UNIDADE BASICA DE SAUDE VILA ESPERANCA': 'VARGEM DAS FLORES',
    'UNIDADE BASICA DE SAUDE DARCY RIBEIRO': 'VARGEM DAS FLORES',
    'UNIDADE BASICA DE SAUDE ICAIVERA': 'VARGEM DAS FLORES',
    'UNIDADE BASICA DE SAUDE NOVA CONTAGEM I': 'VARGEM DAS FLORES',
    'CONTAGEM PENITENCIARIA NELSON HUNGRIA': 'VARGEM DAS FLORES',
    'UNIDADE BASICA DE SAUDE TUPÃ': 'VARGEM DAS FLORES',
    'UNIDADE BASICA DE SAUDE LIBERDADE II': 'VARGEM DAS FLORES',
};

// Variáveis globais
let dados = [];
let filteredData = [];
let tableData = [];
let currentPage = 1;
const itemsPerPage = 15;
let chartEspecialidade = null;
let chartDistrito = null;

// Elementos do DOM
const filtroEspecialidade = document.getElementById('filtroEspecialidade');
const filtroDistrito = document.getElementById('filtroDistrito');
const filtroTipo = document.getElementById('filtroTipo');
const btnLimpar = document.getElementById('btnLimpar');
const btnExcel = document.getElementById('btnExcel');
const btnAtualizar = document.getElementById('btnAtualizar');
const tabelaDados = document.getElementById('tabelaDados');
const btnAnterior = document.getElementById('btnAnterior');
const btnProximo = document.getElementById('btnProximo');
const paginacao = document.getElementById('paginacao');
const infoRegistros = document.getElementById('infoRegistros');

// Event Listeners
filtroEspecialidade.addEventListener('change', aplicarFiltros);
filtroDistrito.addEventListener('change', aplicarFiltros);
filtroTipo.addEventListener('change', aplicarFiltros);
btnLimpar.addEventListener('click', limparFiltros);
btnExcel.addEventListener('click', baixarExcel);
btnAtualizar.addEventListener('click', () => location.reload());
btnAnterior.addEventListener('click', paginaAnterior);
btnProximo.addEventListener('click', proximaPagina);

// Carregar dados ao iniciar
document.addEventListener('DOMContentLoaded', carregarDados);

// Função para carregar dados da planilha Google
async function carregarDados() {
    try {
        const sheetId = '1gGIHpkw9Osr_881n5Vke7Fb3LWs2Z0p1';
        const gid = '1698493941';
        const csvUrl = `https://docs.google.com/spreadsheets/d/${sheetId}/export?format=csv&gid=${gid}`;
        
        const response = await fetch(csvUrl);
        const csv = await response.text();
        
        const lines = csv.split('\n');
        const headers = lines[0].split(',').map(h => h.trim());
        
        for (let i = 1; i < lines.length; i++) {
            if (lines[i].trim()) {
                const values = lines[i].split(',');
                const obj = {};
                headers.forEach((header, index) => {
                    obj[header] = values[index]?.trim() || '';
                });
                dados.push(obj);
            }
        }
        
        filteredData = [...dados];
        preencherFiltros();
        atualizarTabela();
        atualizarGraficos();
    } catch (error) {
        console.error('Erro ao carregar dados:', error);
    }
}

// Preencher opções dos filtros
function preencherFiltros() {
    const especialidades = [...new Set(dados.map(d => d['ESPECIALIDADE']).filter(Boolean))].sort();
    const distritos = [...new Set(Object.values(unidadeDistritoMap))].sort();
    
    especialidades.forEach(esp => {
        const option = document.createElement('option');
        option.value = esp;
        option.textContent = esp;
        filtroEspecialidade.appendChild(option);
    });
    
    distritos.forEach(dist => {
        const option = document.createElement('option');
        option.value = dist;
        option.textContent = dist;
        filtroDistrito.appendChild(option);
    });
}

// Aplicar filtros
function aplicarFiltros() {
    filteredData = dados.filter(item => {
        let match = true;
        
        if (filtroEspecialidade.value) {
            match = match && item['ESPECIALIDADE']?.includes(filtroEspecialidade.value);
        }
        
        if (filtroDistrito.value) {
            const unidade = item['UNIDADE SOLICITANTE'] || '';
            const distrito = unidadeDistritoMap[unidade] || unidade;
            match = match && distrito === filtroDistrito.value;
        }
        
        if (filtroTipo.value) {
            const tipo = item['TIPO DE ATENDIMENTO'] === 'P' ? 'PRIMEIRA CONSULTA' : 'RETORNO';
            match = match && tipo === filtroTipo.value;
        }
        
        return match;
    });
    
    currentPage = 1;
    atualizarTabela();
    atualizarGraficos();
}

// Limpar filtros
function limparFiltros() {
    filtroEspecialidade.value = '';
    filtroDistrito.value = '';
    filtroTipo.value = '';
    filteredData = [...dados];
    currentPage = 1;
    atualizarTabela();
    atualizarGraficos();
}

// Atualizar tabela
function atualizarTabela() {
    // Consolidar dados
    const map = {};
    
    filteredData.forEach(item => {
        const key = `${item['UNIDADE EXECUTANTE']}_${item['NOME CBO']}_${item['ESPECIALIDADE']}_${item['NOME PROFISSIONAL']}`;
        
        if (!map[key]) {
            map[key] = {
                unidadeExecutante: item['UNIDADE EXECUTANTE'] || '',
                cbo: item['NOME CBO'] || '',
                especialidade: item['ESPECIALIDADE'] || '',
                nomeProfissional: item['NOME PROFISSIONAL'] || '',
                primeiraConsulta: 0,
                retorno: 0,
                total: 0,
            };
        }
        
        if (item['TIPO DE ATENDIMENTO'] === 'P') {
            map[key].primeiraConsulta += 1;
        } else if (item['TIPO DE ATENDIMENTO'] === 'R') {
            map[key].retorno += 1;
        }
        
        map[key].total = map[key].primeiraConsulta + map[key].retorno;
    });
    
    tableData = Object.values(map);
    
    // Renderizar tabela
    const tbody = tabelaDados.querySelector('tbody');
    tbody.innerHTML = '';
    
    const startIndex = (currentPage - 1) * itemsPerPage;
    const paginatedData = tableData.slice(startIndex, startIndex + itemsPerPage);
    
    paginatedData.forEach(item => {
        const tr = document.createElement('tr');
        tr.innerHTML = `
            <td class="font-medium">${item.unidadeExecutante}</td>
            <td>${item.cbo}</td>
            <td>${item.especialidade}</td>
            <td>${item.nomeProfissional}</td>
            <td class="text-center font-semibold" style="color: #2563eb;">${item.primeiraConsulta}</td>
            <td class="text-center font-semibold" style="color: #16a34a;">${item.retorno}</td>
            <td class="text-center font-bold">${item.total}</td>
        `;
        tbody.appendChild(tr);
    });
    
    // Atualizar paginação
    const totalPages = Math.ceil(tableData.length / itemsPerPage);
    paginacao.innerHTML = '';
    
    for (let i = 1; i <= Math.min(totalPages, 5); i++) {
        const btn = document.createElement('button');
        btn.className = `btn btn-page ${i === currentPage ? 'btn-blue' : 'btn-outline'}`;
        btn.textContent = i;
        btn.addEventListener('click', () => {
            currentPage = i;
            atualizarTabela();
        });
        paginacao.appendChild(btn);
    }
    
    btnAnterior.disabled = currentPage === 1;
    btnProximo.disabled = currentPage === totalPages;
    
    infoRegistros.textContent = `Mostrando ${startIndex + 1} a ${Math.min(currentPage * itemsPerPage, tableData.length)} de ${tableData.length} registros`;
}

// Atualizar gráficos
function atualizarGraficos() {
    // Dados por especialidade
    const mapEspecialidade = {};
    filteredData.forEach(item => {
        const esp = item['ESPECIALIDADE'] || 'Sem especialidade';
        mapEspecialidade[esp] = (mapEspecialidade[esp] || 0) + 1;
    });
    
    const dataEspecialidade = Object.entries(mapEspecialidade)
        .map(([name, value]) => ({ name, value }))
        .sort((a, b) => b.value - a.value);
    
    const totalEspecialidade = dataEspecialidade.reduce((sum, item) => sum + item.value, 0);
    document.getElementById('totalEspecialidade').textContent = totalEspecialidade;
    
    // Dados por distrito
    const mapDistrito = {};
    filteredData.forEach(item => {
        const unidade = item['UNIDADE SOLICITANTE'] || '';
        const distrito = unidadeDistritoMap[unidade] || 'Outro';
        mapDistrito[distrito] = (mapDistrito[distrito] || 0) + 1;
    });
    
    const dataDistrito = Object.entries(mapDistrito)
        .map(([name, value]) => ({ name, value }))
        .sort((a, b) => b.value - a.value);
    
    const totalDistrito = dataDistrito.reduce((sum, item) => sum + item.value, 0);
    document.getElementById('totalDistrito').textContent = totalDistrito;
    
    // Gráfico de especialidades
    const ctxEspecialidade = document.getElementById('chartEspecialidade').getContext('2d');
    if (chartEspecialidade) chartEspecialidade.destroy();
    
    chartEspecialidade = new Chart(ctxEspecialidade, {
        type: 'barH',
        data: {
            labels: dataEspecialidade.map(d => d.name),
            datasets: [{
                label: 'Quantidade',
                data: dataEspecialidade.map(d => d.value),
                backgroundColor: '#2563eb',
                borderRadius: 5,
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                x: {
                    beginAtZero: true
                }
            }
        }
    });
    
    // Gráfico de distritos
    const ctxDistrito = document.getElementById('chartDistrito').getContext('2d');
    if (chartDistrito) chartDistrito.destroy();
    
    chartDistrito = new Chart(ctxDistrito, {
        type: 'barH',
        data: {
            labels: dataDistrito.map(d => d.name),
            datasets: [{
                label: 'Quantidade',
                data: dataDistrito.map(d => d.value),
                backgroundColor: '#dc2626',
                borderRadius: 5,
            }]
        },
        options: {
            indexAxis: 'y',
            responsive: true,
            maintainAspectRatio: true,
            plugins: {
                legend: {
                    display: false
                }
            },
            scales: {
                x: {
                    beginAtZero: true
                }
            }
        }
    });
}

// Paginação
function paginaAnterior() {
    if (currentPage > 1) {
        currentPage--;
        atualizarTabela();
    }
}

function proximaPagina() {
    const totalPages = Math.ceil(tableData.length / itemsPerPage);
    if (currentPage < totalPages) {
        currentPage++;
        atualizarTabela();
    }
}

// Baixar Excel
function baixarExcel() {
    let csv = 'Unidade Executante,CBO,Especialidade,Nome Profissional,Total 1ª Consulta,Total Retorno,Total Geral\n';
    
    tableData.forEach(item => {
        csv += `"${item.unidadeExecutante}","${item.cbo}","${item.especialidade}","${item.nomeProfissional}",${item.primeiraConsulta},${item.retorno},${item.total}\n`;
    });
    
    const blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    const url = URL.createObjectURL(blob);
    
    link.setAttribute('href', url);
    link.setAttribute('download', `agendamentos_${new Date().toISOString().split('T')[0]}.csv`);
    link.style.visibility = 'hidden';
    
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
}
