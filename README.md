# Painel de Pendências - Atualização Completa

## 📋 Mudanças Implementadas

### ✅ 1. Seletor de Registros por Página
**Local**: Abaixo da caixa de pesquisa na tabela

**Funcionalidade**: Permite ao usuário escolher quantos registros visualizar por página:
- 10 registros
- 20 registros
- **50 registros (padrão)**
- 500 registros
- 1.000 registros

**Arquivos Alterados**:
- `index.html`: Adicionado novo elemento `<div class="table-records-selector">` com select
- `script.js`: 
  - Variável `TABLE_PAGE_SIZE` agora é dinâmica (padrão 50)
  - Nova função `changeRecordsPerPage()` para atualizar o número de registros
- `styles.css`: Novo estilo para `.table-records-selector`

---

### ✅ 2. Destaque Amarelo para Pendências a Vencer
**Local**: Tabela "Todas as Demandas" - ABA PENDÊNCIAS

**Funcionalidade**: 
- Linhas com fundo **amarelo claro (#fef08a)** para pendências com **15 dias ou mais** a partir da "Data Início da Pendência"
- O cálculo é feito automaticamente na função `updateDemandasTable()`
- Somente aplicado em registros do tipo `PENDENTE`

**Arquivos Alterados**:
- `script.js`: 
  - Adicionado cálculo de dias decorridos em `updateDemandasTable()`
  - Aplicado estilo condicional `tr.style.backgroundColor = '#fef08a'`

---

### ✅ 3. Remoção da Coluna Telefone
**Local**: Tabela "Todas as Demandas"

**Funcionalidade**: 
- Coluna "Telefone" **removida** da visualização da tabela
- A coluna ainda existe nos dados, mas não é mais exibida

**Arquivos Alterados**:
- `script.js`: 
  - Função `updateDemandasTable()` não inclui mais `telefone` no mapeamento de linhas
  - Cabeçalho da tabela atualizado sem coluna "Telefone"

---

### ✅ 4. Remoção de Colunas no Excel (Telefone e Usuário)
**Local**: Planilha baixada pelo botão "Excel"

**Funcionalidade**:
- Coluna "Telefone" **removida**
- Coluna "Usuário" **removida**

**Arquivos Alterados**:
- `script.js`: 
  - Função `downloadExcel()` não inclui mais `Telefone` nem `Usuário` no objeto de exportação

---

### ✅ 5. Adição da Coluna "Nº Solicitação" no Excel
**Local**: Planilha baixada pelo botão "Excel"

**Funcionalidade**:
- Nova coluna **"Nº Solicitação"** adicionada ao arquivo Excel exportado
- Busca automaticamente o valor nas possíveis variações de nome da coluna na planilha original

**Arquivos Alterados**:
- `script.js`: 
  - Função `downloadExcel()` agora inclui:
    ```javascript
    'Nº Solicitação': getColumnValue(item, ['Nº Solicitação', 'Numero Solicitação', 'N Solicitação'], '')
    ```

---

## 📦 Arquivos do Repositório

### Arquivos Principais:
1. **index.html** - Estrutura HTML com novo seletor de registros
2. **script.js** - Lógica JavaScript com todas as modificações
3. **styles.css** - Estilos CSS com novo componente table-records-selector

### Estrutura de Colunas:

#### Tabela Visualizada:
1. Origem
2. Data Solicitação
3. Nº Prontuário
4. Unidade Solicitante
5. CBO Especialidade
6. Data Início da Pendência
7. Status

#### Planilha Excel Exportada:
1. Distrito
2. Tipo
3. **Nº Solicitação** ✨ (NOVO)
4. Data Solicitação
5. Nº Prontuário
6. Prestador
7. Unidade Solicitante
8. CBO Especialidade
9. Data Início da Pendência
10. Status

---

## 🚀 Como Usar

1. **Copie os arquivos** para o repositório do GitHub
2. **Substitua** os arquivos existentes:
   - `index.html`
   - `script.js`
   - `styles.css`
3. **Faça commit e push** das alterações
4. O painel estará atualizado automaticamente no GitHub Pages

---

## 🔍 Detalhes Técnicos

### Destaque Amarelo (15 dias):
```javascript
if (r._item['_tipo'] === 'PENDENTE' && r._dataInicio) {
  const diasDecorridos = Math.floor((hoje - r._dataInicio) / (1000 * 60 * 60 * 24));
  if (diasDecorridos >= 15) {
    tr.style.backgroundColor = '#fef08a'; // Amarelo claro
  }
}
```

### Seletor de Registros:
```javascript
function changeRecordsPerPage() {
  const select = document.getElementById('recordsPerPage');
  TABLE_PAGE_SIZE = parseInt(select.value, 10);
  tableCurrentPage = 1;
  document.getElementById('displayedRecords').textContent = TABLE_PAGE_SIZE;
  updateDemandasTable();
}
```

---

## ✨ Resumo das Melhorias

| Mudança | Status | Impacto |
|---------|--------|---------|
| Seletor registros por página (10, 20, 50, 500, 1000) | ✅ | Melhor controle de visualização |
| Destaque amarelo (15 dias) em Pendências | ✅ | Alerta visual para urgências |
| Remover coluna Telefone da tabela | ✅ | Interface mais limpa |
| Remover Telefone e Usuário do Excel | ✅ | Privacidade de dados |
| Adicionar Nº Solicitação ao Excel | ✅ | Rastreabilidade melhorada |

---

## 📝 Notas Importantes

- O destaque amarelo é aplicado **automaticamente** com base na data
- O padrão de registros por página é **50**
- Todas as planilhas originais permanecem **intactas**
- As mudanças afetam apenas a **visualização** e **exportação**

---

**Criado por**: Sistema Automatizado de Atualização  
**Data**: Janeiro 2026  
**Versão**: 2.0
