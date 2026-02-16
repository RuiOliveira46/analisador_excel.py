# 📊 Guia de Utilização - Analisador de Ficheiros Excel

## 🚀 Instalação

Antes de executar o programa, instale as dependências necessárias:

```bash
pip install pandas openpyxl
```

## 📝 Como Usar

### 1. Gerar Ficheiro de Exemplo (Opcional)

Se não tiver um ficheiro Excel para testar, execute:

```bash
python gerar_exemplo.py
```

Isto cria um ficheiro `vendas_exemplo.xlsx` com 200 registos de vendas fictícios.

### 2. Executar o Programa

```bash
python analisador_excel.py
```

## 🎯 Funcionalidades

### 1️⃣ Carregar Ficheiro Excel
- Introduza o caminho do ficheiro (ex: `vendas_exemplo.xlsx`)
- O programa mostra quantas linhas e colunas foram carregadas

### 2️⃣ Informação Básica
- Lista todas as colunas e seus tipos
- Mostra as primeiras 5 linhas
- Útil para conhecer a estrutura dos dados

### 3️⃣ Estatísticas
- Mostra média, mediana, mínimo, máximo
- Identifica valores em falta
- Apenas para colunas numéricas

### 4️⃣ Filtrar Dados
**Passos:**
1. Escolha a coluna para filtrar
2. Selecione o tipo de filtro:
   - Igual a (ex: Cidade = "Lisboa")
   - Maior que (ex: Valor > 1000)
   - Menor que (ex: Quantidade < 5)
   - Contém texto (ex: Nome contém "Silva")
3. Introduza o valor
4. Opção de guardar os resultados filtrados

**Exemplos de uso:**
- Vendas superiores a 500€
- Produtos vendidos em Lisboa
- Vendas de um vendedor específico

### 5️⃣ Agrupar e Resumir
**Passos:**
1. Escolha a coluna para agrupar (ex: Cidade, Vendedor)
2. Escolha a coluna numérica para agregar (ex: Valor_Total)
3. Selecione a operação:
   - Soma: Total de vendas por cidade
   - Média: Valor médio por vendedor
   - Contagem: Número de vendas
   - Máximo/Mínimo: Maior/menor valor

**Exemplos práticos:**
- Total de vendas por cidade
- Número de vendas por produto
- Vendedor com maior venda individual

### 6️⃣ Exportar Colunas Específicas
- Selecione apenas as colunas que precisa
- Útil para criar relatórios simplificados
- Ex: Exportar apenas Nome, Data e Valor

## 💡 Dicas

1. **Nomes de ficheiros:**
   - Use caminhos completos se o ficheiro não estiver na mesma pasta
   - Windows: `C:\Users\Nome\Desktop\dados.xlsx`
   - Mac/Linux: `/home/usuario/documentos/dados.xlsx`

2. **Filtros múltiplos:**
   - Execute a opção 4 várias vezes
   - Cada filtro refina os resultados anteriores

3. **Guardar resultados:**
   - Sempre que filtrar ou agrupar, pode guardar
   - Os ficheiros são guardados na pasta atual

4. **Valores em falta:**
   - Verifique com a opção 3 antes de analisar
   - Colunas com muitos valores em falta podem distorcer estatísticas

## 🔧 Resolução de Problemas

**Erro: "Ficheiro não encontrado"**
- Verifique o caminho do ficheiro
- Use aspas se o caminho tiver espaços

**Erro: "pandas não está instalado"**
```bash
pip install pandas openpyxl
```

**Erro ao filtrar valores numéricos:**
- Não use separadores de milhares
- Use ponto para decimais (500.50 não 500,50)

**Ficheiro muito grande e lento:**
- Use filtros para reduzir os dados
- Exporte apenas colunas necessárias

## 📚 Exemplos de Análises

### Análise de Vendas
1. Carregar ficheiro de vendas
2. Ver estatísticas (opção 3)
3. Filtrar vendas > 1000€ (opção 4)
4. Agrupar total por vendedor (opção 5)

### Relatório por Região
1. Carregar dados
2. Filtrar por cidade específica (opção 4)
3. Exportar apenas colunas relevantes (opção 6)

### Top Performers
1. Carregar dados
2. Agrupar soma de vendas por vendedor (opção 5)
3. Guardar resultado ordenado

## 🎓 Próximos Passos

Pode personalizar o programa adicionando:
- Filtros por intervalo de datas
- Gráficos automáticos
- Exportação para CSV ou PDF
- Cálculos personalizados
- Interface gráfica com tkinter

## 📞 Suporte

Em caso de dúvidas ou sugestões, consulte a documentação do pandas:
https://pandas.pydata.org/docs/
