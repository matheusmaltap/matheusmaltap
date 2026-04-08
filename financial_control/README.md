# 📊 Planilha de Controle Financeiro

Gerador de planilha Excel para controle de **receitas** e **despesas** mensais, com resumo anual e gráficos automáticos.

## ✨ Funcionalidades

- **12 abas mensais** (Janeiro a Dezembro) com:
  - Seção de receitas com categorias pré-definidas
  - Seção de despesas com categorias pré-definidas
  - Cálculo automático de diferença entre previsto e realizado
  - Cálculo automático do saldo mensal
- **Resumo anual** consolidado automaticamente via fórmulas
- **Gráfico de barras**: Receitas vs Despesas realizadas por mês
- **Aba de categorias** personalizáveis
- **Aba de instruções** de uso
- Formatação profissional com cores e bordas

## 🚀 Como usar

### Pré-requisitos

```bash
pip install -r requirements.txt
```

### Gerar a planilha

```bash
# Gera para o ano atual
python financial_control.py

# Gera para um ano específico
python financial_control.py 2025

# Gera para um ano específico com nome de arquivo customizado
python financial_control.py 2025 meu_controle_2025.xlsx
```

### Preencher os dados

1. Abra o arquivo `.xlsx` gerado no Excel, LibreOffice Calc ou Google Sheets.
2. Navegue para a aba do mês desejado (ex: **Janeiro**).
3. Preencha os valores nas colunas **Valor Previsto** e **Valor Realizado**.
4. O saldo e as diferenças são calculados **automaticamente**.
5. O **Resumo Anual** é atualizado automaticamente conforme você preenche os meses.

## 📂 Estrutura da planilha

| Aba | Descrição |
|-----|-----------|
| 📊 Resumo Anual | Consolidado de todos os meses + gráfico de barras |
| Janeiro … Dezembro | Lançamentos mensais de receitas e despesas |
| ⚙️ Categorias | Lista de categorias personalizáveis |
| 📋 Instruções | Guia de uso da planilha |

## 🏷️ Categorias padrão

### Receitas
- Salário
- Freelance / Renda extra
- Investimentos
- Outras receitas

### Despesas
- Moradia (aluguel / financiamento)
- Alimentação
- Transporte
- Saúde
- Educação
- Lazer e entretenimento
- Vestuário
- Serviços (internet, telefone, streaming)
- Seguros
- Outras despesas
