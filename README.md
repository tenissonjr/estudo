

Este projeto contém programas para processar arquivos CSV  e convertê-los para planilhas Excel com análises e estatísticas.

## 📁 Arquivos Incluídos

### 1. `processar_sivis.py` - **Versão Completa**
- **Função**: Processa dados  com análises detalhadas
- **Características**:
  - Gera múltiplas abas no Excel
  - Estatísticas por portaria, destino, data e hora
  - Resumo executivo com métricas principais
  - Formatação otimizada de datas e horários
  - Ordenação cronológica dos dados

### 2. `conversor_simples.py` - **Versão Básica**
- **Função**: Conversão direta CSV → Excel
- **Características**:
  - Conversão rápida sem análises
  - Uma única aba com todos os dados
  - Ideal para casos simples

## 🚀 Como Usar

### Método 1: Versão Completa (Recomendado)

```bash
python processar_sivis.py
```

**O que acontece:**
- Processa o arquivo `Sivis-Entradas-28-10-2025.csv`
- Gera `Sivis_Entradas_28-10-2025_Processado.xlsx`
- Cria 6 abas com análises diferentes

### Método 2: Versão Simples

```bash
# Usando arquivo padrão
python conversor_simples.py

# Ou especificando arquivo
python conversor_simples.py "meu_arquivo.csv"
```

## 📋 Estrutura do Excel Gerado (Versão Completa)

### Aba 1: **Dados Completos**
- Todos os 3.258 registros originais
- Dados ordenados cronologicamente
- Formato de datas otimizado

### Aba 2: **Estatísticas Portaria** 
- Quantidade de entradas por portaria
- Identificação das portarias mais utilizadas

### Aba 3: **Top 50 Destinos**
- Os 50 destinos mais acessados
- Frequência de visitas por local

### Aba 4: **Entradas por Data**
- Distribuição das entradas ao longo dos dias
- Padrão de utilização temporal

### Aba 5: **Entradas por Hora**
- Distribuição das entradas por horário
- Identificação de picos de movimento

### Aba 6: **Resumo Geral**
- Métricas principais consolidadas
- Visão executiva dos dados

## 📊 Dados Processados (28/10/2025)

- **Total de Registros**: 3.258 entradas
- **Portarias Monitoradas**: 6 diferentes
- **Destinos Únicos**: 677 locais
- **Portaria Mais Utilizada**: Anexo IV - A
- **Período**: 28 de outubro de 2025

## 🔧 Requisitos Técnicos

### Bibliotecas Python:
```
pandas      # Manipulação de dados
openpyxl    # Geração de arquivos Excel
```

### Instalação das dependências:
```bash
pip install pandas openpyxl
```

## 📁 Estrutura de Arquivos

```
📁 estudo/
├── 📄 Sivis-Entradas-28-10-2025.csv          # Arquivo original
├── 📄 processar_sivis.py                      # Programa principal
├── 📄 conversor_simples.py                    # Versão simplificada
├── 📄 README.md                               # Este arquivo
├── 📊 Entradas_28-10-2025_Processado.xlsx  # Excel gerado
└── 📁 .venv/                                  # Ambiente virtual
```

## 💡 Casos de Uso

### Para Análises Detalhadas:
- Use `processar_sivis.py`
- Obtenha relatórios completos
- Análise de padrões de acesso

### Para Conversão Rápida:
- Use `conversor_simples.py`
- Apenas conversão de formato
- Processamento mais rápido

## 🎯 Funcionalidades Especiais

### Tratamento de Dados:
- ✅ Conversão automática de datas
- ✅ Formatação de horários
- ✅ Ordenação cronológica
- ✅ Tratamento de valores vazios
- ✅ Encoding UTF-8

### Análises Estatísticas:
- 📊 Frequência por portaria
- 📈 Distribuição temporal
- 🎯 Destinos mais acessados
- ⏰ Padrões horários
- 📋 Métricas consolidadas

## 🔄 Atualizações Futuras

Para processar novos arquivos SIVIS:

1. **Substitua** o arquivo CSV na pasta
2. **Ajuste** o nome no programa (linha com `arquivo_csv = ...`)
3. **Execute** o programa normalmente
4. **Obtenha** nova planilha Excel atualizada

## 📞 Observações

- Os programas foram desenvolvidos especificamente para o formato SIVIS
- Testado com dados de 28/10/2025
- Compatível com Windows, Linux e macOS
- Geração automática de relatórios executivos


