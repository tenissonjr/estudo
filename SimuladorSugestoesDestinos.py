#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Versão Simplificada - Conversor CSV para Excel
Para uso rápido e direto

Estrutura de pastas:
- input/  -> Arquivos CSV de entrada
- output/ -> Planilhas Excel geradas
"""

import pandas as pd
import sys
import os
from datetime import datetime

# Lista de parâmetros para simulações de portaria
simulacoes = [
    {
        "descricao": "Simulação 1",
        "intervalo_minutos": 5,  # Intervalo (em minutos) de apuração da frequência de destinos mais acessados
        "quantidade_minima_entradas": 10,  # Quantidade mínima de entradas para exibição de destinos mais acessados
    },
    {
        "descricao": "Simulação 2",
        "intervalo_minutos": 10,  # Intervalo (em minutos) de apuração da frequência de destinos mais acessados
        "quantidade_minima_entradas": 10,  # Quantidade mínima de entradas para exibição de destinos mais acessados
    },
    {
        "descricao": "Simulação 3",
        "intervalo_minutos": 15,  # Intervalo (em minutos) de apuração da frequência de destinos mais acessados
        "quantidade_minima_entradas": 10,  # Quantidade mínima de entradas para exibição de destinos mais acessados
    },
    {
        "descricao": "Simulação 4",
        "intervalo_minutos": 20,  # Intervalo (em minutos) de apuração da frequência de destinos mais acessados
        "quantidade_minima_entradas": 10,  # Quantidade mínima de entradas para exibição de destinos mais acessados
    },
    {
        "descricao": "Simulação 5",
        "intervalo_minutos": 30,  # Intervalo (em minutos) de apuração da frequência de destinos mais acessados
        "quantidade_minima_entradas": 10,  # Quantidade mínima de entradas para exibição de destinos mais acessados
    },

]

# Variável global para armazenar os dados da planilha
_dados_planilha = None

def carregar_dados_planilha():
    """
    Carrega os dados da planilha em uma variável global para otimizar o acesso
    """
    global _dados_planilha
    if _dados_planilha is None:
        # Procura o arquivo CSV na pasta input
        input_dir = "input"
        arquivos_csv = [f for f in os.listdir(input_dir) if f.endswith('.csv')]
        if arquivos_csv:
            arquivo_csv = os.path.join(input_dir, arquivos_csv[0])
            _dados_planilha = pd.read_csv(arquivo_csv)
            # Converte tim_entrada para datetime para facilitar comparações
            _dados_planilha['tim_entrada'] = pd.to_datetime(_dados_planilha['tim_entrada'])
            # Remove registros sem ide_destino
            _dados_planilha = _dados_planilha.dropna(subset=['ide_destino'])
            _dados_planilha = _dados_planilha[_dados_planilha['ide_destino'].astype(str).str.strip() != '']
    return _dados_planilha

def obterSugestaoDestino(ide_portaria, tim_entrada, intervalo_minutos, quantidade_minima_entradas, ide_destino=None):
    """
    Calcula o destino sugerido com base nos parâmetros da simulação
    
    Args:
        ide_portaria: ID da portaria
        tim_entrada: Timestamp de entrada
        intervalo_minutos: Intervalo em minutos para análise
        quantidade_minima_entradas: Quantidade mínima de entradas
        ide_destino: ID do destino original (opcional)
    
    Returns:
        int ou None: ID do destino sugerido ou None se não encontrar sugestão válida
    """
    # Verifica se ide_destino tem valor válido
    if pd.isna(ide_destino) or ide_destino == '' or ide_destino is None:
        return None
    
    # Carrega os dados da planilha
    dados = carregar_dados_planilha()
    if dados is None or dados.empty:
        return None
    
    # Converte tim_entrada para datetime se for string
    if isinstance(tim_entrada, str):
        tim_entrada_dt = pd.to_datetime(tim_entrada)
    else:
        tim_entrada_dt = pd.to_datetime(tim_entrada)
    
    # Filtra registros da mesma portaria
    dados_portaria = dados[dados['ide_portaria'] == ide_portaria].copy()
    
    # Filtra registros com tim_entrada anterior ao parâmetro
    dados_anteriores = dados_portaria[dados_portaria['tim_entrada'] < tim_entrada_dt].copy()
    
    if dados_anteriores.empty:
        return None
    
    # Calcula a diferença em minutos entre tim_entrada dos registros e o parâmetro
    dados_anteriores['diferenca_minutos'] = (tim_entrada_dt - dados_anteriores['tim_entrada']).dt.total_seconds() / 60
    
    # Filtra registros que estão dentro do intervalo_minutos
    dados_no_intervalo = dados_anteriores[dados_anteriores['diferenca_minutos'] <= intervalo_minutos]
    
    if dados_no_intervalo.empty:
        return None
    
    # Conta a frequência de cada ide_destino
    frequencia_destinos = dados_no_intervalo['ide_destino'].value_counts()
    
    if frequencia_destinos.empty:
        return None
    
    # Obtém o destino mais frequente
    destino_mais_frequente = frequencia_destinos.index[0]
    maior_frequencia = frequencia_destinos.iloc[0]
    
    # Retorna o destino sugerido apenas se a frequência for >= quantidade_minima_entradas
    if maior_frequencia >= quantidade_minima_entradas:
        return destino_mais_frequente
    else:
        return None

def listar_arquivos_input():
    """Lista arquivos CSV disponíveis na pasta input"""
    input_dir = "input"
    if not os.path.exists(input_dir):
        return []
    
    arquivos = [f for f in os.listdir(input_dir) if f.endswith('.csv')]
    return arquivos

def csv_para_excel_simples(arquivo_csv):
    """
    Converte arquivo CSV diretamente para Excel
    """
    try:
        # Lê o CSV
        df = pd.read_csv(arquivo_csv)
        
        # Filtra apenas registros que possuem ide_destino preenchido
        if 'ide_destino' in df.columns:
            total_registros_original = len(df)
            # Remove registros onde ide_destino está vazio, NaN ou None
            df = df.dropna(subset=['ide_destino'])
            # Remove registros onde ide_destino está vazio (string vazia)
            df = df[df['ide_destino'].astype(str).str.strip() != '']
            registros_filtrados = len(df)
            print(f"🔍 Filtro aplicado: registros com 'ide_destino' preenchido")
            print(f"📊 Registros originais: {total_registros_original:,}")
            print(f"📊 Registros filtrados: {registros_filtrados:,}")
            print(f"📊 Registros removidos: {total_registros_original - registros_filtrados:,}")
        else:
            print(f"⚠️ Coluna 'ide_destino' não encontrada. Processando todos os registros.")
        
        # Ordena os dados pelas colunas ide_portaria e tim_entrada
        colunas_ordenacao = []
        if 'ide_portaria' in df.columns:
            colunas_ordenacao.append('ide_portaria')
        if 'tim_entrada' in df.columns:
            colunas_ordenacao.append('tim_entrada')
        
        if colunas_ordenacao:
            df = df.sort_values(by=colunas_ordenacao)
            print(f"📊 Dados ordenados por: {', '.join(colunas_ordenacao)}")
        
        # Criar pasta output se não existir
        output_dir = "output"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        # Nome do arquivo Excel na pasta output
        nome_base = os.path.basename(arquivo_csv).replace('.csv', '.xlsx')
        nome_excel = os.path.join(output_dir, nome_base)
        
        # Salva como Excel inicial
        df.to_excel(nome_excel, index=False)
        
        print(f"✅ Conversão concluída!")
        print(f"📁 Arquivo gerado: {nome_excel}")
        print(f"📊 Registros: {len(df):,}")
        print(f"📋 Colunas: {len(df.columns)}")
        
        # Aplica as simulações se as colunas necessárias existirem
        if 'ide_portaria' in df.columns and 'tim_entrada' in df.columns:
            print(f"\n🔄 Aplicando {len(simulacoes)} simulações...")
            
            # Lê a planilha recém-criada para aplicar simulações
            df_simulacoes = pd.read_excel(nome_excel)
            
            # Para cada simulação, cria uma nova coluna com sugestões de destino
            for i, simulacao in enumerate(simulacoes, 1):
                nome_coluna = f"Simulacao_{i}_Destino"
                nome_coluna_conferencia = f"Simulacao_{i}_Conferencia"
                intervalo = simulacao['intervalo_minutos']
                qtd_min = simulacao['quantidade_minima_entradas']
                
                # Verifica se existe coluna ide_destino
                tem_ide_destino = 'ide_destino' in df_simulacoes.columns
                
                # Aplica a função obterSugestaoDestino para cada linha
                if tem_ide_destino:
                    df_simulacoes[nome_coluna] = df_simulacoes.apply(
                        lambda row: obterSugestaoDestino(
                            row['ide_portaria'], 
                            row['tim_entrada'], 
                            intervalo, 
                            qtd_min,
                            row['ide_destino']
                        ), axis=1
                    )
                else:
                    df_simulacoes[nome_coluna] = df_simulacoes.apply(
                        lambda row: obterSugestaoDestino(
                            row['ide_portaria'], 
                            row['tim_entrada'], 
                            intervalo, 
                            qtd_min
                        ), axis=1
                    )
                
                # Cria coluna de conferência: 1 se simulação == ide_destino, 0 caso contrário
                if tem_ide_destino:
                    df_simulacoes[nome_coluna_conferencia] = df_simulacoes.apply(
                        lambda row: 1 if (pd.notna(row[nome_coluna]) and 
                                        pd.notna(row['ide_destino']) and 
                                        row[nome_coluna] == row['ide_destino']) else 0, axis=1
                    )
                else:
                    # Se não há coluna ide_destino, não é possível fazer conferência
                    df_simulacoes[nome_coluna_conferencia] = 0
                
                print(f"   ✓ {simulacao.get('descricao', f'Simulação {i}')}: {intervalo}min, mín {qtd_min} entradas")
            
            # Cria estatísticas detalhadas
            if tem_ide_destino:
                # Obter informações das portarias
                portarias = sorted(df_simulacoes['ide_portaria'].unique())
                portarias_info = {}
                for portaria in portarias:
                    df_port = df_simulacoes[df_simulacoes['ide_portaria'] == portaria]
                    desc_portaria = df_port['des_portaria'].iloc[0] if 'des_portaria' in df_port.columns else f'Portaria {portaria}'
                    portarias_info[portaria] = {
                        'descricao': desc_portaria,
                        'total_registros': len(df_port)
                    }
                
                # Cria DataFrame para estatísticas por portaria (simulações como linhas, portarias como colunas)
                estatisticas_portaria = []
                
                for i in range(1, len(simulacoes) + 1):
                    simulacao_info = simulacoes[i-1]
                    col_dest = f"Simulacao_{i}_Destino"
                    col_conf = f"Simulacao_{i}_Conferencia"
                    
                    linha_stats = {
                        'Simulacao': f"Simulação {i}",
                        'Descricao': simulacao_info.get('descricao', f'Simulação {i}'),
                        'Intervalo_Min': simulacao_info['intervalo_minutos'],
                        'Qtd_Min_Entradas': simulacao_info['quantidade_minima_entradas']
                    }
                    
                    # Para cada portaria, calcula as estatísticas desta simulação
                    for portaria in portarias:
                        df_port = df_simulacoes[df_simulacoes['ide_portaria'] == portaria]
                        desc_portaria = portarias_info[portaria]['descricao']
                        
                        total_sugestoes = df_port[col_dest].notna().sum()
                        total_acertos = df_port[col_conf].sum()
                        total_registros = len(df_port)
                        
                        if total_sugestoes > 0:
                            precisao = (total_acertos / total_sugestoes) * 100
                            cobertura = (total_sugestoes / total_registros) * 100
                            # Calcula F1-Score (média harmônica entre precisão e cobertura)
                            if precisao > 0 and cobertura > 0:
                                eficiencia = 2 * (precisao * cobertura) / (precisao + cobertura)
                            else:
                                eficiencia = 0
                        else:
                            precisao = 0
                            cobertura = 0
                            eficiencia = 0
                        
                        # Colunas para esta portaria
                        col_prefix = f"Port_{portaria}_{desc_portaria.replace(' ', '_')}"
                        linha_stats[f'{col_prefix}_Registros'] = total_registros
                        linha_stats[f'{col_prefix}_Sugestoes'] = total_sugestoes
                        linha_stats[f'{col_prefix}_Acertos'] = total_acertos
                        linha_stats[f'{col_prefix}_Precisao_Pct'] = round(precisao, 1)
                        linha_stats[f'{col_prefix}_Cobertura_Pct'] = round(cobertura, 1)
                        linha_stats[f'{col_prefix}_Eficiencia_F1'] = round(eficiencia, 1)
                    
                    estatisticas_portaria.append(linha_stats)
                
                df_stats_portaria = pd.DataFrame(estatisticas_portaria)
                
                # Cria estatísticas gerais
                estatisticas_gerais = []
                
                for i in range(1, len(simulacoes) + 1):
                    col_dest = f"Simulacao_{i}_Destino"
                    col_conf = f"Simulacao_{i}_Conferencia"
                    simulacao_info = simulacoes[i-1]
                    
                    total_sugestoes = df_simulacoes[col_dest].notna().sum()
                    total_acertos = df_simulacoes[col_conf].sum()
                    total_registros = len(df_simulacoes)
                    
                    if total_sugestoes > 0:
                        precisao = (total_acertos / total_sugestoes) * 100
                        cobertura = (total_sugestoes / total_registros) * 100
                        # Calcula F1-Score (média harmônica entre precisão e cobertura)
                        if precisao > 0 and cobertura > 0:
                            eficiencia = 2 * (precisao * cobertura) / (precisao + cobertura)
                        else:
                            eficiencia = 0
                    else:
                        precisao = 0
                        cobertura = 0
                        eficiencia = 0
                    
                    linha_geral = {
                        'Simulacao': f"Simulação {i}",
                        'Descricao': simulacao_info.get('descricao', f'Simulação {i}'),
                        'Intervalo_Minutos': simulacao_info['intervalo_minutos'],
                        'Qtd_Min_Entradas': simulacao_info['quantidade_minima_entradas'],
                        'Total_Registros': total_registros,
                        'Total_Sugestoes': total_sugestoes,
                        'Total_Acertos': total_acertos,
                        'Precisao_Pct': round(precisao, 1),
                        'Cobertura_Pct': round(cobertura, 1),
                        'Eficiencia_F1': round(eficiencia, 1)
                    }
                    
                    estatisticas_gerais.append(linha_geral)
                
                df_stats_gerais = pd.DataFrame(estatisticas_gerais)
                
                # Salva o arquivo Excel com múltiplas abas e formatação
                with pd.ExcelWriter(nome_excel, engine='openpyxl') as writer:
                    # Aba principal com dados e simulações
                    df_simulacoes.to_excel(writer, sheet_name='Dados_e_Simulacoes', index=False)
                    
                    # Aba com estatísticas gerais
                    df_stats_gerais.to_excel(writer, sheet_name='Estatisticas_Gerais', index=False)
                    
                    # Aba com estatísticas por portaria
                    df_stats_portaria.to_excel(writer, sheet_name='Estatisticas_por_Portaria', index=False)
                    
                    # Aplica formatação condicional nas abas de estatísticas
                    from openpyxl.formatting.rule import ColorScaleRule
                    from openpyxl.styles import Font, Alignment
                    
                    # Formatação para Estatísticas Gerais
                    ws_gerais = writer.sheets['Estatisticas_Gerais']
                    
                    # Encontra as colunas de Precisão, Cobertura e Eficiência
                    header_row = 1
                    precisao_col = None
                    cobertura_col = None
                    eficiencia_col = None
                    
                    for col_idx, cell in enumerate(ws_gerais[header_row], 1):
                        if cell.value == 'Precisao_Pct':
                            precisao_col = col_idx
                        elif cell.value == 'Cobertura_Pct':
                            cobertura_col = col_idx
                        elif cell.value == 'Eficiencia_F1':
                            eficiencia_col = col_idx
                    
                    # Aplica formatação condicional verde para Precisão
                    if precisao_col:
                        precisao_range = f"{ws_gerais.cell(row=2, column=precisao_col).coordinate}:{ws_gerais.cell(row=len(df_stats_gerais)+1, column=precisao_col).coordinate}"
                        rule_precisao = ColorScaleRule(start_type='min', start_color='FFCCCC',
                                                     mid_type='percentile', mid_value=50, mid_color='FFFF99',
                                                     end_type='max', end_color='90EE90')
                        ws_gerais.conditional_formatting.add(precisao_range, rule_precisao)
                    
                    # Aplica formatação condicional verde para Cobertura
                    if cobertura_col:
                        cobertura_range = f"{ws_gerais.cell(row=2, column=cobertura_col).coordinate}:{ws_gerais.cell(row=len(df_stats_gerais)+1, column=cobertura_col).coordinate}"
                        rule_cobertura = ColorScaleRule(start_type='min', start_color='FFCCCC',
                                                      mid_type='percentile', mid_value=50, mid_color='FFFF99',
                                                      end_type='max', end_color='90EE90')
                        ws_gerais.conditional_formatting.add(cobertura_range, rule_cobertura)
                    
                    # Aplica formatação condicional especial para Eficiência (mapa de calor mais intenso)
                    if eficiencia_col:
                        eficiencia_range = f"{ws_gerais.cell(row=2, column=eficiencia_col).coordinate}:{ws_gerais.cell(row=len(df_stats_gerais)+1, column=eficiencia_col).coordinate}"
                        rule_eficiencia = ColorScaleRule(start_type='min', start_color='FF6B6B',
                                                        mid_type='percentile', mid_value=50, mid_color='FFD93D',
                                                        end_type='max', end_color='6BCF7F')
                        ws_gerais.conditional_formatting.add(eficiencia_range, rule_eficiencia)
                    
                    # Formatação para cabeçalhos
                    for cell in ws_gerais[1]:
                        cell.font = Font(bold=True)
                        cell.alignment = Alignment(horizontal='center')
                    
                    # Formatação para Estatísticas por Portaria
                    ws_portaria = writer.sheets['Estatisticas_por_Portaria']
                    
                    # Aplica formatação condicional para todas as colunas de estatísticas
                    for col_idx, cell in enumerate(ws_portaria[1], 1):
                        if cell.value and 'Precisao_Pct' in str(cell.value):
                            precisao_range = f"{ws_portaria.cell(row=2, column=col_idx).coordinate}:{ws_portaria.cell(row=len(df_stats_portaria)+1, column=col_idx).coordinate}"
                            rule = ColorScaleRule(start_type='min', start_color='FFCCCC',
                                                mid_type='percentile', mid_value=50, mid_color='FFFF99',
                                                end_type='max', end_color='90EE90')
                            ws_portaria.conditional_formatting.add(precisao_range, rule)
                        
                        # Formatação para colunas de cobertura
                        elif cell.value and 'Cobertura_Pct' in str(cell.value):
                            cobertura_range = f"{ws_portaria.cell(row=2, column=col_idx).coordinate}:{ws_portaria.cell(row=len(df_stats_portaria)+1, column=col_idx).coordinate}"
                            rule = ColorScaleRule(start_type='min', start_color='FFCCCC',
                                                mid_type='percentile', mid_value=50, mid_color='FFFF99',
                                                end_type='max', end_color='90EE90')
                            ws_portaria.conditional_formatting.add(cobertura_range, rule)
                        
                        # Formatação especial para colunas de eficiência (mapa de calor mais intenso)
                        elif cell.value and 'Eficiencia_F1' in str(cell.value):
                            eficiencia_range = f"{ws_portaria.cell(row=2, column=col_idx).coordinate}:{ws_portaria.cell(row=len(df_stats_portaria)+1, column=col_idx).coordinate}"
                            rule = ColorScaleRule(start_type='min', start_color='FF6B6B',
                                                mid_type='percentile', mid_value=50, mid_color='FFD93D',
                                                end_type='max', end_color='6BCF7F')
                            ws_portaria.conditional_formatting.add(eficiencia_range, rule)
                    
                    # Formatação para cabeçalhos da aba de portarias
                    for cell in ws_portaria[1]:
                        cell.font = Font(bold=True)
                        cell.alignment = Alignment(horizontal='center')
                    
                    # Ajusta largura das colunas
                    for ws in [ws_gerais, ws_portaria]:
                        for column in ws.columns:
                            max_length = 0
                            column_letter = column[0].column_letter
                            for cell in column:
                                try:
                                    if len(str(cell.value)) > max_length:
                                        max_length = len(str(cell.value))
                                except:
                                    pass
                            adjusted_width = min(max_length + 2, 20)
                            ws.column_dimensions[column_letter].width = adjusted_width
                
                print(f"\n✅ Simulações aplicadas com sucesso!")
                print(f"📊 Novas colunas: {len(simulacoes)} simulações + {len(simulacoes)} conferências adicionadas")
                print(f"📋 Criadas 3 abas: Dados_e_Simulacoes, Estatisticas_Gerais, Estatisticas_por_Portaria")
                print(f"📈 Métrica de Eficiência F1-Score adicionada (combina precisão e cobertura)")
                print(f"🎨 Formatação condicional aplicada: mapa de cores para precisão, cobertura e eficiência")
                
                # Mostra estatísticas de conferência
                print(f"\n📈 Estatísticas de Acertos:")
                for _, row in df_stats_gerais.iterrows():
                    print(f"   {row['Simulacao']}: {row['Total_Acertos']}/{row['Total_Sugestoes']} acertos ({row['Precisao_Pct']:.1f}% precisão, {row['Cobertura_Pct']:.1f}% cobertura, {row['Eficiencia_F1']:.1f} F1-Score)")
            
            else:
                # Se não tem ide_destino, salva apenas os dados
                df_simulacoes.to_excel(nome_excel, index=False)
                print(f"\n✅ Simulações aplicadas com sucesso!")
                print(f"📊 Novas colunas: {len(simulacoes)} simulações adicionadas")
        else:
            print(f"\n⚠️ Colunas 'ide_portaria' ou 'tim_entrada' não encontradas. Simulações não aplicadas.")
        
        print(f"🕐 Processado em: {datetime.now().strftime('%H:%M:%S')}")
        
        return nome_excel
        
    except Exception as e:
        print(f"❌ Erro: {e}")
        return None

if __name__ == "__main__":
    # Usar arquivo como argumento ou padrão
    if len(sys.argv) > 1:
        arquivo = sys.argv[1]
    else:
        arquivo = os.path.join("input", "Entradas-28-10-2025.csv")
    
    if os.path.exists(arquivo):
        csv_para_excel_simples(arquivo)
    else:
        print(f"❌ Arquivo '{arquivo}' não encontrado.")
        print("\n📁 Arquivos disponíveis na pasta 'input':")
        arquivos_disponiveis = listar_arquivos_input()
        if arquivos_disponiveis:
            for i, arq in enumerate(arquivos_disponiveis, 1):
                print(f"   {i}. {arq}")
        else:
            print("   (Nenhum arquivo CSV encontrado)")
        print("\n💡 Dica: Coloque seus arquivos CSV na pasta 'input'")