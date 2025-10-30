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
        int ou None: ID do destino sugerido ou None se ide_destino não tiver valor
    """
    # Verifica se ide_destino tem valor válido
    if pd.isna(ide_destino) or ide_destino == '' or ide_destino is None:
        return None
    
    # Algoritmo simplificado para gerar sugestão de destino
    # Pode ser substituído por lógica mais complexa conforme necessário
    
    # Combina os parâmetros para gerar um ID único
    hash_base = hash(f"{ide_portaria}_{tim_entrada}_{intervalo_minutos}_{quantidade_minima_entradas}")
    
    # Gera um ID de destino entre 1 e 999
    id_destino_sugerido = abs(hash_base) % 999 + 1
    
    return id_destino_sugerido

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
                
                print(f"   ✓ {simulacao.get('descricao', f'Simulação {i}')}: {intervalo}min, mín {qtd_min} entradas")
            
            # Salva a planilha com as simulações
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