#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Script para verificar a estrutura da aba Estatisticas_por_Portaria
"""

import pandas as pd

# Lê a aba de estatísticas por portaria
arquivo_excel = "output/Entradas-28-10-2025.xlsx"
df_portaria = pd.read_excel(arquivo_excel, sheet_name='Estatisticas_por_Portaria')

print("=== ESTRUTURA DA ABA ESTATISTICAS_POR_PORTARIA ===")
print(f"📊 Dimensões: {df_portaria.shape[0]} linhas x {df_portaria.shape[1]} colunas")
print(f"📋 Colunas: {list(df_portaria.columns)}")

print("\n=== PRIMEIRAS 10 LINHAS ===")
print(df_portaria.head(10).to_string(index=False))

print(f"\n=== PORTARIAS IDENTIFICADAS ===")
portarias_unicas = df_portaria[['IDE_Portaria', 'Descricao_Portaria']].drop_duplicates().sort_values('IDE_Portaria')
print(portarias_unicas.to_string(index=False))

print(f"\n=== SIMULAÇÕES IDENTIFICADAS ===")
simulacoes_unicas = df_portaria[['Simulacao', 'Descricao', 'Intervalo_Min', 'Qtd_Min_Entradas']].drop_duplicates().sort_values('Simulacao')
print(simulacoes_unicas.to_string(index=False))

print(f"\n=== EXEMPLOS DE RESULTADOS POR PORTARIA ===")
# Mostra alguns exemplos para cada portaria
for portaria in sorted(df_portaria['IDE_Portaria'].unique()):
    df_port_exemplo = df_portaria[df_portaria['IDE_Portaria'] == portaria].head(3)
    desc_portaria = df_port_exemplo.iloc[0]['Descricao_Portaria']
    print(f"\n🏢 Portaria {portaria} - {desc_portaria}:")
    for _, row in df_port_exemplo.iterrows():
        print(f"   {row['Simulacao']}: {row['Total_Acertos']}/{row['Total_Sugestoes']} acertos "
              f"({row['Precisao_Pct']:.1f}% precisão, {row['Cobertura_Pct']:.1f}% cobertura, "
              f"{row['Eficiencia_F1']:.1f} F1-Score)")