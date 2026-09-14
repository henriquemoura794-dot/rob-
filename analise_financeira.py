#!/usr/bin/env python3
"""
ENGIPEC — SISTEMA DE INTELIGÊNCIA CORPORATIVA v6.9
====================================================
v6.9: Detalhamento completo do Contas a Pagar
- Mostra os 5 componentes: Diversar, ICMS, FGTS/INSS, Comissão/Bônus, DDA
- Tabela do Resumo Planilha com todas as colunas
- Gráfico empilhado com todos os componentes
- Tabela de detalhamento com todas as colunas

Uso:
    python dashboard_engipec.py
"""

import pandas as pd
import numpy as np
import json
import webbrowser
import base64
from pathlib import Path
import re

# ==============================================================================
# CONFIG
# ==============================================================================

ARQ_FIN_2025 = "ANALISE_FINANCEIRA_2025.xlsx"
ARQ_FIN_2026 = "ANALISE_FINANCEIRA_2026.xlsx"
ARQ_TETO = "Demonstrativo_Contas_a_Pagar_x_Teto_Orcamentario.xlsx"
ARQ_FLUXO = "FLUXO_OPERACIONAL_PROJECAO_SEMANAL.xlsx"
ARQ_HTML = "DASHBOARD_ENGIPEC.html"
LOGO_PATH = "NOVA-LOGO-ENGIPEC-HOME-CENTER-01.png"

PERCENTUAL_TETO = 0.65

UNIDADES_ORDEM = ['Oeiras', 'Picos', 'Picos II', 'Simplício', 'Teresina', 'Consolidado']
MESES_VALIDOS = ['JAN','FEV','MAR','ABR','MAI','JUN','JUL','AGO','SET','OUT','NOV','DEZ']
ORDEM_MESES = {m: i for i, m in enumerate(MESES_VALIDOS)}

COL_FIN = [
    'Mes','Venda Liquida $','MD Vendas Liquida $','Venda Liquida $ - GRUP',
    'Vlr. Compra (Entrada)','Compra x Venda %','Titulos Recebidos $',
    'MD Titulos Recebidos $','Contas Pagas $','MD Contas Pagas $',
    'Res. Operacional $','Res. %','Compra Mes $'
]
COL_FIN_NUM = ['Venda Liquida $','Vlr. Compra (Entrada)','Res. Operacional $','Titulos Recebidos $','Contas Pagas $']

ANO_FISCAL_MESES = {'SET': 2026, 'OUT': 2026, 'NOV': 2026, 'DEZ': 2026, 'JAN': 2027, 'FEV': 2027}
ANO_FISCAL_PARA_FIN = {2026: 2025, 2027: 2026}
MAPA_UNIDADES_TETO = {'Oeiras': 'Oeiras', 'Picos': 'Picos', 'Picos II': 'Picos II', 'Simplicio': 'Simplício', 'Teresina': 'Teresina', 'Consolidado': 'Consolidado'}

# ==============================================================================
# LOGO
# ==============================================================================

def carregar_logo_base64(caminho=LOGO_PATH):
    try:
        with open(caminho, "rb") as f:
            return f"data:image/png;base64,{base64.b64encode(f.read()).decode()}"
    except FileNotFoundError:
        return None

# ==============================================================================
# UTILITÁRIOS
# ==============================================================================

def _ler_bruto(caminho):
    p = Path(caminho)
    if p.suffix.lower() in ['.xlsx', '.xls']:
        return pd.read_excel(p, header=None, sheet_name=0, engine='openpyxl')
    elif p.suffix.lower() == '.csv':
        try:
            return pd.read_csv(p, header=None, sep=';', encoding='utf-8-sig')
        except UnicodeDecodeError:
            return pd.read_csv(p, header=None, sep=';', encoding='latin-1')
    raise ValueError(f"Formato não suportado: {p.suffix}")

def _conv_float(v):
    if pd.isna(v): return np.nan
    if isinstance(v, (int, float)): return float(v)
    s = str(v).strip()
    if s in ('', '—', '-', 'nan', 'None'): return np.nan
    s = re.sub(r'[^\d,.\-]', '', s)
    if ',' in s and '.' in s: s = s.replace('.', '').replace(',', '.')
    elif ',' in s: s = s.replace(',', '.')
    try: return float(s)
    except ValueError: return np.nan

def _extrair_mes(df_str, idx):
    for cell in df_str.iloc[idx]:
        m = str(cell).strip().upper()
        if m in MESES_VALIDOS: return m
    return None

def _encontrar_col_mes(df):
    for c in range(df.shape[1]):
        v = df.iloc[:, c].astype(str).str.strip().str.upper()
        if v.isin(MESES_VALIDOS).any(): return c
    return None

# ==============================================================================
# PIPELINE FINANCEIRO (Aba 1)
# ==============================================================================

def processar_financeiro():
    print("\n[FIN] Lendo 2026...")
    d26 = _proc_fin(ARQ_FIN_2026)
    print(f"  ✓ {len(d26)} unidades")
    print("[FIN] Lendo 2025...")
    d25 = _proc_fin(ARQ_FIN_2025)
    print(f"  ✓ {len(d25)} unidades")
    registros = []
    for ano_label, dados_ano in [("2026", d26), ("2025", d25)]:
        for unidade, df_u in dados_ano.items():
            for _, row in df_u.iterrows():
                venda = float(row.get('Venda Liquida $', 0) or 0)
                compra = float(row.get('Vlr. Compra (Entrada)', 0) or 0)
                res = float(row.get('Res. Operacional $', 0) or 0)
                rec = float(row.get('Titulos Recebidos $', 0) or 0)
                pag = float(row.get('Contas Pagas $', 0) or 0)
                margem_pct = float(row.get('Res. %', 0) or 0) * 100
                registros.append({'ano': int(ano_label), 'unidade': unidade, 'mes': str(row['Mes']),
                    'mes_num': ORDEM_MESES.get(str(row['Mes']), 0), 'venda': round(venda, 2),
                    'compra': round(compra, 2), 'res': round(res, 2), 'rec': round(rec, 2),
                    'pag': round(pag, 2), 'saldo': round(rec - pag, 2), 'margem_pct': round(margem_pct, 4)})
    meses_por_ano = {}
    for ano in [2025, 2026]:
        meses_por_ano[ano] = sorted(set(r['mes'] for r in registros if r['ano'] == ano), key=lambda m: ORDEM_MESES.get(m, 0))
    unidades = sorted(set(r['unidade'] for r in registros))
    return {'registros': registros, 'anos_disponiveis': [2025, 2026], 'meses_por_ano': meses_por_ano,
        'unidades': unidades, 'unidades_lojas': [u for u in unidades if u != 'Consolidado'], 'meses_ordem': MESES_VALIDOS}

def _proc_fin(caminho):
    df_raw = _ler_bruto(caminho)
    blocos = _identificar_blocos_fin(df_raw)
    dados = {}
    for nome, df_bloco in blocos.items():
        df_limpo = _limpar_bloco_fin(df_bloco)
        if df_limpo is not None and not df_limpo.empty: dados[nome] = df_limpo
    return dados

def _identificar_blocos_fin(df_raw):
    df_str = df_raw.fillna('').astype(str)
    linhas_total = []
    for idx, row in df_str.iterrows():
        for cell in row:
            if str(cell).strip().upper() == 'TOTAL': linhas_total.append(idx); break
    if not linhas_total: return _identificar_blocos_seq_fin(df_raw)
    linhas_mes = []
    for idx, row in df_str.iterrows():
        for cell in row:
            if str(cell).strip().upper() in MESES_VALIDOS: linhas_mes.append(idx); break
    if not linhas_mes: return {}
    blocos_idx = []; grupo = []; ti = 0
    for mi in linhas_mes:
        if ti < len(linhas_total) and mi > linhas_total[ti]:
            if grupo: blocos_idx.append(grupo); grupo = []
            while ti < len(linhas_total) and mi > linhas_total[ti]: ti += 1
        grupo.append(mi)
    if grupo: blocos_idx.append(grupo)
    blocos = {}
    for i, g in enumerate(blocos_idx):
        nome = UNIDADES_ORDEM[i] if i < len(UNIDADES_ORDEM) else f'Unidade_{i+1}'
        blocos[nome] = df_raw.iloc[g[0]:g[-1]+1].copy()
    return blocos

def _identificar_blocos_seq_fin(df_raw):
    df_str = df_raw.fillna('').astype(str)
    linhas_mes = []
    for idx, row in df_str.iterrows():
        for cell in row:
            if str(cell).strip().upper() in MESES_VALIDOS: linhas_mes.append(idx); break
    if not linhas_mes: return {}
    blocos_idx = []; grupo = [linhas_mes[0]]
    for i in range(1, len(linhas_mes)):
        ma = _extrair_mes(df_str, linhas_mes[i]); mp = _extrair_mes(df_str, linhas_mes[i-1])
        if ma and mp and ORDEM_MESES.get(ma, 99) <= ORDEM_MESES.get(mp, 99): blocos_idx.append(grupo); grupo = [linhas_mes[i]]
        else: grupo.append(linhas_mes[i])
    blocos_idx.append(grupo)
    blocos = {}
    for i, g in enumerate(blocos_idx):
        nome = UNIDADES_ORDEM[i] if i < len(UNIDADES_ORDEM) else f'Unidade_{i+1}'
        blocos[nome] = df_raw.iloc[g[0]:g[-1]+1].copy()
    return blocos

def _limpar_bloco_fin(df_bloco):
    df_str = df_bloco.fillna('').astype(str)
    mask = df_str.apply(lambda r: any(str(c).strip().upper() in MESES_VALIDOS for c in r), axis=1)
    df_m = df_bloco[mask].copy()
    if df_m.empty: return None
    col_mes = _encontrar_col_mes(df_m)
    if col_mes is None: return None
    nc = min(len(COL_FIN), df_m.shape[1] - col_mes)
    df_d = df_m.iloc[:, list(range(col_mes, col_mes + nc))].copy()
    df_d.columns = COL_FIN[:nc]
    df_d['Mes'] = df_d['Mes'].astype(str).str.strip().str.upper()
    df_d = df_d[df_d['Mes'] != 'TOTAL']
    for col in COL_FIN_NUM:
        if col in df_d.columns: df_d[col] = df_d[col].apply(_conv_float)
    df_d['_o'] = df_d['Mes'].map(ORDEM_MESES)
    df_d = df_d.sort_values('_o').drop(columns=['_o']).reset_index(drop=True)
    return df_d

# ==============================================================================
# PIPELINE TETO ORÇAMENTÁRIO (Aba 2)
# ==============================================================================

def processar_teto(dados_fin):
    print("\n[TETO] Lendo arquivo de teto orçamentário...")
    df_raw = _ler_bruto(ARQ_TETO)
    print(f"  ✓ {len(df_raw)} linhas brutas")
    df_str = df_raw.fillna('').astype(str)
    meses_bloco1 = ['SET', 'OUT', 'NOV']; meses_bloco2 = ['DEZ', 'JAN', 'FEV']
    unidades_teto_raw = ['Oeiras', 'Picos', 'Picos II', 'Simplicio', 'Teresina', 'Consolidado']
    registros = []
    linha_dez = None
    for idx in range(df_raw.shape[0]):
        row_vals = [str(v).strip().upper() for v in df_str.iloc[idx]]
        if 'DEZ' in row_vals: linha_dez = idx; break
    if linha_dez is None: linha_dez = df_raw.shape[0] // 2
    linha_header2 = None
    for idx in range(linha_dez, df_raw.shape[0]):
        val = str(df_str.iloc[idx, 0]).strip()
        if val.upper() == 'UNIDADE': linha_header2 = idx; break
    if linha_header2 is None: linha_header2 = linha_dez + 1
    for idx in range(0, linha_dez):
        unidade_raw = str(df_raw.iloc[idx, 0]).strip()
        if unidade_raw not in unidades_teto_raw: continue
        unidade = MAPA_UNIDADES_TETO.get(unidade_raw, unidade_raw)
        for m_idx, mes in enumerate(meses_bloco1):
            base_col = 2 + m_idx * 4
            realizado = _conv_float(df_raw.iloc[idx, base_col]) or 0.0
            previsto = _conv_float(df_raw.iloc[idx, base_col + 1]) or 0.0
            total_prev = _conv_float(df_raw.iloc[idx, base_col + 2]) or 0.0
            ano = ANO_FISCAL_MESES.get(mes, 2026)
            registros.append({'unidade': unidade, 'mes': mes, 'ano': ano, 'mes_num': ORDEM_MESES.get(mes, 0),
                'realizado': round(realizado, 2), 'previsto': round(previsto, 2), 'total_previsto': round(total_prev, 2)})
    for idx in range(linha_header2 + 1, df_raw.shape[0]):
        unidade_raw = str(df_raw.iloc[idx, 0]).strip()
        if unidade_raw not in unidades_teto_raw: continue
        unidade = MAPA_UNIDADES_TETO.get(unidade_raw, unidade_raw)
        for m_idx, mes in enumerate(meses_bloco2):
            base_col = 2 + m_idx * 4
            realizado = _conv_float(df_raw.iloc[idx, base_col]) or 0.0
            previsto = _conv_float(df_raw.iloc[idx, base_col + 1]) or 0.0
            total_prev = _conv_float(df_raw.iloc[idx, base_col + 2]) or 0.0
            ano = ANO_FISCAL_MESES.get(mes, 2027)
            registros.append({'unidade': unidade, 'mes': mes, 'ano': ano, 'mes_num': ORDEM_MESES.get(mes, 0),
                'realizado': round(realizado, 2), 'previsto': round(previsto, 2), 'total_previsto': round(total_prev, 2)})
    registros = [r for r in registros if r['unidade'] != 'Consolidado']
    print("\n[TETO] Calculando teto dinâmico (65% do faturamento) — 4 modos...")
    fat_lookup = {}
    for r in dados_fin['registros']: fat_lookup[(r['unidade'], r['mes'], r['ano'])] = r['venda']
    fat_medio = {}
    for unidade in set(r['unidade'] for r in dados_fin['registros']):
        vendas = [r['venda'] for r in dados_fin['registros'] if r['unidade'] == unidade]
        fat_medio[unidade] = sum(vendas) / len(vendas) if vendas else 0
    fat_ultimo_fechado = {}; ultimo_mes_label = ""
    for unidade in set(r['unidade'] for r in dados_fin['registros']):
        rows_u = [r for r in dados_fin['registros'] if r['unidade'] == unidade and r['venda'] > 0]
        if rows_u:
            rows_u.sort(key=lambda r: (r['ano'], r['mes_num']))
            ultimo = rows_u[-1]; fat_ultimo_fechado[unidade] = ultimo['venda']
            ultimo_mes_label = f"{ultimo['mes']}/{ultimo['ano']}"
        else: fat_ultimo_fechado[unidade] = 0
    print(f"  ✓ Último mês fechado: {ultimo_mes_label}")
    fat_ano_anterior = {}
    for r in registros:
        ano_fin = ANO_FISCAL_PARA_FIN.get(r['ano'], r['ano'] - 1)
        fat_ano_anterior[(r['unidade'], r['mes'], r['ano'])] = fat_lookup.get((r['unidade'], r['mes'], ano_fin), fat_medio.get(r['unidade'], 0))
    fat_ia_tendencia = {}
    for r in registros:
        unidade = r['unidade']; ano_fin = ANO_FISCAL_PARA_FIN.get(r['ano'], r['ano'] - 1)
        fat_ultimo = fat_ultimo_fechado.get(unidade, 0)
        fat_mes_ano_ant = fat_lookup.get((unidade, r['mes'], ano_fin), 0)
        fat_med = fat_medio.get(unidade, 0)
        if fat_mes_ano_ant > 0 and fat_ultimo > 0:
            taxa = (fat_ultimo - fat_mes_ano_ant) / fat_mes_ano_ant
            if taxa > 0.05: fat_proj = fat_ultimo * 0.65 + fat_mes_ano_ant * 0.35
            elif taxa < -0.05: fat_proj = fat_ultimo * 0.35 + fat_mes_ano_ant * 0.65
            else: fat_proj = fat_ultimo * 0.55 + fat_mes_ano_ant * 0.45
        elif fat_ultimo > 0: fat_proj = fat_ultimo
        elif fat_mes_ano_ant > 0: fat_proj = fat_mes_ano_ant
        else: fat_proj = fat_med
        fat_ia_tendencia[(unidade, r['mes'], r['ano'])] = fat_proj
    for r in registros:
        unidade = r['unidade']
        fat_m1 = fat_ultimo_fechado.get(unidade, 0)
        r['faturamento_anterior'] = round(fat_m1, 2); r['teto_mes_anterior'] = round(fat_m1 * PERCENTUAL_TETO, 2)
        fat_m2 = fat_medio.get(unidade, 0)
        r['faturamento_medio'] = round(fat_m2, 2); r['teto_medio'] = round(fat_m2 * PERCENTUAL_TETO, 2)
        fat_m3 = fat_ano_anterior.get((unidade, r['mes'], r['ano']), 0)
        r['faturamento_ano_anterior'] = round(fat_m3, 2); r['teto_ano_anterior'] = round(fat_m3 * PERCENTUAL_TETO, 2)
        fat_m4 = fat_ia_tendencia.get((unidade, r['mes'], r['ano']), 0)
        r['faturamento_ia'] = round(fat_m4, 2); r['teto_ia'] = round(fat_m4 * PERCENTUAL_TETO, 2)
        for modo, teto_key in [('percentual_uso', 'teto_mes_anterior'), ('percentual_uso_medio', 'teto_medio'),
            ('percentual_uso_ano_anterior', 'teto_ano_anterior'), ('percentual_uso_ia', 'teto_ia')]:
            teto_val = r[teto_key]
            r[modo] = round(r['total_previsto'] / teto_val * 100, 2) if teto_val > 0 else 0.0
    registros.sort(key=lambda r: (r['ano'], r['mes_num'], r['unidade']))
    print(f"  ✓ {len(registros)} registros (filiais)")
    meses_unicos = sorted(set((r['ano'], r['mes']) for r in registros), key=lambda x: (x[0], ORDEM_MESES.get(x[1], 0)))
    consolidado_mes = []
    for ano, mes in meses_unicos:
        rows_mes = [r for r in registros if r['mes'] == mes and r['ano'] == ano]
        real = sum(r['realizado'] for r in rows_mes); prev = sum(r['previsto'] for r in rows_mes)
        total = sum(r['total_previsto'] for r in rows_mes)
        teto_m1 = sum(r['teto_mes_anterior'] for r in rows_mes); teto_m2 = sum(r['teto_medio'] for r in rows_mes)
        teto_m3 = sum(r['teto_ano_anterior'] for r in rows_mes); teto_m4 = sum(r['teto_ia'] for r in rows_mes)
        fat_m1 = sum(r['faturamento_anterior'] for r in rows_mes); fat_m2 = sum(r['faturamento_medio'] for r in rows_mes)
        fat_m3 = sum(r['faturamento_ano_anterior'] for r in rows_mes); fat_m4 = sum(r['faturamento_ia'] for r in rows_mes)
        consolidado_mes.append({'mes': mes, 'ano': ano, 'mes_label': f"{mes}/{str(ano)[-2:]}",
            'realizado': round(real, 2), 'previsto': round(prev, 2), 'total_previsto': round(total, 2),
            'teto_mes_anterior': round(teto_m1, 2), 'teto_medio': round(teto_m2, 2),
            'teto_ano_anterior': round(teto_m3, 2), 'teto_ia': round(teto_m4, 2),
            'faturamento_anterior': round(fat_m1, 2), 'faturamento_medio': round(fat_m2, 2),
            'faturamento_ano_anterior': round(fat_m3, 2), 'faturamento_ia': round(fat_m4, 2),
            'percentual_uso': round(total / teto_m1 * 100, 2) if teto_m1 > 0 else 0,
            'percentual_uso_medio': round(total / teto_m2 * 100, 2) if teto_m2 > 0 else 0,
            'percentual_uso_ano_anterior': round(total / teto_m3 * 100, 2) if teto_m3 > 0 else 0,
            'percentual_uso_ia': round(total / teto_m4 * 100, 2) if teto_m4 > 0 else 0})
    return {'registros': registros, 'consolidado_mes': consolidado_mes, 'percentual_teto': PERCENTUAL_TETO,
        'meses': [m['mes_label'] for m in consolidado_mes], 'unidades': sorted(set(r['unidade'] for r in registros)),
        'anos_fiscais': sorted(set(r['ano'] for r in registros))}

# ==============================================================================
# PIPELINE FLUXO OPERACIONAL (Aba 3) — v6.9 COM DETALHAMENTO
# ==============================================================================

def processar_fluxo():
    """
    v6.9: Captura TODOS os componentes do Contas a Pagar:
    - Previsto Diversar
    - Previsto Apuração ICMS
    - Previsto FGTS/INSS
    - Previsto Comissão/Bônus
    - DUP Realizado DDA
    - Total Geral
    """
    print("\n[FLUXO] Lendo arquivo de projeção semanal...")
    df_raw = _ler_bruto(ARQ_FLUXO)
    print(f"  ✓ {len(df_raw)} linhas brutas, {df_raw.shape[1]} colunas")

    df_str = df_raw.fillna('').astype(str)
    total_rows = df_raw.shape[0]
    total_cols = df_raw.shape[1]

    # ===== BLOCO 1: Contas a Pagar Semanal — DETALHADO v6.9 =====
    unidades_contas = ['OEIRAS', 'PICOS', 'PICOS II', 'SIMPLICIO', 'TERESINA', 'Total']
    contas_pagar = []
    for idx in range(total_rows):
        unidade_raw = str(df_raw.iloc[idx, 0]).strip()
        if unidade_raw in unidades_contas:
            # Ler TODAS as colunas numéricas (ignorando Percentual 0-1)
            col_values = []
            for col in range(1, total_cols):
                v = _conv_float(df_raw.iloc[idx, col])
                if v is not None and not np.isnan(v):
                    if 0 <= v <= 1.0:
                        continue
                    col_values.append(v)
            # v6.9: Mapear cada componente individualmente
            # Ordem: [0]=Diversar, [1]=ICMS, [2]=FGTS/INSS, [3]=Comissão/Bônus, [4]=DDA, [5]=Total Geral
            contas_pagar.append({
                'unidade': unidade_raw,
                'previsto_diversar': round(col_values[0], 2) if len(col_values) > 0 else 0.0,
                'previsto_icms': round(col_values[1], 2) if len(col_values) > 1 else 0.0,
                'previsto_fgts_inss': round(col_values[2], 2) if len(col_values) > 2 else 0.0,
                'previsto_comissao': round(col_values[3], 2) if len(col_values) > 3 else 0.0,
                'dup_dda': round(col_values[4], 2) if len(col_values) > 4 else 0.0,
                'total_geral': round(col_values[5], 2) if len(col_values) > 5 else (round(col_values[-1], 2) if col_values else 0.0),
            })
        elif 'Média Diária' in unidade_raw:
            break

    # ===== BLOCO 2: Média Diária =====
    dias_uteis = 5
    for idx in range(total_rows):
        for col in range(total_cols):
            if str(df_str.iloc[idx, col]).strip() == 'Dias Úteis':
                for nc in range(col + 1, total_cols):
                    dv = _conv_float(df_raw.iloc[idx, nc])
                    if dv is not None and not np.isnan(dv):
                        dias_uteis = int(dv)
                        break
                break

    total_geral_pagar = 0.0
    for r in contas_pagar:
        if r['unidade'] == 'Total':
            total_geral_pagar = r['total_geral']
            break

    media_diaria = total_geral_pagar / dias_uteis if dias_uteis > 0 else 0

    print(f"  ✓ Total Geral a Pagar: R$ {total_geral_pagar:,.2f}")
    print(f"  ✓ Dias Úteis: {dias_uteis}")
    print(f"  ✓ Média Diária: R$ {media_diaria:,.2f}")
    if contas_pagar:
        r = contas_pagar[0]
        print(f"  ✓ Detalhamento ({r['unidade']}): Diversar={r['previsto_diversar']:,.2f} | ICMS={r['previsto_icms']:,.2f} | FGTS={r['previsto_fgts_inss']:,.2f} | Comissão={r['previsto_comissao']:,.2f} | DDA={r['dup_dda']:,.2f} | Total={r['total_geral']:,.2f}")

    # ===== BLOCO 3: Fluxo de Caixa Semanal =====
    fluxo_caixa = []
    header_found = False
    for idx in range(total_rows):
        val_col0 = str(df_str.iloc[idx, 0]).strip()
        has_saldo = any('Saldo Inicial' in str(df_str.iloc[idx, c]).strip() for c in range(total_cols))
        if val_col0 == 'Unidade' and has_saldo:
            header_found = True
            continue
        if header_found:
            unidade_raw = str(df_raw.iloc[idx, 0]).strip()
            if not unidade_raw or unidade_raw in ('nan', 'None', '—', ''):
                continue
            if 'Demonstrativo' in unidade_raw or 'DPS' in unidade_raw:
                break
            vals = []
            for col in range(1, total_cols):
                v = _conv_float(df_raw.iloc[idx, col])
                if v is not None and not np.isnan(v):
                    vals.append(v)
            fluxo_caixa.append({
                'unidade': unidade_raw,
                'saldo_inicial': round(vals[0], 2) if len(vals) > 0 else 0.0,
                'previsao_recebimento': round(vals[1], 2) if len(vals) > 1 else 0.0,
                'disponibilidade': round(vals[2], 2) if len(vals) > 2 else 0.0,
                'contas_pagar': round(vals[3], 2) if len(vals) > 3 else 0.0,
                'saldo_final': round(vals[4], 2) if len(vals) > 4 else 0.0,
            })

    # ===== BLOCO 4: DPS =====
    dps = []
    for idx in range(total_rows):
        val = str(df_str.iloc[idx, 0]).strip()
        if val.startswith('(+)') or val.startswith('(-)') or val.startswith('(=)'):
            for col in range(1, total_cols):
                v = _conv_float(df_raw.iloc[idx, col])
                if v is not None and not np.isnan(v):
                    dps.append({'label': val, 'valor': round(v, 2)})
                    break

    # ===== BLOCO 5: Inadimplência =====
    inadimplencia = []
    inad_col = None
    inad_header_row = -1
    for idx in range(total_rows):
        for col in range(total_cols):
            val = str(df_str.iloc[idx, col]).strip()
            if 'Inadimpl' in val or 'Diagnóstico' in val or 'Diafnóstico' in val:
                inad_col = col
                inad_header_row = idx
                break
        if inad_col is not None:
            break

    if inad_col is not None and inad_header_row >= 0:
        col_unidade = inad_col
        col_valor = inad_col + 1
        if col_valor < total_cols:
            for idx in range(inad_header_row + 1, total_rows):
                unidade = str(df_str.iloc[idx, col_unidade]).strip()
                if not unidade or unidade in ('nan', 'None', '—', ''):
                    continue
                if 'Inadimpl' in unidade or 'Diagnóstico' in unidade or 'Diafnóstico' in unidade:
                    continue
                v = _conv_float(df_raw.iloc[idx, col_valor])
                if v is not None and not np.isnan(v):
                    inadimplencia.append({'unidade': unidade, 'valor': round(v, 2)})

    print(f"  ✓ Contas a Pagar: {len(contas_pagar)} | Fluxo: {len(fluxo_caixa)} | DPS: {len(dps)} | Inadimplência: {len(inadimplencia)}")

    return {
        'contas_pagar': contas_pagar,
        'fluxo_caixa': fluxo_caixa,
        'dps': dps,
        'inadimplencia': inadimplencia,
        'dias_uteis': dias_uteis,
        'media_diaria': round(media_diaria, 2),
        'saldo_a_pagar': round(total_geral_pagar, 2),
    }

# ==============================================================================
# GERAÇÃO HTML
# ==============================================================================

def gerar_html(data_fin, data_teto, data_fluxo, logo_src):
    json_fin = json.dumps(data_fin, ensure_ascii=False)
    json_teto = json.dumps(data_teto, ensure_ascii=False)
    json_fluxo = json.dumps(data_fluxo, ensure_ascii=False)
    html = _HTML_TEMPLATE
    html = html.replace('/* FIN_DATA */', json_fin)
    html = html.replace('/* TETO_DATA */', json_teto)
    html = html.replace('/* FLUXO_DATA */', json_fluxo)
    if logo_src: html = html.replace('/* LOGO_PLACEHOLDER */', logo_src)
    return html

# ==============================================================================
# TEMPLATE HTML v6.9
# ==============================================================================

_HTML_TEMPLATE = r'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>ENGIPEC BI — Inteligência Corporativa</title>
<link rel="icon" href="data:image/svg+xml,<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 100 100'><text y='.9em' font-size='90'>📊</text></svg>">
<script src="https://cdn.tailwindcss.com"></script>
<script src="https://cdn.jsdelivr.net/npm/apexcharts"></script>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap" rel="stylesheet">
<style>
*{font-family:'Inter',sans-serif;-webkit-font-smoothing:antialiased;-moz-osx-font-smoothing:grayscale;text-rendering:optimizeLegibility}
body{background-color:#0B1120;color:#E2E8F0}
.kpi-value,.audit-table td,.audit-table th,.section-title,.kpi-label,.tag,.filter-label,.filter-select,.filter-toggle span,.bench-badge span,#periodo-info,footer,.tab-btn{text-shadow:0 1px 3px rgba(0,0,0,0.5)}
.glass{background:rgba(15,23,42,.55);backdrop-filter:blur(20px);-webkit-backdrop-filter:blur(20px);border:1px solid rgba(51,65,85,.35);border-radius:16px;box-shadow:0 4px 24px rgba(0,0,0,.25)}
.kpi-card{background:linear-gradient(145deg,rgba(30,41,59,.7) 0%,rgba(15,23,42,.5) 100%);backdrop-filter:blur(20px);border:1px solid rgba(51,65,85,.35);border-radius:14px;box-shadow:0 4px 20px rgba(0,0,0,.2);position:relative;overflow:hidden;transition:transform .2s,box-shadow .2s}
.kpi-card:hover{transform:translateY(-2px);box-shadow:0 8px 28px rgba(0,0,0,.3)}
.kpi-card::before{content:'';position:absolute;top:0;left:0;right:0;height:2px;background:linear-gradient(90deg,transparent 10%,#0EA5E9 50%,transparent 90%)}
.kpi-label{font-size:.62rem;font-weight:600;text-transform:uppercase;letter-spacing:1.2px;color:#94A3B8;margin-bottom:.4rem}
.kpi-value{font-size:1.2rem;font-weight:800;color:#FFFFFF;letter-spacing:-.5px;line-height:1.2;text-shadow:0 1px 4px rgba(0,0,0,0.6)}
.kpi-value.neg{color:#FCA5A5;text-shadow:0 1px 4px rgba(220,38,38,0.3)}
.kpi-value.pos{color:#6EE7B7;text-shadow:0 1px 4px rgba(16,185,129,0.3)}
.tag{display:inline-flex;align-items:center;gap:3px;font-size:.56rem;font-weight:700;padding:2px 7px;border-radius:5px}
.tag-pos{background:rgba(16,185,129,.12);color:#4ADE80}.tag-neg{background:rgba(239,68,68,.12);color:#FCA5A5}
.tag-neu{background:rgba(100,116,139,.12);color:#CBD5E1}.tag-yoy{background:rgba(14,165,233,.12);color:#7DD3FC}
.tag-yoy-neg{background:rgba(168,85,247,.12);color:#D8B4FE}
.section-title{font-size:.68rem;font-weight:700;color:#CBD5E1;text-transform:uppercase;letter-spacing:1.5px}
.filter-label{font-size:.6rem;font-weight:600;text-transform:uppercase;letter-spacing:.8px;color:#94A3B8;margin-bottom:3px;display:block}
.filter-select{background-color:rgba(30,41,59,.8);color:#F1F5F9;border:1px solid rgba(51,65,85,.5);border-radius:8px;padding:6px 12px;font-size:.78rem;font-weight:600;cursor:pointer;outline:none;transition:all .2s}
.filter-select:hover,.filter-select:focus{border-color:#0EA5E9;box-shadow:0 0 0 2px rgba(14,165,233,.15)}
.filter-group{background:rgba(15,23,42,.4);border:1px solid rgba(51,65,85,.3);border-radius:10px;padding:10px 14px}
.filter-toggle{display:flex;align-items:center;gap:8px;cursor:pointer}
.filter-toggle input{appearance:none;width:36px;height:20px;background:rgba(51,65,85,.5);border-radius:10px;position:relative;cursor:pointer;transition:background .2s}
.filter-toggle input:checked{background:rgba(14,165,233,.4)}
.filter-toggle input::before{content:'';position:absolute;top:2px;left:2px;width:16px;height:16px;background:#94A3B8;border-radius:50%;transition:transform .2s,background .2s}
.filter-toggle input:checked::before{transform:translateX(16px);background:#0EA5E9}
.filter-toggle span{font-size:.72rem;font-weight:600;color:#CBD5E1}
.pulse-dot{width:7px;height:7px;border-radius:50%;background-color:#10B981;animation:pulse 2s infinite}
@keyframes pulse{0%,100%{opacity:1}50%{opacity:.4}}
.audit-table th{background-color:rgba(30,41,59,.9);color:#CBD5E1;font-weight:700;font-size:.62rem;text-transform:uppercase;letter-spacing:.5px;padding:8px 10px;text-align:right;border-bottom:2px solid rgba(51,65,85,.5);text-shadow:0 1px 2px rgba(0,0,0,0.5)}
.audit-table th:first-child{text-align:left}
.audit-table td{padding:6px 10px;font-size:.74rem;text-align:right;border-bottom:1px solid rgba(51,65,85,.15);color:#E2E8F0;text-shadow:0 1px 2px rgba(0,0,0,0.4)}
.audit-table td:first-child{text-align:left;font-weight:600;color:#F1F5F9}
.audit-table tr:nth-child(even) td{background-color:rgba(15,23,42,.2)}
.td-neg{color:#FCA5A5 !important;font-weight:700;text-shadow:0 1px 3px rgba(220,38,38,0.4) !important}
.td-pos{color:#6EE7B7 !important;font-weight:600;text-shadow:0 1px 3px rgba(16,185,129,0.4) !important}
.td-neu{color:#94A3B8}
.legend-dot{width:8px;height:8px;border-radius:50%;display:inline-block}
::-webkit-scrollbar{width:5px;height:5px}::-webkit-scrollbar-track{background:#0B1120}::-webkit-scrollbar-thumb{background:#334155;border-radius:3px}
.bench-badge{display:inline-flex;align-items:center;gap:5px;font-size:.6rem;font-weight:600;padding:3px 9px;border-radius:20px;background:rgba(14,165,233,.1);color:#7DD3FC;border:1px solid rgba(14,165,233,.25)}
.bench-badge.off{background:rgba(100,116,139,.1);color:#94A3B8;border-color:rgba(100,116,139,.2)}
.logo-container{background:linear-gradient(135deg,#FFFFFF 0%,#F8FAFC 100%);border-radius:10px;padding:6px 10px;box-shadow:0 2px 8px rgba(0,0,0,0.3),0 0 0 1px rgba(255,255,255,0.1);display:inline-flex;align-items:center;justify-content:center;transition:transform .2s,box-shadow .2s}
.logo-container:hover{transform:scale(1.05);box-shadow:0 4px 12px rgba(0,0,0,0.4),0 0 0 1px rgba(14,165,233,0.3)}
.exec-logo{height:36px;width:auto;object-fit:contain}
.tab-bar{display:flex;gap:8px;padding:0 0 12px 0;border-bottom:1px solid rgba(51,65,85,.4);margin-bottom:16px;flex-wrap:wrap}
.tab-btn{padding:10px 20px;border-radius:10px;font-size:.82rem;font-weight:700;cursor:pointer;transition:all .2s;border:1px solid transparent;background:rgba(30,41,59,.4);color:#94A3B8;display:inline-flex;align-items:center;gap:8px}
.tab-btn:hover{background:rgba(30,41,59,.6);color:#E2E8F0;border-color:rgba(51,65,85,.5)}
.tab-btn.active{background:linear-gradient(135deg,rgba(14,165,233,.15) 0%,rgba(99,102,241,.1) 100%);color:#7DD3FC;border-color:rgba(14,165,233,.3);box-shadow:0 2px 12px rgba(14,165,233,.1)}
.tab-content{display:none}.tab-content.active{display:block}
.teto-bar-wrap{background:rgba(15,23,42,.5);border-radius:8px;height:28px;overflow:hidden;position:relative;border:1px solid rgba(51,65,85,.3)}
.teto-bar-fill{height:100%;border-radius:8px;transition:width .6s ease;position:relative;display:flex;align-items:center;justify-content:flex-end;padding-right:8px}
.teto-bar-fill.green{background:linear-gradient(90deg,#10B981,#34D399)}.teto-bar-fill.orange{background:linear-gradient(90deg,#F97316,#FB923C)}.teto-bar-fill.red{background:linear-gradient(90deg,#DC143C,#EF4444)}
.teto-bar-label{font-size:.68rem;font-weight:700;color:#FFF;text-shadow:0 1px 2px rgba(0,0,0,0.5);white-space:nowrap}
.teto-row{display:flex;align-items:center;gap:12px;padding:8px 0;border-bottom:1px solid rgba(51,65,85,.15)}.teto-row:last-child{border-bottom:none}
.teto-unit-name{font-size:.82rem;font-weight:600;color:#E2E8F0;width:100px;flex-shrink:0}.teto-bar-container{flex-grow:1}.teto-value{font-size:.78rem;font-weight:700;width:120px;text-align:right;flex-shrink:0}
.accordion-item{background:rgba(15,23,42,.4);border:1px solid rgba(51,65,85,.3);border-radius:10px;margin-bottom:8px;overflow:hidden}
.accordion-header{padding:12px 16px;cursor:pointer;display:flex;align-items:center;justify-content:space-between;transition:background .2s}
.accordion-header:hover{background:rgba(30,41,59,.5)}
.accordion-header-left{display:flex;align-items:center;gap:10px}
.accordion-arrow{transition:transform .2s;font-size:.7rem;color:#94A3B8}.accordion-item.open .accordion-arrow{transform:rotate(90deg)}
.accordion-body{display:none;padding:0 16px 12px}.accordion-item.open .accordion-body{display:block}
.accordion-unit-name{font-size:.88rem;font-weight:700;color:#E2E8F0}.accordion-unit-badge{font-size:.62rem;font-weight:700;padding:2px 8px;border-radius:5px}
.teto-mode-btn{padding:6px 14px;border-radius:8px;font-size:.72rem;font-weight:700;cursor:pointer;transition:all .2s;border:1px solid rgba(51,65,85,.5);background:rgba(30,41,59,.4);color:#94A3B8}
.teto-mode-btn.active{background:rgba(14,165,233,.15);color:#7DD3FC;border-color:rgba(14,165,233,.3)}
.dps-row{display:flex;justify-content:space-between;align-items:center;padding:8px 12px;border-bottom:1px solid rgba(51,65,85,.15)}
.dps-label{font-size:.78rem;color:#CBD5E1}.dps-value{font-size:.88rem;font-weight:700}
.alert-box{background:rgba(220,20,60,0.08);border:1px solid rgba(220,20,60,0.3);border-radius:10px;padding:12px 16px;display:flex;align-items:center;gap:10px}
.fluxo-view-btn{padding:6px 14px;border-radius:8px;font-size:.72rem;font-weight:700;cursor:pointer;transition:all .2s;border:1px solid rgba(51,65,85,.5);background:rgba(30,41,59,.4);color:#94A3B8}
.fluxo-view-btn.active{background:rgba(14,165,233,.15);color:#7DD3FC;border-color:rgba(14,165,233,.3)}
.fluxo-section{display:none}.fluxo-section.active{display:block}
.resumo-block-title{font-size:.82rem;font-weight:800;color:#7DD3FC;text-transform:uppercase;letter-spacing:1px;padding:12px 0 8px;border-bottom:1px solid rgba(14,165,233,.2);margin-bottom:8px}
.resumo-obs{font-size:.68rem;color:#64748B;padding:8px 12px;background:rgba(15,23,42,.3);border-radius:6px;margin-top:4px}
</style>
</head>
<body class="min-h-screen p-4 md:p-6">

<header class="flex flex-col md:flex-row items-start md:items-center justify-between mb-4 pb-3 border-b border-slate-700/40 gap-3">
    <div class="flex items-center gap-3">
        <div class="logo-container" id="logo-wrap" style="display:none"><img src="/* LOGO_PLACEHOLDER */" alt="Engipec" class="exec-logo" id="engipec-logo"></div>
        <div><h1 class="text-lg md:text-xl font-black tracking-tight text-white" style="text-shadow:0 2px 4px rgba(0,0,0,0.5)">ENGIPEC</h1>
        <p class="text-xs font-semibold text-sky-400 tracking-widest uppercase">Sistema de Inteligência Corporativa</p></div>
    </div>
    <div class="flex items-center gap-4">
        <div class="bench-badge off" id="bench-badge"><span id="bench-text">📊 2026 only</span></div>
        <div class="flex items-center gap-2"><div class="pulse-dot"></div><span class="text-xs font-bold text-green-400 tracking-wider">LIVE DATA</span></div>
        <div id="periodo-info" class="text-xs text-slate-400"></div>
    </div>
</header>

<div class="tab-bar">
    <button class="tab-btn active" onclick="switchTab('fin')" id="tab-fin">📊 Painel Financeiro Global</button>
    <button class="tab-btn" onclick="switchTab('teto')" id="tab-teto">⚠️ Controle de Teto Orçamentário (OTB)</button>
    <button class="tab-btn" onclick="switchTab('fluxo')" id="tab-fluxo">🗓️ Projeção Semanal de Caixa</button>
</div>

<!-- ===== ABA 1: PAINEL FINANCEIRO ===== -->
<div class="tab-content active" id="content-fin">
    <div class="glass p-4 mb-4"><div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-6 gap-3">
        <div class="filter-group"><label class="filter-label">Unidade</label><select id="f-unidade" class="filter-select w-full" onchange="atualizarTudo()"><option value="Consolidado">Consolidado (Todas)</option><option value="Oeiras">Oeiras</option><option value="Picos">Picos</option><option value="Picos II">Picos II</option><option value="Simplício">Simplício</option><option value="Teresina">Teresina</option></select></div>
        <div class="filter-group"><label class="filter-label">Ano</label><select id="f-ano" class="filter-select w-full" onchange="atualizarMeses(); atualizarTudo()"><option value="2026">2026</option><option value="2025">2025</option></select></div>
        <div class="filter-group"><label class="filter-label">Mês De</label><select id="f-mes-de" class="filter-select w-full" onchange="atualizarTudo()"></select></div>
        <div class="filter-group"><label class="filter-label">Mês Até</label><select id="f-mes-ate" class="filter-select w-full" onchange="atualizarTudo()"></select></div>
        <div class="filter-group"><label class="filter-label">Comparação YoY</label><label class="filter-toggle"><input type="checkbox" id="f-comp-toggle" onchange="toggleComp(); atualizarTudo()"><span>Ativar Benchmark</span></label></div>
        <div class="filter-group" id="f-comp-wrap" style="opacity:.4;pointer-events:none"><label class="filter-label">Comparar com Ano</label><select id="f-ano-comp" class="filter-select w-full" onchange="atualizarTudo()"><option value="2025">2025</option><option value="2026">2026</option></select></div>
    </div></div>
    <div class="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3 mb-4">
        <div class="kpi-card p-4"><div class="kpi-label">Venda Líquida</div><div id="k-fat" class="kpi-value mb-1"></div><div class="flex items-center gap-1 flex-wrap"><div id="k-fat-mom" class="tag"></div><div id="k-fat-yoy" class="tag"></div></div><div id="s-fat" class="mt-2"></div></div>
        <div class="kpi-card p-4"><div class="kpi-label">Resultado Oper.</div><div id="k-res" class="kpi-value mb-1"></div><div class="flex items-center gap-1 flex-wrap"><div id="k-res-mom" class="tag"></div><div id="k-res-yoy" class="tag"></div></div><div id="s-res" class="mt-2"></div></div>
        <div class="kpi-card p-4"><div class="kpi-label">Vlr. Compra</div><div id="k-comp" class="kpi-value mb-1"></div><div class="flex items-center gap-1 flex-wrap"><div id="k-comp-mom" class="tag"></div><div id="k-comp-yoy" class="tag"></div></div><div id="s-comp" class="mt-2"></div></div>
        <div class="kpi-card p-4"><div class="kpi-label">Títulos Receb.</div><div id="k-rec" class="kpi-value mb-1"></div><div class="flex items-center gap-1 flex-wrap"><div id="k-rec-mom" class="tag"></div><div id="k-rec-yoy" class="tag"></div></div><div id="s-rec" class="mt-2"></div></div>
        <div class="kpi-card p-4"><div class="kpi-label">Saldo de Caixa</div><div id="k-saldo" class="kpi-value mb-1"></div><div class="flex items-center gap-1 flex-wrap"><div id="k-saldo-mom" class="tag"></div><div id="k-saldo-yoy" class="tag"></div></div><div id="s-saldo" class="mt-2"></div></div>
        <div class="kpi-card p-4"><div class="kpi-label">Margem Oper.</div><div id="k-marg" class="kpi-value mb-1"></div><div class="flex items-center gap-1 flex-wrap"><div id="k-marg-mom" class="tag"></div><div id="k-marg-yoy" class="tag"></div></div><div id="s-marg" class="mt-2"></div></div>
    </div>
    <div class="glass p-5 mb-4"><div class="section-title mb-3">📊 Benchmarking YoY — Vendas por Mês</div><div id="c-bench"></div></div>
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-4">
        <div class="glass p-5"><div class="section-title mb-3">📈 Evolução Mensal — Vendas vs Compras</div><div id="c-evol"></div></div>
        <div class="glass p-5"><div class="section-title mb-3">📊 Fluxo de Caixa — Entradas vs Saídas</div><div id="c-fluxo"></div></div>
    </div>
    <div class="grid grid-cols-1 lg:grid-cols-3 gap-4 mb-4">
        <div class="glass p-5"><div class="section-title mb-3">🏆 Ranking por Margem Operacional</div><div id="c-rank"></div></div>
        <div class="glass p-5"><div class="section-title mb-3">🥧 Participação nas Vendas</div><div id="c-donut"></div></div>
        <div class="glass p-5"><div class="section-title mb-3">⚖️ Comparação entre Unidades</div><div id="c-comp"></div></div>
    </div>
    <details class="glass p-5 mb-4">
        <summary class="cursor-pointer text-sm font-bold text-slate-300 uppercase tracking-wider" style="text-shadow:0 1px 3px rgba(0,0,0,0.5)">📋 Audit Data • Demonstrativo Mensal</summary>
        <div class="mt-3 flex items-center gap-4 text-xs">
            <div class="flex items-center gap-1.5"><span class="legend-dot" style="background:#6EE7B7;box-shadow:0 0 6px rgba(16,185,129,0.4)"></span><span class="text-slate-300">Positivo (R$ > 0)</span></div>
            <div class="flex items-center gap-1.5"><span class="legend-dot" style="background:#FCA5A5;box-shadow:0 0 6px rgba(220,38,38,0.4)"></span><span class="text-slate-300">Negativo (R$ < 0)</span></div>
            <div class="flex items-center gap-1.5"><span class="legend-dot" style="background:#94A3B8"></span><span class="text-slate-300">Neutro</span></div>
        </div>
        <div class="mt-4 overflow-x-auto"><table class="audit-table w-full"><thead><tr id="th-audit"></tr></thead><tbody id="tb-audit"></tbody></table></div>
    </details>
</div>

<!-- ===== ABA 2: TETO ORÇAMENTÁRIO ===== -->
<div class="tab-content" id="content-teto">
    <div class="glass p-4 mb-4"><div class="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-5 gap-3">
        <div class="filter-group"><label class="filter-label">Unidade</label><select id="teto-f-unidade" class="filter-select w-full" onchange="renderTeto()"><option value="">Todas as Unidades</option></select></div>
        <div class="filter-group"><label class="filter-label">Mês</label><select id="teto-f-mes" class="filter-select w-full" onchange="renderTeto()"><option value="">Todos os Meses</option></select></div>
        <div class="filter-group"><label class="filter-label">Ano Fiscal</label><select id="teto-f-ano" class="filter-select w-full" onchange="renderTeto()"><option value="">Todos os Anos</option></select></div>
        <div class="filter-group" style="grid-column:span 2"><label class="filter-label">Modo de Cálculo do Teto (65% do Faturamento)</label>
        <div class="flex gap-2 mt-1 flex-wrap">
            <button class="teto-mode-btn active" id="mode-anterior" onclick="setTetoMode('anterior')">📊 Mês Anterior (Fechado)</button>
            <button class="teto-mode-btn" id="mode-medio" onclick="setTetoMode('medio')">📈 Faturamento Médio</button>
            <button class="teto-mode-btn" id="mode-ano" onclick="setTetoMode('ano')">📅 Ano Anterior</button>
            <button class="teto-mode-btn" id="mode-ia" onclick="setTetoMode('ia')">🤖 IA — Tendência</button>
        </div></div>
    </div></div>
    <div class="glass p-5 mb-4"><div class="section-title mb-3" id="teto-title-barras">⚠️ Teto Orçamentário (65% — Mês Anterior) — Consolidado por Mês</div><div id="c-teto-barras"></div></div>
    <div class="grid grid-cols-1 lg:grid-cols-2 gap-4 mb-4">
        <div class="glass p-5"><div class="section-title mb-3">📊 Consumo do Orçamento por Unidade</div><div id="teto-barras-unidades"></div></div>
        <div class="glass p-5"><div class="section-title mb-3">📈 Evolução do Consumo — Realizado vs Previsto</div><div id="c-teto-evol"></div></div>
    </div>
    <details class="glass p-5 mb-4" open>
        <summary class="cursor-pointer text-sm font-bold text-slate-300 uppercase tracking-wider" style="text-shadow:0 1px 3px rgba(0,0,0,0.5)">📋 Auditoria Orçamentária — Detalhamento por Unidade</summary>
        <div class="mt-3 flex items-center gap-4 text-xs">
            <div class="flex items-center gap-1.5"><span class="legend-dot" style="background:#10B981"></span><span class="text-slate-300">🟢 Até 80% do teto</span></div>
            <div class="flex items-center gap-1.5"><span class="legend-dot" style="background:#F97316"></span><span class="text-slate-300">🟡 81% a 99% — Atenção</span></div>
            <div class="flex items-center gap-1.5"><span class="legend-dot" style="background:#DC143C"></span><span class="text-slate-300">🔴 ≥100% — Teto Estourado</span></div>
        </div>
        <div class="mt-3 text-xs text-slate-500">💡 Clique em cada unidade para expandir os meses.</div>
        <div id="teto-accordion" class="mt-4"></div>
    </details>
</div>

<!-- ===== ABA 3: FLUXO OPERACIONAL — PROJEÇÃO SEMANAL ===== -->
<div class="tab-content" id="content-fluxo">
    <div class="glass p-4 mb-4">
        <div class="flex items-center gap-3 flex-wrap">
            <label class="filter-label" style="margin-bottom:0">Visualizar:</label>
            <button class="fluxo-view-btn active" id="view-resumo" onclick="setFluxoView('resumo')">📄 Resumo Planilha</button>
            <button class="fluxo-view-btn" id="view-contas" onclick="setFluxoView('contas')">💸 Contas a Pagar</button>
            <button class="fluxo-view-btn" id="view-fluxo" onclick="setFluxoView('fluxo')">💰 Fluxo de Caixa</button>
            <button class="fluxo-view-btn" id="view-dps" onclick="setFluxoView('dps')">📝 DPS</button>
            <button class="fluxo-view-btn" id="view-inad" onclick="setFluxoView('inad')">🔴 Inadimplência</button>
        </div>
    </div>

    <div class="fluxo-section active" id="sec-resumo">
        <div class="grid grid-cols-2 md:grid-cols-4 gap-3 mb-4">
            <div class="kpi-card p-4"><div class="kpi-label">Disponibilidade Total</div><div id="fk-disp" class="kpi-value mb-1"></div><div id="fk-disp-sub" class="text-xs text-slate-500 mt-1"></div></div>
            <div class="kpi-card p-4"><div class="kpi-label">Contas a Pagar</div><div id="fk-pagar" class="kpi-value mb-1"></div><div id="fk-pagar-sub" class="text-xs text-slate-500 mt-1"></div></div>
            <div class="kpi-card p-4"><div class="kpi-label">Média Diária Pagamento</div><div id="fk-media" class="kpi-value mb-1"></div><div id="fk-media-sub" class="text-xs text-slate-500 mt-1"></div></div>
            <div class="kpi-card p-4"><div class="kpi-label">Saldo Final Projetado</div><div id="fk-saldo" class="kpi-value mb-1"></div><div id="fk-saldo-sub" class="text-xs mt-1"></div></div>
        </div>
        <div id="fluxo-alerta" class="mb-4"></div>
        <div class="glass p-5 mb-4">
            <div id="resumo-container"></div>
        </div>
    </div>

    <div class="fluxo-section" id="sec-contas">
        <div class="glass p-5 mb-4"><div class="section-title mb-3">💸 Contas a Pagar Semanal — Composição Detalhada</div><div id="c-fluxo-contas"></div></div>
        <div class="glass p-5 mb-4"><div class="section-title mb-3">📋 Detalhamento — Contas a Pagar por Unidade</div><div class="mt-3 overflow-x-auto"><table class="audit-table w-full"><thead><tr id="th-contas"></tr></thead><tbody id="tb-contas"></tbody></table></div></div>
    </div>

    <div class="fluxo-section" id="sec-fluxo-caixa">
        <div class="glass p-5 mb-4"><div class="section-title mb-3">📊 Balanço por Unidade</div><div id="c-fluxo-balanco"></div></div>
        <div class="glass p-5 mb-4"><div class="section-title mb-3">📋 Detalhamento do Fluxo de Caixa</div><div class="mt-3 overflow-x-auto"><table class="audit-table w-full"><thead><tr id="th-fluxo"></tr></thead><tbody id="tb-fluxo"></tbody></table></div></div>
    </div>

    <div class="fluxo-section" id="sec-dps">
        <div class="glass p-5 mb-4"><div class="section-title mb-3">📝 Demonstrativo de Previsão de Saldo (DPS)</div><div id="fluxo-dps"></div></div>
    </div>

    <div class="fluxo-section" id="sec-inad">
        <div class="glass p-5 mb-4"><div class="section-title mb-3">🔴 Ranking de Inadimplência</div><div id="c-fluxo-inad"></div></div>
    </div>
</div>

<footer class="text-center text-xs text-slate-500 pb-4">ENGIPEC HOME CENTER — Sistema de Inteligência Corporativa v6.9 · 3 Abas · Gerado por Python</footer>

<script>
const FIN_DATA = /* FIN_DATA */;
const TETO_DATA = /* TETO_DATA */;
const FLUXO_DATA = /* FLUXO_DATA */;
let charts = {};
let estado = { unidade:'Consolidado', ano:2026, mesDe:0, mesAte:6, compAtivo:false, anoComp:2025 };
let tetoMode = 'anterior';
let fluxoView = 'resumo';

function switchTab(tab) {
    document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
    document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
    if (tab === 'fin') { document.getElementById('tab-fin').classList.add('active'); document.getElementById('content-fin').classList.add('active'); }
    else if (tab === 'teto') { document.getElementById('tab-teto').classList.add('active'); document.getElementById('content-teto').classList.add('active'); initTetoFiltros(); renderTeto(); }
    else { document.getElementById('tab-fluxo').classList.add('active'); document.getElementById('content-fluxo').classList.add('active'); renderFluxo(); }
}

(function(){ var logo=document.getElementById('engipec-logo'); var wrap=document.getElementById('logo-wrap');
    if(logo && logo.src && !logo.src.endsWith('/*%20LOGO_PLACEHOLDER%20*/') && !logo.src.endsWith('/')){ wrap.style.display='inline-flex'; } })();

function fmtBRL(v){if(v===null||v===undefined||isNaN(v))return 'R$ 0,00';return 'R$ '+v.toLocaleString('pt-BR',{minimumFractionDigits:2,maximumFractionDigits:2})}
function fmtBRL0(v){if(v===null||v===undefined||isNaN(v))return 'R$ 0';return 'R$ '+v.toLocaleString('pt-BR',{minimumFractionDigits:0,maximumFractionDigits:0})}
function fmtPct(v){if(v===null||v===undefined||isNaN(v))return '0,00%';return v.toFixed(2).replace('.',',')+'%'}
function fmtC(v){if(Math.abs(v)>=1e6)return 'R$ '+(v/1e6).toFixed(1)+'M';if(Math.abs(v)>=1e3)return 'R$ '+(v/1e3).toFixed(0)+'K';return fmtBRL(v)}
function tagH(val,label,type){if(val===null||val===0)return '<span class="tag tag-neu">— '+label+'</span>';const a=val>0?'▲':'▼';const cls=type==='yoy'?(val>0?'tag-yoy':'tag-yoy-neg'):(val>0?'tag-pos':'tag-neg');return '<span class="tag '+cls+'">'+a+' '+(val>0?'+':'')+val.toFixed(1)+'% '+label+'</span>';}

// ===== ABA 1: FINANCEIRO =====
function filtrar(ano,mesDe,mesAte,uni){return FIN_DATA.registros.filter(r=>r.ano===ano&&r.mes_num>=mesDe&&r.mes_num<=mesAte&&r.unidade===uni)}
function filtrarComp(anoC,mesDe,mesAte,uni){if(!anoC||!estado.compAtivo)return[];return FIN_DATA.registros.filter(r=>r.ano===anoC&&r.mes_num>=mesDe&&r.mes_num<=mesAte&&r.unidade===uni)}
function filtrarLojas(ano,mesDe,mesAte){return FIN_DATA.registros.filter(r=>r.ano===ano&&r.mes_num>=mesDe&&r.mes_num<=mesAte&&FIN_DATA.unidades_lojas.includes(r.unidade))}
function calcKPIs(d){if(!d.length)return null;const venda=d.reduce((s,r)=>s+r.venda,0),compra=d.reduce((s,r)=>s+r.compra,0),res=d.reduce((s,r)=>s+r.res,0);const rec=d.reduce((s,r)=>s+r.rec,0),pag=d.reduce((s,r)=>s+r.pag,0),saldo=rec-pag;const margem=venda!==0?(res/venda*100):0;const o=[...d].sort((a,b)=>a.mes_num-b.mes_num);let mV=0,mR=0,mS=0,mM=0,mC=0,mRc=0;if(o.length>=2){const u=o[o.length-1],p=o[o.length-2];mV=p.venda!==0?((u.venda-p.venda)/Math.abs(p.venda)*100):0;mR=p.res!==0?((u.res-p.res)/Math.abs(p.res)*100):0;mS=(p.rec-p.pag)!==0?((u.rec-u.pag)-(p.rec-p.pag))/Math.abs(p.rec-p.pag)*100:0;mM=p.margem_pct!==0?(u.margem_pct-p.margem_pct):0;mC=p.compra!==0?((u.compra-p.compra)/Math.abs(p.compra)*100):0;mRc=p.rec!==0?((u.rec-p.rec)/Math.abs(p.rec)*100):0;}return{venda,compra,res,rec,pag,saldo,margem,mV,mR,mS,mM,mC,mRc,sV:o.map(r=>r.venda),sR:o.map(r=>r.res),sS:o.map(r=>r.saldo),sM:o.map(r=>r.margem_pct),sC:o.map(r=>r.compra),sRc:o.map(r=>r.rec),meses:o.map(r=>r.mes)};}
function calcYoY(k,kc){if(!kc)return{yV:0,yR:0,yS:0,yM:0,yC:0,yRc:0};const y=(a,b)=>b!==0?((a-b)/Math.abs(b)*100):0;return{yV:y(k.venda,kc.venda),yR:y(k.res,kc.res),yS:y(k.saldo,kc.saldo),yM:k.margem-kc.margem,yC:y(k.compra,kc.compra),yRc:y(k.rec,kc.rec)};}
if(FIN_DATA===null){document.body.innerHTML='<div class="flex items-center justify-center min-h-screen"><div class="text-center"><div class="text-2xl font-bold text-red-500 mb-2">⚠️ Dados não carregados</div></div></div>';}
else{if(document.readyState!=='loading')init();else document.addEventListener('DOMContentLoaded',init);}
function init(){atualizarMeses();atualizarTudo();}
function toggleComp(){const cb=document.getElementById('f-comp-toggle');const wrap=document.getElementById('f-comp-wrap');if(cb.checked){wrap.style.opacity='1';wrap.style.pointerEvents='auto';estado.compAtivo=true;}else{wrap.style.opacity='.4';wrap.style.pointerEvents='none';estado.compAtivo=false;}}
function atualizarMeses(){const ano=parseInt(document.getElementById('f-ano').value);const meses=FIN_DATA.meses_por_ano[ano]||[];const selDe=document.getElementById('f-mes-de');const selAte=document.getElementById('f-mes-ate');const curDe=parseInt(selDe.value)||0;const curAte=parseInt(selAte.value)||6;selDe.innerHTML='';selAte.innerHTML='';meses.forEach(m=>{const n=FIN_DATA.meses_ordem.indexOf(m);const o1=document.createElement('option');o1.value=n;o1.textContent=m;if(n===curDe||(n===0&&curDe===0))o1.selected=true;selDe.appendChild(o1);const o2=document.createElement('option');o2.value=n;o2.textContent=m;if(n===curAte||(n===6&&curAte===6))o2.selected=true;selAte.appendChild(o2);});if(!selDe.value)selDe.value=0;if(!selAte.value){const u=meses[meses.length-1];selAte.value=FIN_DATA.meses_ordem.indexOf(u);}}
function atualizarTudo(){estado.unidade=document.getElementById('f-unidade').value;estado.ano=parseInt(document.getElementById('f-ano').value);estado.mesDe=parseInt(document.getElementById('f-mes-de').value)||0;estado.mesAte=parseInt(document.getElementById('f-mes-ate').value)||11;estado.anoComp=parseInt(document.getElementById('f-ano-comp').value)||2025;if(estado.mesDe>estado.mesAte){estado.mesDe=estado.mesAte;document.getElementById('f-mes-de').value=estado.mesAte;}const d=filtrar(estado.ano,estado.mesDe,estado.mesAte,estado.unidade);const dc=filtrarComp(estado.anoComp,estado.mesDe,estado.mesAte,estado.unidade);const k=calcKPIs(d),kc=calcKPIs(dc),y=calcYoY(k,kc);if(!k)return;const mn=FIN_DATA.meses_ordem.slice(estado.mesDe,estado.mesAte+1);document.getElementById('periodo-info').textContent=mn[0]+'–'+mn[mn.length-1]+' · '+estado.ano;const bb=document.getElementById('bench-badge');const bt=document.getElementById('bench-text');if(estado.compAtivo){bb.classList.remove('off');bt.textContent='📊 '+estado.anoComp+' vs '+estado.ano;}else{bb.classList.add('off');bt.textContent='📊 '+estado.ano+' only';}atualizarKPIs(k,y);Object.keys(charts).forEach(k2=>{if(charts[k2]&&k2.startsWith('fin_')){charts[k2].destroy();delete charts[k2]}});renderBench(d,dc);renderEvol(d);renderFluxoFin(d);renderRank();renderDonut();renderComp();renderTabela(d,dc);}
function atualizarKPIs(k,y){const set=(id,v,neg,mom,yoy,sid,sd,cor)=>{const el=document.getElementById('k-'+id);el.textContent=fmtBRL(v);el.className='kpi-value mb-1 '+(neg?'neg':'pos');document.getElementById('k-'+id+'-mom').innerHTML=tagH(mom,'MoM','mom');document.getElementById('k-'+id+'-yoy').innerHTML=(estado.compAtivo&&yoy!==0)?tagH(yoy,'YoY','yoy'):'';spark(sid,sd,cor);};set('fat',k.venda,k.venda<0,k.mV,y.yV,'s-fat',k.sV,'#0EA5E9');set('res',k.res,k.res<0,k.mR,y.yR,'s-res',k.sR,'#6366F1');set('comp',k.compra,false,k.mC,y.yC,'s-comp',k.sC,'#8B5CF6');set('rec',k.rec,false,k.mRc,y.yRc,'s-rec',k.sRc,'#06B6D4');set('saldo',k.saldo,k.saldo<0,k.mS,y.yS,'s-saldo',k.sS,k.saldo<0?'#F87171':'#34D399');const em=document.getElementById('k-marg');em.textContent=fmtPct(k.margem);em.className='kpi-value mb-1 '+(k.margem<0?'neg':'pos');document.getElementById('k-marg-mom').innerHTML=tagH(k.mM,'pp MoM','mom');document.getElementById('k-marg-yoy').innerHTML=(estado.compAtivo&&y.yM!==0)?tagH(y.yM,'pp YoY','yoy'):'';spark('s-marg',k.sM,'#FBBF24');}

const F="'Inter',sans-serif",G='#1E293B',A='#94A3B8',L='#CBD5E1';
function spark(id,data,cor){charts['fin_'+id]=new ApexCharts(document.querySelector('#'+id),{chart:{type:'area',height:34,sparkline:{enabled:true},animations:{enabled:false},fontFamily:F},series:[{data:data,color:cor}],fill:{type:'gradient',gradient:{shadeIntensity:1,opacityFrom:.2,opacityTo:.02,stops:[0,100]}},tooltip:{enabled:false}});charts['fin_'+id].render();}
function renderBench(d,dc){const meses=[...new Set(d.map(r=>r.mes))].sort((a,b)=>FIN_DATA.meses_ordem.indexOf(a)-FIN_DATA.meses_ordem.indexOf(b));const vA=meses.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.venda,0));const vC=meses.map(m=>dc.length?dc.filter(r=>r.mes===m).reduce((s,r)=>s+r.venda,0):0);const series=[{name:String(estado.ano),data:vA,color:'#0EA5E9'}];if(dc.length)series.unshift({name:String(estado.anoComp),data:vC,color:'#475569'});charts.fin_bench=new ApexCharts(document.querySelector('#c-bench'),{chart:{type:'bar',height:340,background:'transparent',toolbar:{show:false},fontFamily:F,animations:{enabled:true,speed:600}},plotOptions:{bar:{borderRadius:5,columnWidth:'50%'}},series:series,dataLabels:{enabled:false},xaxis:{categories:meses,labels:{style:{colors:A,fontSize:'12px',fontFamily:F}},axisBorder:{show:false},axisTicks:{show:false}},yaxis:{labels:{style:{colors:A,fontSize:'11px',fontFamily:F},formatter:v=>fmtC(v)}},grid:{borderColor:G,strokeDashArray:4,padding:{top:-20,right:10,bottom:0,left:10}},legend:{position:'top',horizontalAlign:'right',labels:{colors:L},fontSize:'12px',fontFamily:F,markers:{size:6,radius:3}},tooltip:{theme:'dark',y:{formatter:v=>fmtBRL(v)}},fill:{opacity:[.6,.95]}});charts.fin_bench.render();}
function renderEvol(d){const meses=[...new Set(d.map(r=>r.mes))].sort((a,b)=>FIN_DATA.meses_ordem.indexOf(a)-FIN_DATA.meses_ordem.indexOf(b));const v=meses.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.venda,0));const c=meses.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.compra,0));charts.fin_evol=new ApexCharts(document.querySelector('#c-evol'),{chart:{type:'bar',height:320,background:'transparent',toolbar:{show:false},fontFamily:F,animations:{enabled:true,speed:600}},plotOptions:{bar:{borderRadius:4,columnWidth:'55%',dataLabels:{position:'top'}}},series:[{name:'Venda Líquida',data:v,color:'#0EA5E9'},{name:'Compra (Entrada)',data:c,color:'#8B5CF6'}],dataLabels:{enabled:true,formatter:function(val){return fmtC(val);},style:{colors:['#7DD3FC','#C4B5FD'],fontSize:'9px',fontFamily:F,fontWeight:600},offsetY:-16,background:{enabled:false}},xaxis:{categories:meses,labels:{style:{colors:A,fontSize:'12px',fontFamily:F}},axisBorder:{show:false},axisTicks:{show:false}},yaxis:{labels:{style:{colors:A,fontSize:'11px',fontFamily:F},formatter:v=>fmtC(v)}},grid:{borderColor:G,strokeDashArray:4,padding:{top:-20,right:10,bottom:0,left:10}},legend:{position:'top',horizontalAlign:'right',labels:{colors:L},fontSize:'12px',fontFamily:F},tooltip:{theme:'dark',y:{formatter:v=>fmtBRL(v)}},fill:{opacity:[.9,.7]}});charts.fin_evol.render();}
function renderFluxoFin(d){const meses=[...new Set(d.map(r=>r.mes))].sort((a,b)=>FIN_DATA.meses_ordem.indexOf(a)-FIN_DATA.meses_ordem.indexOf(b));const e=meses.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.rec,0));const s=meses.map(m=>d.filter(r=>r.mes===m).reduce((s,r)=>s+r.pag,0));const sl=meses.map((m,i)=>e[i]-s[i]);charts.fin_fluxo=new ApexCharts(document.querySelector('#c-fluxo'),{chart:{type:'bar',height:320,background:'transparent',toolbar:{show:false},fontFamily:F,animations:{enabled:true,speed:600}},plotOptions:{bar:{borderRadius:4,columnWidth:'55%'}},series:[{name:'Entradas',data:e,color:'#10B981'},{name:'Saídas',data:s,color:'#EF4444'},{name:'Saldo',data:sl,type:'line',color:'#0EA5E9'}],dataLabels:{enabled:true,enabledOnSeries:[2],formatter:v=>fmtC(v),style:{colors:['#7DD3FC'],fontSize:'10px',fontFamily:F,fontWeight:600},offsetY:-8},xaxis:{categories:meses,labels:{style:{colors:A,fontSize:'12px',fontFamily:F}},axisBorder:{show:false},axisTicks:{show:false}},yaxis:{labels:{style:{colors:A,fontSize:'11px',fontFamily:F},formatter:v=>fmtC(v)}},grid:{borderColor:G,strokeDashArray:4,padding:{top:-20,right:10,bottom:0,left:10}},legend:{position:'top',horizontalAlign:'right',labels:{colors:L},fontSize:'12px',fontFamily:F},tooltip:{theme:'dark',y:{formatter:v=>fmtBRL(v)}},stroke:{width:[0,0,3],curve:'smooth'},markers:{size:[0,0,5],colors:['#10B981','#EF4444','#0EA5E9'],strokeColors:'#0B1120',strokeWidth:2},fill:{opacity:[.85,.85,1]}});charts.fin_fluxo.render();}
function renderRank(){const lojas=filtrarLojas(estado.ano,estado.mesDe,estado.mesAte);const ranking=FIN_DATA.unidades_lojas.map(u=>{const rows=lojas.filter(r=>r.unidade===u);const fat=rows.reduce((s,r)=>s+r.venda,0),res=rows.reduce((s,r)=>s+r.res,0);const marg=fat!==0?(res/fat*100):0;return{u,marg,fat,res};}).sort((a,b)=>b.marg-a.marg);const nomes=ranking.map(r=>r.u),margens=ranking.map(r=>r.marg);const cores=margens.map(m=>m>=10?'#0EA5E9':(m>=0?'#FBBF24':'#F87171'));charts.fin_rank=new ApexCharts(document.querySelector('#c-rank'),{chart:{type:'bar',height:300,background:'transparent',toolbar:{show:false},fontFamily:F,animations:{enabled:true,speed:600}},plotOptions:{bar:{horizontal:true,borderRadius:6,barHeight:'55%',distributed:true,dataLabels:{position:'center'}}},series:[{name:'Margem %',data:margens}],dataLabels:{enabled:true,formatter:v=>v.toFixed(2).replace('.',',')+'%',style:{colors:['#FFFFFF'],fontSize:'11px',fontFamily:F,fontWeight:700}},xaxis:{categories:nomes,labels:{style:{colors:A},formatter:v=>v.toFixed(1)+'%'}},yaxis:{labels:{style:{colors:'#F1F5F9',fontWeight:600,fontSize:'12px'}}},grid:{borderColor:G,strokeDashArray:4},legend:{show:false},tooltip:{theme:'dark',y:{formatter:v=>v.toFixed(2)+'%'}},colors:cores,fill:{colors:cores}});charts.fin_rank.render();}
function renderDonut(){const lojas=filtrarLojas(estado.ano,estado.mesDe,estado.mesAte);const dados=FIN_DATA.unidades_lojas.map(u=>({u,v:lojas.filter(r=>r.unidade===u).reduce((s,r)=>s+r.venda,0)}));const labels=dados.map(d=>d.u),values=dados.map(d=>d.v);const total=values.reduce((s,v)=>s+v,0);const cores=['#0EA5E9','#3B82F6','#6366F1','#8B5CF6','#06B6D4'];charts.fin_donut=new ApexCharts(document.querySelector('#c-donut'),{chart:{type:'donut',height:320,background:'transparent',toolbar:{show:false},fontFamily:F,animations:{enabled:true,speed:600}},series:values,labels:labels,colors:cores,stroke:{colors:['#0B1120'],width:4},plotOptions:{pie:{donut:{size:'62%',labels:{show:true,name:{show:true,fontSize:'13px',color:L,fontFamily:F},value:{show:true,fontSize:'15px',fontWeight:700,color:'#F1F5F9',fontFamily:F,formatter:function(val){return fmtBRL0(parseFloat(val));}},total:{show:true,showAlways:true,fontSize:'11px',color:'#94A3B8',fontFamily:F,label:'Faturamento Total',formatter:function(){return 'R$ '+(total/1e6).toFixed(2).replace('.',',')+'M';}}}}}},dataLabels:{enabled:true,formatter:function(val){return val.toFixed(1).replace('.',',')+'%';},style:{colors:['#FFFFFF'],fontSize:'12px',fontFamily:F,fontWeight:700},background:{enabled:false},dropShadow:{enabled:true,top:1,left:0,blur:3,opacity:0.5}},legend:{position:'bottom',fontSize:'12px',fontFamily:F,labels:{colors:L},markers:{size:6,radius:3,offsetX:-2},itemMargin:{horizontal:8,vertical:4},formatter:function(seriesName,opts){var val=opts.w.globals.series[opts.seriesIndex];return seriesName+':  '+fmtBRL0(val);}},tooltip:{theme:'dark',y:{formatter:function(val){var pct=(val/total*100).toFixed(2).replace('.',',');return fmtBRL(val)+' ('+pct+'%)';}}},responsive:[{breakpoint:480,options:{legend:{position:'bottom',fontSize:'10px'}}}]});charts.fin_donut.render();}
function renderComp(){const lojas=filtrarLojas(estado.ano,estado.mesDe,estado.mesAte);const dados=FIN_DATA.unidades_lojas.map(u=>{const rows=lojas.filter(r=>r.unidade===u);return{u,v:rows.reduce((s,r)=>s+r.venda,0),r:rows.reduce((s,r)=>s+r.res,0)};});charts.fin_comp=new ApexCharts(document.querySelector('#c-comp'),{chart:{type:'bar',height:300,background:'transparent',toolbar:{show:false},fontFamily:F,animations:{enabled:true,speed:600}},plotOptions:{bar:{borderRadius:4,columnWidth:'50%',dataLabels:{position:'top'}}},series:[{name:'Venda Líquida',data:dados.map(d=>d.v),color:'#0EA5E9'},{name:'Resultado',data:dados.map(d=>d.r),color:'#6366F1'}],dataLabels:{enabled:true,formatter:function(val){return fmtC(val);},style:{colors:['#7DD3FC','#A5B4FC'],fontSize:'9px',fontFamily:F,fontWeight:600},offsetY:-16,background:{enabled:false}},xaxis:{categories:dados.map(d=>d.u),labels:{style:{colors:A,fontSize:'10px'}}},yaxis:{labels:{style:{colors:A},formatter:v=>fmtC(v)}},grid:{borderColor:G,strokeDashArray:4},legend:{position:'top',horizontalAlign:'left',labels:{colors:L},fontSize:'11px'},tooltip:{theme:'dark',y:{formatter:v=>fmtBRL(v)}},fill:{opacity:[.9,.8]}});charts.fin_comp.render();}
function renderTabela(d,dc){const meses=[...new Set(d.map(r=>r.mes))].sort((a,b)=>FIN_DATA.meses_ordem.indexOf(a)-FIN_DATA.meses_ordem.indexOf(b));const th=document.getElementById('th-audit');th.innerHTML='<th>Mês</th><th>Venda '+estado.ano+'</th>';if(dc.length)th.innerHTML+='<th style="color:#64748B">Venda '+estado.anoComp+'</th><th>Var. YoY</th>';th.innerHTML+='<th>Compra</th><th>Res. Oper.</th>';if(dc.length)th.innerHTML+='<th style="color:#64748B">Res. '+estado.anoComp+'</th>';th.innerHTML+='<th>Títulos Rec.</th><th>Contas Pagas</th>';const tb=document.getElementById('tb-audit');tb.innerHTML='';meses.forEach(m=>{const rd=d.filter(r=>r.mes===m);const v=rd.reduce((s,r)=>s+r.venda,0),c=rd.reduce((s,r)=>s+r.compra,0),r=rd.reduce((s,r)=>s+r.res,0),rec=rd.reduce((s,r)=>s+r.rec,0),pag=rd.reduce((s,r)=>s+r.pag,0);const resArrow=r>=0?'▲':'▼';const resClass=r<0?'td-neg':'td-pos';let html='<tr><td>'+m+'</td><td>'+fmtBRL(v)+'</td>';if(dc.length){const rc=dc.filter(x=>x.mes===m);const vc=rc.length>0?rc.reduce((s,x)=>s+x.venda,0):0;const rc2=rc.length>0?rc.reduce((s,x)=>s+x.res,0):0;const yoy=vc!==0?((v-vc)/Math.abs(vc)*100):0;const yoyArrow=yoy>=0?'▲':'▼';const yoyClass=yoy>=0?'td-pos':'td-neg';const res25Arrow=rc2>=0?'▲':'▼';const res25Class=rc2<0?'td-neg':'td-pos';html+='<td class="td-neu">'+fmtBRL(vc)+'</td>';html+='<td class="'+yoyClass+'">'+yoyArrow+' '+(yoy>=0?'+':'')+yoy.toFixed(1)+'%</td>';html+='<td>'+fmtBRL(c)+'</td>';html+='<td class="'+resClass+'">'+resArrow+' '+fmtBRL(r)+'</td>';html+='<td class="'+res25Class+'">'+res25Arrow+' '+fmtBRL(rc2)+'</td>';}else{html+='<td>'+fmtBRL(c)+'</td>';html+='<td class="'+resClass+'">'+resArrow+' '+fmtBRL(r)+'</td>';}html+='<td class="td-neu">'+fmtBRL(rec)+'</td>';html+='<td class="td-neu">▼ '+fmtBRL(pag)+'</td>';html+='</tr>';tb.innerHTML+=html;});}

// ===== ABA 2: TETO =====
function setTetoMode(mode){tetoMode=mode;['anterior','medio','ano','ia'].forEach(m=>{var el=document.getElementById('mode-'+m);if(el)el.classList.toggle('active',mode===m);});renderTeto();}
function getTeto(r){if(tetoMode==='medio')return r.teto_medio;if(tetoMode==='ano')return r.teto_ano_anterior;if(tetoMode==='ia')return r.teto_ia;return r.teto_mes_anterior;}
function getPct(r){if(tetoMode==='medio')return r.percentual_uso_medio;if(tetoMode==='ano')return r.percentual_uso_ano_anterior;if(tetoMode==='ia')return r.percentual_uso_ia;return r.percentual_uso;}
function getFat(r){if(tetoMode==='medio')return r.faturamento_medio;if(tetoMode==='ano')return r.faturamento_ano_anterior;if(tetoMode==='ia')return r.faturamento_ia;return r.faturamento_anterior;}
function initTetoFiltros(){var selU=document.getElementById('teto-f-unidade');if(selU.options.length<=1){TETO_DATA.unidades.forEach(function(u){var o=document.createElement('option');o.value=u;o.textContent=u;selU.appendChild(o);});}var selM=document.getElementById('teto-f-mes');if(selM.options.length<=1){var meses=[...new Set(TETO_DATA.consolidado_mes.map(function(m){return m.mes_label;}))];meses.forEach(function(m){var o=document.createElement('option');o.value=m;o.textContent=m;selM.appendChild(o);});}var selA=document.getElementById('teto-f-ano');if(selA.options.length<=1){TETO_DATA.anos_fiscais.forEach(function(a){var o=document.createElement('option');o.value=a;o.textContent=a;selA.appendChild(o);});}}
function filtrarTeto(){var fU=document.getElementById('teto-f-unidade').value;var fM=document.getElementById('teto-f-mes').value;var fA=document.getElementById('teto-f-ano').value;var reg=TETO_DATA.registros;if(fU)reg=reg.filter(function(r){return r.unidade===fU;});if(fM){var parts=fM.split('/');var mes=parts[0];var anoFull=2000+parseInt(parts[1]);reg=reg.filter(function(r){return r.mes===mes&&r.ano===anoFull;});}if(fA)reg=reg.filter(function(r){return r.ano===parseInt(fA);});return reg;}
function renderTeto(){if(!TETO_DATA||!TETO_DATA.registros)return;Object.keys(charts).forEach(function(k){if(k.startsWith('teto_')){charts[k].destroy();delete charts[k];}});var registros=filtrarTeto();if(registros.length===0){document.getElementById('c-teto-barras').innerHTML='<div class="text-center text-slate-500 py-8">Nenhum registro encontrado.</div>';document.getElementById('teto-barras-unidades').innerHTML='';document.getElementById('c-teto-evol').innerHTML='';document.getElementById('teto-accordion').innerHTML='';return;}var mesesUnicos=[...new Set(registros.map(function(r){return r.mes+'|'+r.ano;}))];var consolidadoMes=[];mesesUnicos.forEach(function(key){var parts=key.split('|');var mes=parts[0];var ano=parseInt(parts[1]);var rows=registros.filter(function(r){return r.mes===mes&&r.ano===ano;});var real=rows.reduce(function(s,r){return s+r.realizado;},0);var prev=rows.reduce(function(s,r){return s+r.previsto;},0);var total=rows.reduce(function(s,r){return s+r.total_previsto;},0);var teto=rows.reduce(function(s,r){return s+getTeto(r);},0);var fat=rows.reduce(function(s,r){return s+getFat(r);},0);var pct=teto>0?total/teto*100:0;consolidadoMes.push({mes:mes,ano:ano,mes_label:mes+'/'+String(ano).slice(-2),realizado:real,previsto:prev,total_previsto:total,teto:teto,faturamento:fat,percentual_uso:pct});});var ordemMeses=FIN_DATA.meses_ordem;consolidadoMes.sort(function(a,b){return a.ano!==b.ano?a.ano-b.ano:ordemMeses.indexOf(a.mes)-ordemMeses.indexOf(b.mes);});var modeLabels={'anterior':'Mês Anterior (Fechado)','medio':'Faturamento Médio','ano':'Ano Anterior','ia':'IA — Tendência'};document.getElementById('teto-title-barras').textContent='⚠️ Teto Orçamentário (65% — '+modeLabels[tetoMode]+') — Consolidado por Mês';renderTetoBarras(consolidadoMes);renderTetoUnidades(registros);renderTetoEvol(consolidadoMes);renderTetoAccordion(registros);}
function renderTetoBarras(dados){var meses=dados.map(function(d){return d.mes_label;});var realizado=dados.map(function(d){return d.realizado;});var previsto=dados.map(function(d){return d.previsto;});var tetoValores=dados.map(function(d){return d.teto;});charts.teto_barras=new ApexCharts(document.querySelector('#c-teto-barras'),{chart:{type:'bar',height:380,background:'transparent',toolbar:{show:false},fontFamily:F,animations:{enabled:true,speed:600}},plotOptions:{bar:{borderRadius:4,columnWidth:'55%',stacked:true}},series:[{name:'Realizado (ERP)',data:realizado,color:'#0EA5E9'},{name:'Previsto (Pedidos)',data:previsto,color:'#8B5CF6'}],dataLabels:{enabled:false},xaxis:{categories:meses,labels:{style:{colors:A,fontSize:'11px',fontFamily:F}},axisBorder:{show:false},axisTicks:{show:false}},yaxis:{labels:{style:{colors:A,fontSize:'11px',fontFamily:F},formatter:function(v){return fmtC(v);}}},grid:{borderColor:G,strokeDashArray:4,padding:{top:-20,right:10,bottom:0,left:10}},legend:{position:'top',horizontalAlign:'right',labels:{colors:L},fontSize:'12px',fontFamily:F},tooltip:{theme:'dark',y:{formatter:function(v){return fmtBRL(v);}}},fill:{opacity:[.9,.7]},annotations:{yaxis:[{y:dados[0].teto,borderColor:'#DC143C',borderWidth:3,strokeDashArray:6,label:{text:'TETO 65%',style:{color:'#FFF',background:'#DC143C',fontSize:'10px',fontWeight:700}}}]}});charts.teto_barras.render();}
function renderTetoUnidades(registros){var unidades=[...new Set(registros.map(function(r){return r.unidade;}))];var dadosUnidades=unidades.map(function(u){var rows=registros.filter(function(r){return r.unidade===u;});var total=rows.reduce(function(s,r){return s+r.total_previsto;},0);var teto=rows.reduce(function(s,r){return s+getTeto(r);},0);var pct=total/teto*100;return{u:u,total:total,teto:teto,pct:pct};}).sort(function(a,b){return b.pct-a.pct;});var container=document.getElementById('teto-barras-unidades');container.innerHTML='';dadosUnidades.forEach(function(d){var cor=d.pct>=100?'red':(d.pct>=81?'orange':'green');var largura=Math.min(d.pct,100);var corTexto=d.pct>=100?'#FCA5A5':(d.pct>=81?'#FDBA74':'#6EE7B7');var row=document.createElement('div');row.className='teto-row';row.innerHTML='<div class="teto-unit-name">'+d.u+'</div><div class="teto-bar-container"><div class="teto-bar-wrap"><div class="teto-bar-fill '+cor+'" style="width:'+largura+'%"><span class="teto-bar-label">'+d.pct.toFixed(1)+'%</span></div></div></div><div class="teto-value" style="color:'+corTexto+'">'+fmtBRL(d.total)+'</div>';container.appendChild(row);});}
function renderTetoEvol(dados){var meses=dados.map(function(d){return d.mes_label;});var realizado=dados.map(function(d){return d.realizado;});var previsto=dados.map(function(d){return d.previsto;});var total=dados.map(function(d){return d.total_previsto;});var teto=dados.map(function(d){return d.teto;});charts.teto_evol=new ApexCharts(document.querySelector('#c-teto-evol'),{chart:{type:'bar',height:320,background:'transparent',toolbar:{show:false},fontFamily:F,animations:{enabled:true,speed:600}},plotOptions:{bar:{borderRadius:4,columnWidth:'55%',stacked:true}},series:[{name:'Realizado',data:realizado,color:'#0EA5E9'},{name:'Previsto',data:previsto,color:'#8B5CF6'},{name:'Total Previsto',data:total,type:'line',color:'#F97316'},{name:'Teto (65%)',data:teto,type:'line',color:'#DC143C'}],dataLabels:{enabled:true,enabledOnSeries:[2],formatter:function(v){return fmtC(v);},style:{colors:['#FDBA74'],fontSize:'10px',fontFamily:F,fontWeight:600},offsetY:-8},xaxis:{categories:meses,labels:{style:{colors:A,fontSize:'11px',fontFamily:F}},axisBorder:{show:false},axisTicks:{show:false}},yaxis:{labels:{style:{colors:A,fontSize:'11px',fontFamily:F},formatter:function(v){return fmtC(v);}}},grid:{borderColor:G,strokeDashArray:4,padding:{top:-20,right:10,bottom:0,left:10}},legend:{position:'top',horizontalAlign:'right',labels:{colors:L},fontSize:'12px',fontFamily:F},tooltip:{theme:'dark',y:{formatter:function(v){return fmtBRL(v);}}},stroke:{width:[0,0,3,3],curve:'smooth',dashArray:[0,0,0,6]},markers:{size:[0,0,5,0],colors:['#0EA5E9','#8B5CF6','#F97316','#DC143C'],strokeColors:'#0B1120',strokeWidth:2},fill:{opacity:[.9,.7,1,1]},colors:['#0EA5E9','#8B5CF6','#F97316','#DC143C']});charts.teto_evol.render();}
function renderTetoAccordion(registros){var container=document.getElementById('teto-accordion');container.innerHTML='';var unidades=[...new Set(registros.map(function(r){return r.unidade;}))].sort();unidades.forEach(function(unidade,idx){var rows=registros.filter(function(r){return r.unidade===unidade;});var ordemMeses=FIN_DATA.meses_ordem;rows.sort(function(a,b){return a.ano!==b.ano?a.ano-b.ano:ordemMeses.indexOf(a.mes)-ordemMeses.indexOf(b.mes);});var totalReal=rows.reduce(function(s,r){return s+r.realizado;},0);var totalPrev=rows.reduce(function(s,r){return s+r.previsto;},0);var totalGeral=rows.reduce(function(s,r){return s+r.total_previsto;},0);var tetoTotal=rows.reduce(function(s,r){return s+getTeto(r);},0);var pctGeral=totalGeral/tetoTotal*100;var corGeral=pctGeral>=100?'red':(pctGeral>=81?'orange':'green');var corBgGeral=pctGeral>=100?'rgba(220,20,60,0.12)':(pctGeral>=81?'rgba(249,115,22,0.12)':'rgba(16,185,129,0.12)');var corTextoGeral=pctGeral>=100?'#FCA5A5':(pctGeral>=81?'#FDBA74':'#6EE7B7');var item=document.createElement('div');item.className='accordion-item';if(idx===0)item.classList.add('open');var header=document.createElement('div');header.className='accordion-header';header.onclick=function(){item.classList.toggle('open');};header.innerHTML='<div class="accordion-header-left"><span class="accordion-arrow">▶</span><span class="accordion-unit-name">'+unidade+'</span><span class="accordion-unit-badge" style="background:'+corBgGeral+';color:'+corTextoGeral+'">'+pctGeral.toFixed(1)+'% do teto</span></div><div style="font-size:.78rem;font-weight:700;color:'+corTextoGeral+'">'+fmtBRL(totalGeral)+'</div>';item.appendChild(header);var body=document.createElement('div');body.className='accordion-body';var tableHtml='<table class="audit-table w-full"><thead><tr>';tableHtml+='<th>Mês/Ano</th><th>Faturamento</th><th>Teto (65%)</th><th>Realizado</th><th>Previsto</th><th>Total Previsto</th><th>% do Teto</th><th>Status</th>';tableHtml+='</tr></thead><tbody>';rows.forEach(function(r){var teto=getTeto(r);var pct=getPct(r);var fat=getFat(r);var corClass=pct>=100?'td-neg':(pct>=81?'td-warn':'td-pos');var status=pct>=100?'🔴 ESTOURADO':(pct>=81?'🟡 ATENÇÃO':'🟢 OK');var bgStyle=pct>=100?'background-color:rgba(220,20,60,0.08)':'';tableHtml+='<tr><td style="text-align:left">'+r.mes+'/'+String(r.ano).slice(-2)+'</td><td>'+fmtBRL(fat)+'</td><td>'+fmtBRL(teto)+'</td><td>'+fmtBRL(r.realizado)+'</td><td>'+fmtBRL(r.previsto)+'</td><td>'+fmtBRL(r.total_previsto)+'</td><td class="'+corClass+'" style="'+bgStyle+'">'+pct.toFixed(2).replace('.',',')+'%</td><td style="text-align:center;font-size:.68rem">'+status+'</td></tr>';});tableHtml+='<tr style="border-top:2px solid rgba(51,65,85,.5)"><td style="text-align:left;font-weight:800;color:#F1F5F9">TOTAL</td><td style="font-weight:700">—</td><td style="font-weight:700">'+fmtBRL(tetoTotal)+'</td><td style="font-weight:700">'+fmtBRL(totalReal)+'</td><td style="font-weight:700">'+fmtBRL(totalPrev)+'</td><td style="font-weight:700">'+fmtBRL(totalGeral)+'</td><td class="'+(corGeral==='red'?'td-neg':(corGeral==='orange'?'td-warn':'td-pos'))+'" style="font-weight:800">'+pctGeral.toFixed(2).replace('.',',')+'%</td><td style="text-align:center;font-size:.68rem">'+(pctGeral>=100?'🔴 ESTOURADO':(pctGeral>=81?'🟡 ATENÇÃO':'🟢 OK'))+'</td></tr>';tableHtml+='</tbody></table>';body.innerHTML=tableHtml;item.appendChild(body);container.appendChild(item);});}

// ==============================================================================
// ===== ABA 3: FLUXO OPERACIONAL — v6.9 COM DETALHAMENTO =====
// ==============================================================================

function setFluxoView(view){
    fluxoView=view;
    ['resumo','contas','fluxo','dps','inad'].forEach(function(v){
        var el=document.getElementById('view-'+v);
        if(el)el.classList.toggle('active',view===v);
    });
    document.getElementById('sec-resumo').classList.toggle('active',view==='resumo');
    document.getElementById('sec-contas').classList.toggle('active',view==='contas');
    document.getElementById('sec-fluxo-caixa').classList.toggle('active',view==='fluxo');
    document.getElementById('sec-dps').classList.toggle('active',view==='dps');
    document.getElementById('sec-inad').classList.toggle('active',view==='inad');
}

function renderFluxo(){
    if(!FLUXO_DATA||!FLUXO_DATA.fluxo_caixa)return;
    Object.keys(charts).forEach(function(k){if(k.startsWith('fluxo_')){charts[k].destroy();delete charts[k];}});
    renderFluxoKPIs();
    renderFluxoAlerta();
    renderResumo();
    renderFluxoBalanco();
    renderFluxoContas();
    renderContasTabela();
    renderFluxoInad();
    renderFluxoDPS();
    renderFluxoTabela();
}

// ===== RESUMO PLANILHA — v6.9 COM DETALHAMENTO =====
function renderResumo(){
    var container=document.getElementById('resumo-container');
    if(!container)return;
    var html='';

    // Bloco 1: Contas a Pagar — DETALHADO v6.9
    html+='<div class="resumo-block-title">💸 Contas a Pagar Semanal — Composição Detalhada</div>';
    html+='<table class="audit-table w-full"><thead><tr>';
    html+='<th>Unidade</th><th>Prev. Diversar</th><th>Prev. ICMS</th><th>Prev. FGTS/INSS</th><th>Prev. Comissão/Bônus</th><th>DUP. DDA</th><th>Total Geral</th>';
    html+='</tr></thead><tbody>';
    FLUXO_DATA.contas_pagar.forEach(function(r){
        var isTotal=r.unidade==='Total';
        var style=isTotal?' style="border-top:2px solid rgba(51,65,85,.5);font-weight:800"':'';
        var tdStyle=isTotal?' style="color:#F1F5F9"':'';
        html+='<tr'+style+'><td'+tdStyle+'>'+r.unidade+'</td>';
        html+='<td>'+fmtBRL(r.previsto_diversar)+'</td>';
        html+='<td>'+fmtBRL(r.previsto_icms)+'</td>';
        html+='<td>'+fmtBRL(r.previsto_fgts_inss)+'</td>';
        html+='<td>'+fmtBRL(r.previsto_comissao)+'</td>';
        html+='<td>'+fmtBRL(r.dup_dda)+'</td>';
        html+='<td style="font-weight:700">'+fmtBRL(r.total_geral)+'</td>';
        html+='</tr>';
    });
    html+='</tbody></table>';

    // Bloco 2: Média Diária
    html+='<div class="resumo-block-title">📊 Média Diária de Pagamento</div>';
    html+='<table class="audit-table w-full"><tbody>';
    html+='<tr><td>Dias Úteis</td><td>'+FLUXO_DATA.dias_uteis+'</td></tr>';
    html+='<tr><td>Total a Pagar</td><td>'+fmtBRL(FLUXO_DATA.saldo_a_pagar)+'</td></tr>';
    html+='<tr style="border-top:2px solid rgba(51,65,85,.5)"><td style="font-weight:800;color:#F1F5F9">Média Diária</td><td style="font-weight:800;color:#7DD3FC">'+fmtBRL(FLUXO_DATA.media_diaria)+'</td></tr>';
    html+='</tbody></table>';
    html+='<div class="resumo-obs">Observações: Para cálculo da média de pagamentos diários, considera-se '+FLUXO_DATA.dias_uteis+' dias úteis. As informações estão sujeitas a variações, pois ainda passarão pelo processo de conciliação entre o WinThor e o banco.</div>';

    // Bloco 3: Fluxo de Caixa
    html+='<div class="resumo-block-title">💰 Fluxo de Caixa Semanal</div>';
    html+='<table class="audit-table w-full"><thead><tr><th>Unidade</th><th>Saldo Inicial de Caixa</th><th>Previsão de Recebimento</th><th>Disponibilidade Total</th><th>Previsão de Contas a Pagar</th><th>Saldo Final</th></tr></thead><tbody>';
    FLUXO_DATA.fluxo_caixa.forEach(function(r){
        var isConsol=r.unidade==='CONSOLIDADO';
        var neg=r.saldo_final<0;
        html+='<tr'+(isConsol?' style="border-top:2px solid rgba(51,65,85,.5);font-weight:800"':'')+'><td'+(isConsol?' style="color:#F1F5F9"':'')+'>'+r.unidade+'</td><td>'+fmtBRL(r.saldo_inicial)+'</td><td>'+fmtBRL(r.previsao_recebimento)+'</td><td>'+fmtBRL(r.disponibilidade)+'</td><td>'+fmtBRL(r.contas_pagar)+'</td><td class="'+(neg?'td-neg':'td-pos')+'">'+fmtBRL(r.saldo_final)+'</td></tr>';
    });
    html+='</tbody></table>';

    // Bloco 4: DPS
    if(FLUXO_DATA.dps.length>0){
        html+='<div class="resumo-block-title">📝 Demonstrativo de Previsão de Saldo (DPS)</div>';
        html+='<table class="audit-table w-full"><thead><tr><th>Descrição</th><th>Valor</th></tr></thead><tbody>';
        FLUXO_DATA.dps.forEach(function(d){
            var isNeg=d.valor<0;
            var bold=d.label.startsWith('(=)');
            html+='<tr'+(bold?' style="border-top:1px solid rgba(51,65,85,.3);font-weight:800"':'')+'><td style="font-weight:'+(bold?'700':'400')+'">'+d.label+'</td><td class="'+(isNeg?'td-neg':(bold?'td-pos':''))+'" style="font-weight:'+(bold?'800':'600')+'">'+fmtBRL(d.valor)+'</td></tr>';
        });
        html+='</tbody></table>';
    }

    // Bloco 5: Inadimplência
    if(FLUXO_DATA.inadimplencia.length>0){
        html+='<div class="resumo-block-title">🔴 Diagnóstico de Inadimplência</div>';
        html+='<table class="audit-table w-full"><thead><tr><th>Unidade</th><th>Valor em Aberto</th></tr></thead><tbody>';
        var acumulado=null;
        FLUXO_DATA.inadimplencia.forEach(function(r){
            if(r.unidade==='ACUMULADO'){acumulado=r;return;}
            var isConsol=r.unidade.includes('CONSOLIDADO');
            html+='<tr'+(isConsol?' style="border-top:1px solid rgba(51,65,85,.3);font-weight:700"':'')+'><td'+(isConsol?' style="color:#F1F5F9"':'')+'>'+r.unidade+'</td><td class="td-neg">'+fmtBRL(r.valor)+'</td></tr>';
        });
        if(acumulado){
            html+='<tr style="border-top:2px solid rgba(220,20,60,0.3);font-weight:800"><td style="color:#FCA5A5">'+acumulado.unidade+'</td><td class="td-neg" style="font-size:.88rem">'+fmtBRL(acumulado.valor)+'</td></tr>';
        }
        html+='</tbody></table>';
    }

    container.innerHTML=html;
}

function renderFluxoKPIs(){
    var fc=FLUXO_DATA.fluxo_caixa;
    var consol=fc.find(function(r){return r.unidade==='CONSOLIDADO';})||fc[fc.length-1];
    if(!consol)return;
    var disp=consol.disponibilidade;var pagar=consol.contas_pagar;var saldo=consol.saldo_final;
    var media=FLUXO_DATA.media_diaria;var dias=FLUXO_DATA.dias_uteis;
    document.getElementById('fk-disp').textContent=fmtBRL(disp);
    document.getElementById('fk-disp').className='kpi-value mb-1 pos';
    document.getElementById('fk-disp-sub').textContent='Saldo Inicial + Recebimentos';
    document.getElementById('fk-pagar').textContent=fmtBRL(pagar);
    document.getElementById('fk-pagar').className='kpi-value mb-1 neg';
    document.getElementById('fk-pagar-sub').textContent='Total Geral a Pagar';
    document.getElementById('fk-media').textContent=fmtBRL(media);
    document.getElementById('fk-media').className='kpi-value mb-1';
    document.getElementById('fk-media-sub').textContent=dias+' dias úteis';
    var elSaldo=document.getElementById('fk-saldo');
    elSaldo.textContent=fmtBRL(saldo);
    elSaldo.className='kpi-value mb-1 '+(saldo<0?'neg':'pos');
    var subSaldo=document.getElementById('fk-saldo-sub');
    if(saldo<0){subSaldo.innerHTML='<span style="color:#FCA5A5;font-weight:700">⚠️ RISCO DE CAIXA</span>';}
    else{subSaldo.innerHTML='<span style="color:#6EE7B7">✅ Caixa Positivo</span>';}
}

function renderFluxoAlerta(){
    var fc=FLUXO_DATA.fluxo_caixa;
    var alertas=fc.filter(function(r){return r.saldo_final<0 && r.unidade!=='CONSOLIDADO';});
    var consol=fc.find(function(r){return r.unidade==='CONSOLIDADO';});
    var container=document.getElementById('fluxo-alerta');
    if(alertas.length===0 && (!consol||consol.saldo_final>=0)){container.innerHTML='';return;}
    var html='<div class="alert-box"><div style="font-size:1.2rem">🚨</div><div><div style="font-size:.88rem;font-weight:700;color:#FCA5A5">ALERTA DE CAIXA — Saldo Final Negativo</div><div style="font-size:.74rem;color:#CBD5E1;margin-top:4px">';
    alertas.forEach(function(a){html+=a.unidade+': '+fmtBRL(a.saldo_final)+' · ';});
    if(consol&&consol.saldo_final<0)html+='CONSOLIDADO: '+fmtBRL(consol.saldo_final);
    html+='</div></div></div>';
    container.innerHTML=html;
}

function renderFluxoBalanco(){
    var fc=FLUXO_DATA.fluxo_caixa.filter(function(r){return r.unidade!=='CONSOLIDADO';});
    if(fc.length===0)return;
    var nomes=fc.map(function(r){return r.unidade;});
    var disp=fc.map(function(r){return r.disponibilidade;});
    var pagar=fc.map(function(r){return r.contas_pagar;});
    var saldo=fc.map(function(r){return r.saldo_final;});
    charts.fluxo_balanco=new ApexCharts(document.querySelector('#c-fluxo-balanco'),{
        chart:{type:'bar',height:380,background:'transparent',toolbar:{show:false},fontFamily:F,animations:{enabled:true,speed:600}},
        plotOptions:{bar:{borderRadius:4,columnWidth:'50%'}},
        series:[{name:'Disponibilidade Total',data:disp,color:'#0EA5E9'},{name:'Contas a Pagar',data:pagar,color:'#F97316'},{name:'Saldo Final',data:saldo,type:'line',color:'#DC143C'}],
        dataLabels:{enabled:true,enabledOnSeries:[2],formatter:function(v){return fmtC(v);},style:{colors:['#FCA5A5'],fontSize:'10px',fontFamily:F,fontWeight:600},offsetY:-8},
        xaxis:{categories:nomes,labels:{style:{colors:A,fontSize:'11px',fontFamily:F}},axisBorder:{show:false},axisTicks:{show:false}},
        yaxis:{labels:{style:{colors:A,fontSize:'11px',fontFamily:F},formatter:function(v){return fmtC(v);}}},
        grid:{borderColor:G,strokeDashArray:4,padding:{top:-20,right:10,bottom:0,left:10}},
        legend:{position:'top',horizontalAlign:'right',labels:{colors:L},fontSize:'12px',fontFamily:F},
        tooltip:{theme:'dark',y:{formatter:function(v){return fmtBRL(v);}}},
        stroke:{width:[0,0,3],curve:'smooth'},
        markers:{size:[0,0,5],colors:['#0EA5E9','#F97316','#DC143C'],strokeColors:'#0B1120',strokeWidth:2},
        fill:{opacity:[.9,.8,1]},colors:['#0EA5E9','#F97316','#DC143C']
    });
    charts.fluxo_balanco.render();
}

// ===== GRÁFICO CONTAS A PAGAR — EMPILHADO v6.9 =====
function renderFluxoContas(){
    var cp=FLUXO_DATA.contas_pagar.filter(function(r){return r.unidade!=='Total';});
    if(cp.length===0)return;
    var nomes=cp.map(function(r){return r.unidade;});
    var diversar=cp.map(function(r){return r.previsto_diversar;});
    var icms=cp.map(function(r){return r.previsto_icms;});
    var fgts=cp.map(function(r){return r.previsto_fgts_inss;});
    var comissao=cp.map(function(r){return r.previsto_comissao;});
    var dda=cp.map(function(r){return r.dup_dda;});

    charts.fluxo_contas=new ApexCharts(document.querySelector('#c-fluxo-contas'),{
        chart:{type:'bar',height:380,background:'transparent',toolbar:{show:false},fontFamily:F,animations:{enabled:true,speed:600}},
        plotOptions:{bar:{borderRadius:4,columnWidth:'55%',stacked:true}},
        series:[
            {name:'Prev. Diversar',data:diversar,color:'#8B5CF6'},
            {name:'Prev. ICMS',data:icms,color:'#0EA5E9'},
            {name:'Prev. FGTS/INSS',data:fgts,color:'#10B981'},
            {name:'Prev. Comissão/Bônus',data:comissao,color:'#F97316'},
            {name:'DUP. DDA',data:dda,color:'#EF4444'}
        ],
        dataLabels:{enabled:false},
        xaxis:{categories:nomes,labels:{style:{colors:A,fontSize:'11px',fontFamily:F}},axisBorder:{show:false},axisTicks:{show:false}},
        yaxis:{labels:{style:{colors:A,fontSize:'11px',fontFamily:F},formatter:function(v){return fmtC(v);}}},
        grid:{borderColor:G,strokeDashArray:4,padding:{top:-20,right:10,bottom:0,left:10}},
        legend:{position:'top',horizontalAlign:'right',labels:{colors:L},fontSize:'11px',fontFamily:F},
        tooltip:{theme:'dark',y:{formatter:function(v){return fmtBRL(v);}}},
        fill:{opacity:[.7,.7,.7,.7,.7]}
    });
    charts.fluxo_contas.render();
}

// ===== TABELA CONTAS A PAGAR — DETALHADA v6.9 =====
function renderContasTabela(){
    var cp=FLUXO_DATA.contas_pagar;
    if(cp.length===0)return;
    var th=document.getElementById('th-contas');
    th.innerHTML='<th>Unidade</th><th>Prev. Diversar</th><th>Prev. ICMS</th><th>Prev. FGTS/INSS</th><th>Prev. Comissão/Bônus</th><th>DUP. DDA</th><th>Total Geral</th>';
    var tb=document.getElementById('tb-contas');tb.innerHTML='';
    cp.forEach(function(r){
        var isTotal=r.unidade==='Total';
        var style=isTotal?' style="border-top:2px solid rgba(51,65,85,.5);font-weight:800"':'';
        var tdStyle=isTotal?' style="color:#F1F5F9"':'';
        var html='<tr'+style+'><td'+tdStyle+'>'+r.unidade+'</td>';
        html+='<td>'+fmtBRL(r.previsto_diversar)+'</td>';
        html+='<td>'+fmtBRL(r.previsto_icms)+'</td>';
        html+='<td>'+fmtBRL(r.previsto_fgts_inss)+'</td>';
        html+='<td>'+fmtBRL(r.previsto_comissao)+'</td>';
        html+='<td>'+fmtBRL(r.dup_dda)+'</td>';
        html+='<td style="font-weight:700">'+fmtBRL(r.total_geral)+'</td>';
        html+='</tr>';
        tb.innerHTML+=html;
    });
}

function renderFluxoInad(){
    var inad=FLUXO_DATA.inadimplencia.filter(function(r){return r.unidade!=='ACUMULADO' && !r.unidade.includes('CONSOLIDADO');});
    if(inad.length===0)return;
    inad.sort(function(a,b){return b.valor-a.valor;});
    var nomes=inad.map(function(r){return r.unidade;});
    var valores=inad.map(function(r){return r.valor;});
    var acumulado=FLUXO_DATA.inadimplencia.find(function(r){return r.unidade==='ACUMULADO';});
    charts.fluxo_inad=new ApexCharts(document.querySelector('#c-fluxo-inad'),{
        chart:{type:'bar',height:300,background:'transparent',toolbar:{show:false},fontFamily:F,animations:{enabled:true,speed:600}},
        plotOptions:{bar:{horizontal:true,borderRadius:6,barHeight:'55%',distributed:true,dataLabels:{position:'center'}}},
        series:[{name:'Inadimplência',data:valores}],
        dataLabels:{enabled:true,formatter:function(v){return fmtBRL0(v);},style:{colors:['#FFF'],fontSize:'11px',fontFamily:F,fontWeight:700}},
        xaxis:{categories:nomes,labels:{style:{colors:A},formatter:function(v){return fmtC(v);}}},
        yaxis:{labels:{style:{colors:'#F1F5F9',fontWeight:600,fontSize:'12px'}}},
        grid:{borderColor:G,strokeDashArray:4},legend:{show:false},
        tooltip:{theme:'dark',y:{formatter:function(v){return fmtBRL(v);}}},
        colors:['#DC143C','#EF4444','#F87171','#FCA5A5','#FECACA'],
        fill:{colors:['#DC143C','#EF4444','#F87171','#FCA5A5','#FECACA']}
    });
    charts.fluxo_inad.render();
    if(acumulado){
        var container=document.querySelector('#c-fluxo-inad');
        var existing=container.parentNode.querySelector('.inad-acumulado-info');
        if(existing)existing.remove();
        var info=document.createElement('div');
        info.style.cssText='text-align:center;padding:12px;margin-top:8px;background:rgba(220,20,60,0.08);border-radius:8px;';
        info.innerHTML='<span style="font-size:.72rem;color:#94A3B8">ACUMULADO HISTÓRICO: </span><span style="font-size:.88rem;font-weight:700;color:#FCA5A5">'+fmtBRL(acumulado.valor)+'</span>';
        info.className='inad-acumulado-info';
        container.parentNode.appendChild(info);
    }
}

function renderFluxoDPS(){
    var dps=FLUXO_DATA.dps;
    if(dps.length===0)return;
    var container=document.getElementById('fluxo-dps');
    var html='<div style="background:rgba(15,23,42,.4);border:1px solid rgba(51,65,85,.3);border-radius:10px;padding:16px">';
    dps.forEach(function(d){
        var isNegativo=d.valor<0;
        var cor=isNegativo?'#FCA5A5':'#E2E8F0';
        var bold=d.label.startsWith('(=)');
        html+='<div class="dps-row">';
        html+='<div class="dps-label" style="font-weight:'+(bold?'700':'400')+'">'+d.label+'</div>';
        html+='<div class="dps-value" style="color:'+cor+';font-weight:'+(bold?'800':'600')+'">'+fmtBRL(d.valor)+'</div>';
        html+='</div>';
    });
    html+='</div>';
    container.innerHTML=html;
}

function renderFluxoTabela(){
    var fc=FLUXO_DATA.fluxo_caixa;
    if(fc.length===0)return;
    var th=document.getElementById('th-fluxo');
    th.innerHTML='<th>Unidade</th><th>Saldo Inicial</th><th>Previsão Receb.</th><th>Disponibilidade</th><th>Contas a Pagar</th><th>Saldo Final</th><th>Status</th>';
    var tb=document.getElementById('tb-fluxo');tb.innerHTML='';
    fc.forEach(function(r){
        var neg=r.saldo_final<0;
        var status=neg?'🔴 NEGATIVO':'🟢 POSITIVO';
        var corClass=neg?'td-neg':'td-pos';
        var isConsol=r.unidade==='CONSOLIDADO';
        var html='<tr'+(isConsol?' style="border-top:2px solid rgba(51,65,85,.5);font-weight:800"':'')+'><td'+(isConsol?' style="color:#F1F5F9"':'')+'>'+r.unidade+'</td><td>'+fmtBRL(r.saldo_inicial)+'</td><td>'+fmtBRL(r.previsao_recebimento)+'</td><td>'+fmtBRL(r.disponibilidade)+'</td><td>'+fmtBRL(r.contas_pagar)+'</td><td class="'+corClass+'">'+fmtBRL(r.saldo_final)+'</td><td style="text-align:center;font-size:.68rem">'+status+'</td></tr>';
        tb.innerHTML+=html;
    });
}
</script>
</body>
</html>'''

# ==============================================================================
# EXECUÇÃO PRINCIPAL
# ==============================================================================

def gerar_dashboard():
    print("=" * 60)
    print("ENGIPEC — SISTEMA DE INTELIGÊNCIA CORPORATIVA v6.9")
    print("=" * 60)

    print("\n[1/6] Carregando logo...")
    logo_src = carregar_logo_base64()
    if logo_src: print("  ✓ Logo carregada")
    else: print("  ⚠ Logo não encontrada")

    print("\n[2/6] Processando dados financeiros (Aba 1)...")
    data_fin = processar_financeiro()
    print(f"  ✓ {len(data_fin['registros'])} registros financeiros")

    print("\n[3/6] Processando teto orçamentário (Aba 2)...")
    data_teto = processar_teto(data_fin)
    print(f"  ✓ {len(data_teto['registros'])} registros de teto")

    print("\n[4/6] Processando fluxo operacional (Aba 3)...")
    data_fluxo = processar_fluxo()
    print(f"  ✓ {len(data_fluxo['fluxo_caixa'])} registros de fluxo")

    print("\n[5/6] Gerando HTML unificado com 3 abas...")
    html = gerar_html(data_fin, data_teto, data_fluxo, logo_src)
    Path(ARQ_HTML).write_text(html, encoding='utf-8')
    print(f"  ✓ {ARQ_HTML}")

    print("\n[6/6] Abrindo navegador...")
    webbrowser.open(Path(ARQ_HTML).resolve().as_uri())
    print("  ✓ Aberto")

    print("\n" + "=" * 60)
    print("PLATAFORMA UNIFICADA GERADA — v6.9")
    print("  Aba 1: Painel Financeiro Global")
    print("  Aba 2: Controle de Teto Orçamentário (OTB)")
    print("  Aba 3: Fluxo Operacional — Projeção Semanal")
    print("  v6.9: Contas a Pagar detalhado com 5 componentes")
    print("=" * 60)

if __name__ == '__main__':
    gerar_dashboard()
