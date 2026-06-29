"""
AGERSA - Sistema de Extração e Relatório
=========================================
Interface grafica completa usando Playwright.
Execute: python agersa_sistema.py

Dependencias:
  playwright, pandas, openpyxl
"""
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import threading
import os, shutil, math
from pathlib import Path
from datetime import datetime
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import time
import json

# =====================================================
#  CONFIGURACOES FIXAS
# =====================================================
URL_LOGIN       = "https://siafe2-flexvision.fazenda.rj.gov.br/Flexvision/"
CODIGO_CONSULTA = "084308"
TITULO_CONSULTA = "MODELO - RELATORIO CONSULTA - NE - NL - PD - CONTRATO - RPP - RPNP Copia"
UG_FIXA         = "043500"
PASTA_SAIDA     = str(Path.home() / "Downloads" / "AGERSA")

MESES = {
    "1 - Janeiro": "1", "2 - Fevereiro": "2", "3 - Marco": "3",
    "4 - Abril": "4",   "5 - Maio": "5",       "6 - Junho": "6",
    "7 - Julho": "7",   "8 - Agosto": "8",     "9 - Setembro": "9",
    "10 - Outubro":"10","11 - Novembro":"11",   "12 - Dezembro":"12",
}

COR_HEADER  = "1F497D"; COR_HEADER2 = "2E75B6"; COR_TOTAL = "D6E4F0"
COR_ALERTA = "FCE4D6"; COR_OK = "E2EFDA"; FONTE_XL = "Arial"

COLS_VALOR = [
    "Despesas Empenhadas", "Desp. Emp. a Liquidar", "Despesas Liquidadas",
    "Desp. Liq. a Pagar", "Despesas a Pagar", "Despesas Pagas",
    "Restos a Pagar Nao Processados a Liquidar",
    "Restos a Pagar Nao Processados Inscritos",
    "Restos a Pagar Nao Processados Inscritos - Exercicios Anteriores",
    "Restos a Pagar Nao Processados Liquidados a Pagar",
    "Restos a Pagar Processados - Pagos", "Restos a Pagar Processados a Pagar",
    "Restos a Pagar Processados Inscritos", "RP Pagos",
    "Restos a Pagar Nao Processados Pagos", "Retencoes Pagas", "PD PAGA RETENCAO",
]

# =====================================================
#  FUNCOES EXCEL
# =====================================================
def borda():
    s = Side(style="thin", color="BFBFBF")
    return Border(left=s, right=s, top=s, bottom=s)

def fmt_header(cell, cor=COR_HEADER):
    cell.font = Font(name=FONTE_XL, bold=True, color="FFFFFF", size=9)
    cell.fill = PatternFill("solid", fgColor=cor)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border = borda()

def fmt_valor(cell, cor_bg=None):
    cell.number_format = '#,##0.00'
    cell.alignment = Alignment(horizontal="right", vertical="center")
    cell.border = borda()
    cell.font = Font(name=FONTE_XL, size=9)
    if cor_bg:
        cell.fill = PatternFill("solid", fgColor=cor_bg)

def fmt_texto(cell, bold=False):
    cell.alignment = Alignment(horizontal="left", vertical="center", wrap_text=False)
    cell.border = borda()
    cell.font = Font(name=FONTE_XL, size=9, bold=bold)

def ajustar_col(ws, min_w=8, max_w=45):
    for col in ws.columns:
        cl = get_column_letter(col[0].column)
        mx = max((len(str(c.value or "")) for c in col), default=min_w)
        ws.column_dimensions[cl].width = max(min_w, min(mx + 2, max_w))

def titulo_aba(ws, titulo, subtitulo, n_cols=20):
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    c = ws.cell(row=1, column=1, value=titulo)
    c.font = Font(name=FONTE_XL, bold=True, size=13, color="FFFFFF")
    c.fill = PatternFill("solid", fgColor=COR_HEADER)
    c.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 28
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    c2 = ws.cell(row=2, column=1, value=subtitulo)
    c2.font = Font(name=FONTE_XL, italic=True, size=9, color="595959")
    c2.alignment = Alignment(horizontal="center")
    ws.row_dimensions[2].height = 14
    return 3

def preencher_aba(wb, nome_aba, titulo, df, cols_dim, cols_val_subset=None, params=None):
    ws = wb.create_sheet(nome_aba)
    p = params or {}
    periodo = ""
    if p.get('mes_ini_nome') and p.get('exercicio'):
        periodo = f" | Periodo: {p['mes_ini_nome']}/{p['exercicio']} a {p.get('mes_nome','')}/{p['exercicio']}"
    sub = f"UG: {UG_FIXA} - AGERSA{periodo} | Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    dim_ok = [c for c in cols_dim if c in df.columns]
    val_ok = [c for c in (cols_val_subset or COLS_VALOR) if c in df.columns]
    cols = dim_ok + val_ok
    if not cols:
        ws['A1'] = "Sem dados"; return ws
    df2 = df[cols].copy()
    n = len(cols)
    start = titulo_aba(ws, titulo, sub, n)
    val_idx = [i+1 for i, c in enumerate(cols) if c in val_ok]
    for i, c in enumerate(cols, 1):
        fmt_header(ws.cell(row=start, column=i, value=c), cor=COR_HEADER2 if c in val_ok else COR_HEADER)
    ws.row_dimensions[start].height = 40
    for ri, row in enumerate(df2.itertuples(index=False), start+1):
        for ci, val in enumerate(row, 1):
            if isinstance(val, float) and math.isnan(val): val = None
            cell = ws.cell(row=ri, column=ci, value=val)
            fmt_valor(cell) if ci in val_idx else fmt_texto(cell)
    tr = start + len(df2) + 1
    ws.cell(row=tr, column=1, value="TOTAL")
    for c in val_idx:
        cl = get_column_letter(c)
        ws.cell(row=tr, column=c, value=f"=SUM({cl}{start+1}:{cl}{start+len(df2)})")
    for c in range(1, n+1):
        cell = ws.cell(row=tr, column=c)
        cell.font = Font(name=FONTE_XL, bold=True, size=9)
        cell.fill = PatternFill("solid", fgColor=COR_TOTAL)
        cell.border = borda()
    ws.freeze_panes = ws.cell(row=start+1, column=1)
    ws.auto_filter.ref = f"A{start}:{get_column_letter(n)}{start}"
    ws.sheet_view.showGridLines = False
    ajustar_col(ws)
    return ws

def gerar_resumo(wb, df, params=None):
    ws = wb.create_sheet("RESUMO", 0)
    p = params or {}
    periodo = ""
    if p.get('mes_ini_nome') and p.get('exercicio'):
        periodo = f" | Periodo: {p['mes_ini_nome']}/{p['exercicio']} a {p.get('mes_nome','')}/{p['exercicio']}"
    sub = f"UG: {UG_FIXA} - AGERSA{periodo} | Gerado em: {datetime.now().strftime('%d/%m/%Y %H:%M')}"
    start = titulo_aba(ws, "RESUMO EXECUTIVO - EXECUCAO ORCAMENTARIA AGERSA", sub, 4)
    def bloco(tit, row):
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=4)
        fmt_header(ws.cell(row=row, column=1, value=tit), cor=COR_HEADER2)
        return row + 1
    def linha(nome, valor, row, cor=None, bold=False, pct=None):
        c1 = ws.cell(row=row, column=1, value=nome)
        c2 = ws.cell(row=row, column=2, value=valor)
        c1.font = Font(name=FONTE_XL, bold=bold, size=10); c1.border = borda()
        c2.number_format = '#,##0.00'; c2.alignment = Alignment(horizontal="right")
        c2.border = borda(); c2.font = Font(name=FONTE_XL, bold=bold, size=10)
        if pct:
            c3 = ws.cell(row=row, column=3, value=f"=B{row}/B{pct}")
            c3.number_format = '0.0%'; c3.border = borda(); c3.font = Font(name=FONTE_XL, size=10)
        if cor:
            for c in [c1, c2]:
                c.fill = PatternFill("solid", fgColor=cor)
            if pct: ws.cell(row=row, column=3).fill = PatternFill("solid", fgColor=cor)
        return row + 1
    r = start
    r = bloco("CICLO DA DESPESA - EXERCICIO CORRENTE", r)
    emp_row = None
    for nome, cor, bold in [
        ("Despesas Empenhadas", "F2F2F2", True),
        ("Desp. Emp. a Liquidar", COR_ALERTA, False),
        ("Despesas Liquidadas", "F2F2F2", False),
        ("Desp. Liq. a Pagar", COR_ALERTA, False),
        ("Despesas Pagas", COR_OK, True),
        ("Despesas a Pagar", COR_ALERTA, False),
    ]:
        if nome in df.columns:
            val = df[nome].sum()
            r = linha(nome, val, r, cor=cor, bold=bold, pct=emp_row)
            if nome == "Despesas Empenhadas": emp_row = r - 1
    r += 1
    r = bloco("RESTOS A PAGAR", r)
    for nome in ["Restos a Pagar Processados Inscritos", "Restos a Pagar Processados a Pagar",
                 "Restos a Pagar Processados - Pagos", "Restos a Pagar Nao Processados Inscritos",
                 "Restos a Pagar Nao Processados Pagos", "RP Pagos", "Retencoes Pagas", "PD PAGA RETENCAO"]:
        if nome in df.columns and df[nome].sum() != 0:
            r = linha(nome, df[nome].sum(), r)
    r += 1
    r = bloco("CONTADORES", r)
    for label, col in [
        ("Contratos distintos", "Contrato SIAFE"), ("Credores distintos", "Credor"),
        ("Notas de Empenho", "Nota de Empenho"), ("Notas de Liquidacao", "Nota de Liquidaacao"),
        ("Programacoes de Desembolso", "Programacao de Desembolso"),
        ("Ordens Bancarias", "Ordem Bancaria"),
    ]:
        if col in df.columns:
            c1 = ws.cell(row=r, column=1, value=label)
            c2 = ws.cell(row=r, column=2, value=int(df[col].nunique()))
            c1.font = Font(name=FONTE_XL, size=10); c1.border = borda()
            c2.font = Font(name=FONTE_XL, size=10, bold=True); c2.border = borda()
            c2.alignment = Alignment(horizontal="right"); r += 1
    ajustar_col(ws, min_w=10, max_w=55)
    ws.column_dimensions['B'].width = 20
    ws.column_dimensions['C'].width = 14
    ws.sheet_view.showGridLines = False
    return ws

def processar_csv(csv_path, params, log):
    log("Lendo CSV...")
    df = None
    for enc in ['utf-8-sig', 'latin-1', 'cp1252']:
        for sep in [';', ',', '\t']:
            try:
                df = pd.read_csv(csv_path, encoding=enc, sep=sep, dtype=str)
                if len(df.columns) > 5:
                    log(f"CSV: {len(df)} linhas, {len(df.columns)} colunas (enc={enc})"); break
            except: continue
        if df is not None and len(df.columns) > 5: break
    if df is None or df.empty:
        raise ValueError("Nao foi possivel ler o CSV.")
    df.columns = df.columns.str.strip()
    df.rename(columns={"Contratoa SIAFE": "Contrato SIAFE"}, inplace=True)
    df.dropna(how='all', inplace=True)
    for col in df.columns:
        if any(v in col for v in ["Despesa","Restos","Retenc","RP ","PD PAGA","Liquidad","Pagar","Pagas","Empenhad"]):
            df[col] = (df[col].astype(str).str.replace(r'\s','',regex=True)
                       .str.replace('.','',regex=False).str.replace(',','.',regex=False))
            df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
    log("Gerando Excel...")
    Path(PASTA_SAIDA).mkdir(parents=True, exist_ok=True)
    mes_ini = params.get('mes_ini','01').zfill(2)
    mes_fmt = params['mes'].zfill(2)
    nome = f"AGERSA_Execucao_{params['exercicio']}_{mes_ini}a{mes_fmt}_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
    caminho = str(Path(PASTA_SAIDA) / nome)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    log("  > Aba RESUMO"); gerar_resumo(wb, df, params)
    log("  > Aba CONSOLIDADO"); preencher_aba(wb, "CONSOLIDADO", "CONSOLIDADO - EXECUCAO COMPLETA", df,
        [c for c in df.columns if c not in COLS_VALOR], params=params)
    log("  > Aba NE"); preencher_aba(wb, "NE - Empenhos", "NOTAS DE EMPENHO", df,
        ["Unidade Gestora","Credor","Nota de Empenho","Historico","Processo (NE)","Contrato SIAFE","Contrato ORIGINAL"],
        ["Despesas Empenhadas","Desp. Emp. a Liquidar","Despesas Liquidadas","Despesas Pagas"], params=params)
    log("  > Aba NL"); preencher_aba(wb, "NL - Liquidacoes", "NOTAS DE LIQUIDACAO", df,
        ["Unidade Gestora","Credor","Nota de Empenho","Nota de Liquidacao","NF","Contrato SIAFE"],
        ["Despesas Liquidadas","Desp. Liq. a Pagar","Despesas Pagas"], params=params)
    log("  > Aba PD"); preencher_aba(wb, "PD - Desembolso", "PROGRAMACAO DE DESEMBOLSO", df,
        ["Unidade Gestora","Credor","Nota de Empenho","Nota de Liquidacao","Programacao de Desembolso","Contrato SIAFE"],
        ["Despesas a Pagar","Despesas Pagas","Desp. Liq. a Pagar",
         "Restos a Pagar Processados a Pagar","Restos a Pagar Processados - Pagos"], params=params)
    log("  > Aba OB"); preencher_aba(wb, "OB - Ordens Bancarias", "ORDENS BANCARIAS", df,
        ["Unidade Gestora","Credor","Nota de Empenho","Nota de Liquidacao","Programacao de Desembolso","Ordem Bancaria","Contrato SIAFE","Credor da Retencao"],
        ["Despesas Pagas","RP Pagos","Restos a Pagar Nao Processados Pagos","Retencoes Pagas","PD PAGA RETENCAO"], params=params)
    wb.save(caminho)
    log(f"Excel salvo: {caminho}")
    return caminho

# =====================================================
#  AUTOMACAO FLEXVISION - PLAYWRIGHT
# =====================================================

def flexvision_playwright(params, credenciais, log):
    """
    Executa fluxo completo no FlexVision usando Playwright.
    Consulta 084308: Exercicio (dropdown Vaadin), Mes (dropdown Vaadin), UG (input).
    """
    from playwright.sync_api import sync_playwright

    Path(PASTA_SAIDA).mkdir(parents=True, exist_ok=True)
    arquivo_baixado = None

    def scr(page, nome):
        try: page.screenshot(path=str(Path(PASTA_SAIDA) / f"debug_{nome}.png"))
        except: pass

    def aguardar_vaadin(page, segs=2):
        try: page.wait_for_selector(".v-loading-indicator", state="hidden", timeout=20000)
        except: pass
        if segs > 0: time.sleep(segs)

    def selecionar_dropdown_vaadin(page, indice_combo, valor_texto, log_label):
        """
        Clica no dropdown Vaadin (indice 0=Exercicio, 1=Mes),
        aguarda lista abrir e clica na opcao com o texto correspondente.
        """
        try:
            # Clica no botao seta do combo (v-filterselect-button)
            botoes = page.locator(".v-filterselect-button").all()
            if len(botoes) <= indice_combo:
                log(f"  [AVISO] Dropdown {log_label}: botao nao encontrado (total={len(botoes)})")
                return False

            botoes[indice_combo].click()
            time.sleep(1)

            # Aguarda lista de opcoes aparecer
            page.wait_for_selector(".v-filterselect-suggestpopup", timeout=5000)
            time.sleep(0.5)

            # Clica na opcao que contem o texto desejado
            opcoes = page.locator(".v-filterselect-suggestpopup td.gwt-MenuItem").all()
            for op in opcoes:
                txt = (op.text_content() or "").strip()
                if valor_texto in txt or txt == valor_texto:
                    op.click()
                    log(f"  {log_label}: {txt}")
                    time.sleep(0.5)
                    return True

            # Se nao achou, loga opcoes disponiveis para debug
            textos = [(op.text_content() or "").strip() for op in opcoes]
            log(f"  [AVISO] {log_label}: '{valor_texto}' nao encontrado. Opcoes: {textos[:5]}")

            # Tenta clicar na primeira opcao que contenha o numero do mes/ano
            for op in opcoes:
                txt = (op.text_content() or "").strip()
                if str(valor_texto) in txt:
                    op.click()
                    log(f"  {log_label} (parcial): {txt}")
                    time.sleep(0.5)
                    return True

            # Fecha dropdown sem selecionar
            page.keyboard.press("Escape")
            return False

        except Exception as e:
            log(f"  [AVISO] Dropdown {log_label}: {e}")
            try: page.keyboard.press("Escape")
            except: pass
            return False

    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            headless=False,
            args=["--start-maximized"]
        )
        ctx = browser.new_context(
            ignore_https_errors=True,
            no_viewport=True,  # deixa o browser maximizar sem viewport fixo
            locale="pt-BR"
        )
        page = ctx.new_page()
        # Força maximização real da janela
        page.set_default_timeout(60000)

        try:
            # ── 1. LOGIN ─────────────────────────────────────────────
            log("Abrindo FlexVision...")
            page.goto(URL_LOGIN, wait_until="load", timeout=60000)
            page.wait_for_selector("input[type='text']", timeout=30000)
            time.sleep(2)
            scr(page, "01_login")

            log("Preenchendo credenciais...")
            page.fill("input[type='text']", credenciais['usuario'])
            page.fill("input[type='password']", credenciais['senha'])
            time.sleep(1)
            page.click(".v-button-friendly")
            aguardar_vaadin(page, 5)
            scr(page, "02_pos_login")
            log(f"  Login OK.")

            # ── 2. MENU CONSULTAS ─────────────────────────────────────
            log("Clicando em Consultas...")
            page.evaluate("""
                () => {
                    const items = document.querySelectorAll('.valo-menu-item[role="button"]');
                    for (const el of items) {
                        const cap = el.querySelector('.valo-menu-item-caption');
                        if (cap && cap.textContent.trim() === 'Consultas') {
                            el.click(); return;
                        }
                    }
                }
            """)
            aguardar_vaadin(page, 4)
            scr(page, "03_consultas")

            # ── 3. FILTRAR CONSULTA ───────────────────────────────────
            log(f"Filtrando {CODIGO_CONSULTA}...")
            for tentativa in range(3):
                try:
                    filtro = page.locator("input[placeholder*='filtr']").first
                    filtro.click()
                    time.sleep(0.3)
                    filtro.fill("")
                    filtro.fill(CODIGO_CONSULTA)
                    aguardar_vaadin(page, 3)
                    break
                except Exception as e:
                    log(f"  Tentativa {tentativa+1}: {e}")
                    time.sleep(2)
            scr(page, "04_filtro")

            # ── 4. SELECIONAR LINHA ───────────────────────────────────
            log(f"Selecionando linha {CODIGO_CONSULTA}...")
            linha_el = page.get_by_text(CODIGO_CONSULTA, exact=True).first
            linha_el.click(timeout=15000)
            time.sleep(2)
            scr(page, "05_linha")

            # Aguarda toolbar ter pelo menos 2 botoes habilitados apos clear-button
            # Isso confirma que a linha esta selecionada
            for espera in range(15):
                n_hab = page.evaluate(
                    "() => { "
                    "  var btns = Array.from(document.querySelectorAll('.v-button')); "
                    "  var clearBtn = btns.find(function(b){ return b.className.includes('clear-button'); }); "
                    "  if (!clearBtn) return 0; "
                    "  var cl = clearBtn.getBoundingClientRect().left; "
                    "  return btns.filter(function(b){ "
                    "    var r = b.getBoundingClientRect(); "
                    "    return r.left > cl && r.width > 0 "
                    "        && !b.className.includes('v-disabled') "
                    "        && !b.className.includes('clear-button'); "
                    "  }).length; "
                    "}"
                )
                log(f"  Botoes toolbar habilitados: {n_hab}")
                if n_hab >= 2:
                    break
                time.sleep(1)
                linha_el.click(timeout=5000)
                time.sleep(1)

            log("  Linha selecionada.")

            # ── 5. EXECUTAR CONSULTA ─────────────────────────────────
            # Fluxo manual: clique simples na linha → clique no icone engrenagem
            # O icone engrenagem fica na toolbar do PAINEL DIREITO apos selecionar a linha

            log("Clicando Executar (icone engrenagem)...")

            # Screenshot antes — para ver posicao dos botoes com linha selecionada
            scr(page, "05b_antes_executar")
            time.sleep(1)

            # Loga TODOS os botoes visiveis para identificar o correto
            todos = page.evaluate(
                "() => { "
                "  var btns = Array.from(document.querySelectorAll('.v-button')); "
                "  var info = []; "
                "  btns.forEach(function(b,i) { "
                "    var r = b.getBoundingClientRect(); "
                "    if (r.width > 0 && r.height > 0 && r.top < 120 && r.top > 0) { "
                "      info.push(i + ':left=' + Math.round(r.left) + ',top=' + Math.round(r.top) "
                "        + ',dis=' + b.className.includes('v-disabled') "
                "        + ',cls=' + b.className.substring(0,60)); "
                "    } "
                "  }); "
                "  return info.join(' | '); "
                "}"
            )
            log(f"  Botoes visiveis: {todos}")

            # Estrategia: pega botoes icon-only NAO disabled, ordena por left
            # O executar fica na toolbar do painel direito (left > 500)
            # Tenta todos os botoes habilitados da direita para a esquerda
            clicou = False
            for tentativa in range(5):
                result = page.evaluate(
                    "() => { "
                    "  var btns = Array.from(document.querySelectorAll('.v-button')); "
                    "  var clearBtn = null; "
                    "  var allBtns = Array.from(document.querySelectorAll('.v-button')); "
                    "  for (var b of allBtns) { "
                    "    if (b.className.includes('clear-button')) { clearBtn = b; break; } "
                    "  } "
                    "  if (!clearBtn) return 'nao:sem_clear_button'; "
                    "  var clearLeft = clearBtn.getBoundingClientRect().left; "
                    "  var direita = allBtns.filter(function(b) { "
                    "    var r = b.getBoundingClientRect(); "
                    "    return r.width > 0 && r.height > 0 "
                    "        && r.top > 0 && r.top < 120 "
                    "        && r.left > clearLeft "
                    "        && !b.className.includes('v-disabled') "
                    "        && !b.className.includes('clear-button'); "
                    "  }); "
                    "  direita.sort(function(a,b){ "
                    "    return a.getBoundingClientRect().left - b.getBoundingClientRect().left; "
                    "  }); "
                    "  var lefts = direita.map(function(b){ return Math.round(b.getBoundingClientRect().left); }); "
                    "  var idx = direita.length - 1; "
                    "  if (idx >= 0) { "
                    "    direita[idx].click(); "
                    "    return 'ok:left=' + Math.round(direita[idx].getBoundingClientRect().left) + '|clearLeft=' + Math.round(clearLeft) + '|todos=' + JSON.stringify(lefts); "
                    "  } "
                    "  return 'nao:clearLeft=' + Math.round(clearLeft) + '|lefts=' + JSON.stringify(lefts); "
                    "}"
                )
                log(f"  Clique tentativa {tentativa+1}: {result}")
                if result.startswith('ok'):
                    clicou = True
                    break
                time.sleep(2)

            scr(page, "06_apos_executar")

            if not clicou:
                raise RuntimeError("Botao Executar nao encontrado.")

            # ── 6. AGUARDAR MODAL ─────────────────────────────────────
            log("Aguardando modal de parametros...")
            modal_ok = False
            for i in range(30):
                time.sleep(1)
                detectado = page.evaluate("""
                    () => {
                        // Verifica botoes de dropdown Vaadin visiveis (v-filterselect-button)
                        const combos = document.querySelectorAll('.v-filterselect-button');
                        let visiveis = 0;
                        for (const c of combos) {
                            const r = c.getBoundingClientRect();
                            if (r.width > 0 && r.height > 0) visiveis++;
                        }
                        if (visiveis >= 2) return 'combos_visiveis';
                        // Fallback: texto do modal
                        const body = document.body.textContent || '';
                        if (body.includes('Informe o Exerc') || body.includes('Parametros da consulta'))
                            return 'texto_encontrado';
                        return false;
                    }
                """)
                if detectado:
                    modal_ok = True
                    log(f"  Modal detectado: {detectado}")
                    break
                if i % 5 == 0:
                    log(f"  Aguardando modal... ({i+1}s)")

            if not modal_ok:
                scr(page, "07_modal_nao_detectado")
                raise RuntimeError("Modal de parametros nao apareceu apos 30s.")

            scr(page, "07_parametros")
            # Aguarda modal estar completamente carregado
            time.sleep(3)
            # Verifica se o modal real abriu (deve ter selects ou dropdowns)
            tem_selects = page.evaluate(
                "() => document.querySelectorAll('select').length"
            )
            log(f"  Selects no modal: {tem_selects}")

            # ── 7. PREENCHER PARÂMETROS ───────────────────────────────
            # Modal tem: 2 dropdowns Vaadin + 1 input texto (UG)
            # Os dropdowns NAO usam .v-filterselect-button — usam div clicavel com seta
            log("Preenchendo parametros...")
            time.sleep(2)
            scr(page, "07b_modal_html")

            # Loga estrutura do modal para debug
            estrutura = page.evaluate(
                "() => { "
                "  var modal = document.querySelector('.v-window, .v-overlay-container, [class*=window], [class*=dialog]'); "
                "  if (!modal) modal = document.body; "
                "  var els = modal.querySelectorAll('input, select, .v-select, .v-combobox, [class*=combo], [class*=select], button, .v-button'); "
                "  var info = []; "
                "  els.forEach(function(el) { "
                "    var r = el.getBoundingClientRect(); "
                "    if (r.width > 0) info.push(el.tagName + '.' + (el.className||'').substring(0,40) + ':left=' + Math.round(r.left)); "
                "  }); "
                "  return info.join(' | '); "
                "}"
            )
            log(f"  Estrutura modal: {estrutura[:300]}")

            # Dropdowns Vaadin filterselect
            # Estrategia: clicar no botao seta (v-filterselect-button) para abrir
            # e depois clicar na opcao correta no popup

            def preencher_filterselect(indice, valor_texto, label):
                """
                Preenche dropdown Vaadin filterselect.
                Abre com seta, rola ate encontrar o valor e clica.
                """
                try:
                    # Abre o dropdown clicando na seta
                    btn_seta = page.locator(".v-filterselect-button").nth(indice)
                    btn_seta.click(timeout=5000)
                    time.sleep(1)

                    # Aguarda popup aparecer
                    popup = page.locator(".v-filterselect-suggestpopup")
                    try:
                        popup.wait_for(state="visible", timeout=4000)
                    except Exception:
                        pass
                    time.sleep(0.5)

                    # Tenta achar o valor nas opcoes visiveis primeiro
                    def achar_e_clicar():
                        opcoes = page.locator(".v-filterselect-suggestpopup td.gwt-MenuItem").all()
                        for op in opcoes:
                            txt = (op.text_content() or "").strip()
                            if valor_texto in txt:
                                op.click()
                                log(f"  {label} selecionado: {txt}")
                                return True
                        return False

                    if achar_e_clicar():
                        time.sleep(0.3)
                        return True

                    # Nao achou — rola via JS e verifica a cada passo
                    log(f"  {label}: rolando lista via JS para achar '{valor_texto}'...")
                    for scroll_step in range(30):
                        # Rola o container da lista via JavaScript
                        page.evaluate(
                            "() => { "
                            "  var popup = document.querySelector('.v-filterselect-suggestpopup'); "
                            "  if (popup) { "
                            "    var scroller = popup.querySelector('[class*=scroll], .popupContent, div'); "
                            "    if (scroller) scroller.scrollTop += 120; "
                            "    else popup.scrollTop += 120; "
                            "  } "
                            "}"
                        )
                        time.sleep(0.2)
                        if achar_e_clicar():
                            time.sleep(0.3)
                            return True

                    # Ultima tentativa: fecha e digita direto no input
                    log(f"  {label}: nao encontrado na lista — digitando direto")
                    page.keyboard.press("Escape")
                    time.sleep(0.3)
                    inp = page.locator("input.v-filterselect-input").nth(indice)
                    inp.click(timeout=3000)
                    inp.press("Control+a")
                    inp.type(valor_texto, delay=100)  # digita devagar
                    time.sleep(1)
                    # Tenta clicar na opcao filtrada
                    try:
                        page.locator(".v-filterselect-suggestpopup td.gwt-MenuItem").filter(
                            has_text=valor_texto).first.click(timeout=3000)
                        log(f"  {label} (digitado+clique): {valor_texto}")
                    except Exception:
                        page.keyboard.press("Enter")
                        log(f"  {label} (digitado+enter): {valor_texto}")
                    time.sleep(0.5)
                    return True

                except Exception as e:
                    log(f"  [AVISO] {label}: {str(e)[:150]}")
                    return False

            # Conta filterselect buttons disponiveis
            n_fs = len(page.locator(".v-filterselect-button").all())
            log(f"  Filterselect buttons: {n_fs}")

            # Exercicio (indice 0)
            preencher_filterselect(0, params['exercicio'], "Exercicio")
            time.sleep(1)

            # Mes (indice 1) — tenta pelo nome e pelo padrao "N - Nome"
            mes_nome = params['mes_nome']
            mes_num  = params['mes']
            if not preencher_filterselect(1, mes_nome, "Mes"):
                preencher_filterselect(1, f"{mes_num} -", "Mes (num)")
            time.sleep(1)

            # Preenche UG — terceiro input visivel (apos os 2 filterselects)
            # O campo UG e um v-textfield simples, nao filterselect
            try:
                # Busca inputs que NAO sejam filterselect
                inputs_all = page.locator("input").all()
                preencheu = False
                for inp in inputs_all:
                    try:
                        r = inp.bounding_box()
                        if not r or r['width'] == 0: continue
                        cls = inp.get_attribute("class") or ""
                        typ = inp.get_attribute("type") or "text"
                        if typ in ["hidden","checkbox","radio"]: continue
                        # Pula os filterselect inputs (Exercicio e Mes)
                        if "filterselect" in cls: continue
                        inp.click(timeout=3000)
                        inp.fill(params['ug'])
                        log(f"  UG: {params['ug']} (cls={cls[:40]})")
                        preencheu = True
                        break
                    except Exception:
                        continue
                if not preencheu:
                    # Fallback: usa o terceiro input da pagina
                    page.locator("input").nth(2).fill(params['ug'])
                    log(f"  UG (nth 2): {params['ug']}")
            except Exception as e:
                log(f"  [AVISO] UG: {str(e)[:100]}")

            time.sleep(1)
            scr(page, "09_param_preenchidos")

            # ── 8. CLICAR OK ──────────────────────────────────────────
            log("Clicando OK...")
            ok_clicado = False
            # Tenta por role=button com texto Ok
            try:
                page.get_by_role("button", name="Ok").click(timeout=5000)
                ok_clicado = True
                log("  OK clicado (get_by_role)")
            except Exception:
                pass

            # Fallback: evaluate
            if not ok_clicado:
                for tentativa in range(10):
                    result = page.evaluate(
                        "() => { "
                        "  var all = document.querySelectorAll('.v-button, button, [role=\"button\"]'); "
                        "  for (var b of all) { "
                        "    var r = b.getBoundingClientRect(); "
                        "    if (r.width === 0 || r.height === 0) continue; "
                        "    var txt = (b.textContent || '').replace(/[^a-zA-Z]/g,'').toLowerCase(); "
                        "    if (txt === 'ok') { b.click(); return 'ok:' + Math.round(r.left); } "
                        "  } "
                        "  return 'nao'; "
                        "}"
                    )
                    log(f"  OK tentativa {tentativa+1}: {result}")
                    if result.startswith('ok'):
                        ok_clicado = True
                        break
                    time.sleep(1)

            if not ok_clicado:
                scr(page, "10_ok_nao_encontrado")
                raise RuntimeError("Botao OK nao encontrado.")

            scr(page, "10_ok_clicado")
            aguardar_vaadin(page, 3)

            # ── 8b. SEGUNDO MODAL: "Como prefere executar?" ───────────
            log("Verificando segundo modal...")
            for i in range(10):
                time.sleep(1)
                tem_segundo = page.evaluate(
                    "() => { var b = document.body.textContent || ''; "
                    "return b.includes('Como prefere executar') || b.includes('Dados processados'); }"
                )
                if tem_segundo:
                    log("  Segundo modal detectado — selecionando 'mais rapido' e clicando OK...")
                    scr(page, "10b_segundo_modal")
                    # Seleciona primeiro radio (Dados processados - mais rapido)
                    page.evaluate(
                        "() => { var inputs = document.querySelectorAll('input'); "
                        "for (var i of inputs) { if (i.type === 'radio') { i.click(); break; } } }"
                    )
                    time.sleep(0.5)
                    # Clica OK
                    try:
                        page.get_by_role('button', name='Ok').click(timeout=3000)
                    except Exception:
                        page.evaluate(
                            "() => { var all = document.querySelectorAll('.v-button, button'); "
                            "for (var b of all) { "
                            "  var txt = (b.textContent||'').replace(/[^a-zA-Z]/g,'').toLowerCase(); "
                            "  var r = b.getBoundingClientRect(); "
                            "  if (r.width > 0 && txt === 'ok') { b.click(); return; } "
                            "} }"
                        )
                    log("  OK do segundo modal clicado!")
                    aguardar_vaadin(page, 3)
                    break
                if i == 9:
                    log("  Segundo modal nao detectado — continuando.")

            # ── 9. AGUARDAR RESULTADO ─────────────────────────────────────────
            log("Aguardando resultado...")
            resultado_ok = False
            for i in range(150):  # ate 300s = 5 minutos
                time.sleep(2)
                try:
                    texto = page.evaluate("() => document.body.innerText")
                except Exception as e:
                    log(f"  [AVISO] Erro ao ler pagina ({i*2}s): {str(e)[:80]}")
                    time.sleep(3)
                    continue
                carregando = any(p in texto for p in ['Carregando','carregando','Aguarde','Processando'])
                if carregando:
                    if i % 10 == 0: log(f"  Processando... ({i*2}s)")
                    continue
                tem_dados = (
                    len(texto) > 500 and
                    any(p in texto for p in ['Credor','Empenho','Despesa',UG_FIXA,'Liquidac','043500'])
                )
                if tem_dados:
                    resultado_ok = True
                    log("  Resultado carregado!")
                    break
                if i % 10 == 0:
                    log(f"  Aguardando dados... ({i*2}s)")

            scr(page, "11_resultado")
            if not resultado_ok:
                log("  [AVISO] Resultado nao detectado — tentando exportar mesmo assim...")
                resultado_ok = True

            # ── 10. EXPORTAR CSV ──────────────────────────────────────
            log("Exportando CSV...")

            def clicar_export():
                # Toolbar de resultados fica entre top=55 e top=90 (linha do filtro)
                # Icones: filtro, atualizar, engrenagem, impressora, DOWNLOAD, info
                # O download e o penultimo icone — identificado pela posicao
                return page.evaluate(
                    "() => { "
                    # Tenta pelo title em qualquer elemento pai
                    "  var els = document.querySelectorAll('.v-button, button, [role=button]'); "
                    "  for (var el of els) { "
                    "    var t = (el.title || el.getAttribute('aria-label') || ''); "
                    "    var p = el.parentElement; "
                    "    var pt = p ? (p.title || p.getAttribute('aria-label') || '') : ''; "
                    "    var txt = (t + ' ' + pt).toLowerCase(); "
                    "    if (txt.includes('export') || txt.includes('download') || txt.includes('baixar')) { "
                    "      var r = el.getBoundingClientRect(); "
                    "      if (r.width > 0) { el.click(); return 'ok_title:' + txt.substring(0,40); } "
                    "    } "
                    "  } "
                    # Toolbar de resultados: botoes na faixa top 55-90, left > 200, nao disabled
                    "  var btns = Array.from(document.querySelectorAll('.v-button')); "
                    "  var toolbar = btns.filter(function(b) { "
                    "    var r = b.getBoundingClientRect(); "
                    "    return r.width > 0 && r.width < 50 "
                    "        && r.top >= 55 && r.top <= 95 "
                    "        && r.left > 200 "
                    "        && !b.className.includes('v-disabled') "
                    "        && !b.className.includes('clear-button'); "
                    "  }); "
                    "  toolbar.sort(function(a,b){ return a.getBoundingClientRect().left - b.getBoundingClientRect().left; }); "
                    "  var lefts = toolbar.map(function(b){ return Math.round(b.getBoundingClientRect().left); }); "
                    # Penultimo = download (o ultimo e informacao)
                    "  if (toolbar.length >= 2) { "
                    "    var btn = toolbar[toolbar.length-2]; "
                    "    btn.click(); "
                    "    return 'ok_toolbar|left=' + Math.round(btn.getBoundingClientRect().left) + '|todos=' + JSON.stringify(lefts); "
                    "  } "
                    "  if (toolbar.length === 1) { toolbar[0].click(); return 'ok_1|' + JSON.stringify(lefts); } "
                    "  return 'nao|lefts=' + JSON.stringify(lefts); "
                    "}"
                )

            # Abre menu export ANTES do expect_download
            for tentativa in range(5):
                result = clicar_export()
                log(f"  Export tentativa {tentativa+1}: {result}")
                if result.startswith('ok'):
                    break
                time.sleep(2)
            aguardar_vaadin(page, 2)
            scr(page, "11b_export_menu")
            time.sleep(1)

            # Captura download enquanto clica em CSV
            with page.expect_download(timeout=120000) as dl_info:
                for tentativa in range(15):
                    try:
                        csv_el = page.get_by_text("Csv", exact=True).first
                        if csv_el.is_visible(timeout=1000):
                            csv_el.click()
                            log("  CSV clicado!")
                            break
                    except Exception:
                        pass
                    r = page.evaluate(
                        "() => {"
                        "  var w = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);"
                        "  var n; while (n = w.nextNode()) {"
                        "    if (n.textContent.trim() === 'Csv') {"
                        "      var e = n.parentElement; var r = e.getBoundingClientRect();"
                        "      if (r.width > 0) { e.click(); return 'ok'; }"
                        "    }"
                        "  } return 'nao';"
                        "}"
                    )
                    log(f"  CSV ev {tentativa+1}: {r}")
                    if r == 'ok':
                        log("  CSV selecionado!")
                        break
                    if tentativa > 0 and tentativa % 4 == 0:
                        log("  Reabrindo menu...")
                        clicar_export()
                        time.sleep(1)
                    time.sleep(1)
            try:
                download = dl_info.value
                arquivo_baixado = str(Path(PASTA_SAIDA) / download.suggested_filename)
                download.save_as(arquivo_baixado)
                log(f"  Download: {download.suggested_filename}")
            except Exception as e:
                raise RuntimeError(f"Download nao capturado: {e}")

            scr(page, "12_export")
            return arquivo_baixado

        except Exception as e:
            scr(page, "99_erro")
            raise

        finally:
            try: browser.close()
            except: pass


# =====================================================
#  INTERFACE GRAFICA
# =====================================================
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("AGERSA - Extracao FlexVision")
        self.resizable(False, False)
        self.configure(bg="#F0F4F8")
        self._build_ui()
        self._centralizar()

    def _centralizar(self):
        self.update_idletasks()
        w, h = self.winfo_width(), self.winfo_height()
        sw, sh = self.winfo_screenwidth(), self.winfo_screenheight()
        self.geometry(f"+{(sw-w)//2}+{(sh-h)//2}")

    def _build_ui(self):
        AZUL = "#1F497D"; AZUL2 = "#2E75B6"; BG = "#F0F4F8"; BG_CARD = "#FFFFFF"

        header = tk.Frame(self, bg=AZUL, pady=14)
        header.pack(fill="x")
        tk.Label(header, text="AGERSA", font=("Arial", 20, "bold"), fg="white", bg=AZUL).pack()
        tk.Label(header, text="Sistema de Extracao - FlexVision -> Excel",
                 font=("Arial", 10), fg="#BDD7EE", bg=AZUL).pack()

        body = tk.Frame(self, bg=BG, padx=20, pady=15)
        body.pack(fill="both")

        def card(parent, titulo):
            f = tk.LabelFrame(parent, text=f"  {titulo}  ",
                              font=("Arial", 9, "bold"), fg=AZUL2,
                              bg=BG_CARD, bd=1, relief="solid", padx=12, pady=10)
            f.pack(fill="x", pady=6)
            return f

        def row(parent, label, widget_fn, **kw):
            fr = tk.Frame(parent, bg=BG_CARD); fr.pack(fill="x", pady=3)
            tk.Label(fr, text=label, font=("Arial", 9), fg="#374151", bg=BG_CARD, width=22, anchor="w").pack(side="left")
            w = widget_fn(fr, **kw); w.pack(side="left", fill="x", expand=True)
            return w

        c1 = card(body, "Acesso FlexVision")
        self.v_usuario = tk.StringVar(); self.v_senha = tk.StringVar()
        row(c1, "Usuario:", lambda p: tk.Entry(p, textvariable=self.v_usuario, font=("Arial",10), relief="solid", bd=1))
        row(c1, "Senha:", lambda p: tk.Entry(p, textvariable=self.v_senha, show="o", font=("Arial",10), relief="solid", bd=1))

        c2 = card(body, "Periodo da Consulta")
        ano_atual = datetime.now().year
        mes_atual = datetime.now().month
        self.v_exercicio = tk.StringVar(value=str(ano_atual))
        row(c2, "Exercicio (Ano):", lambda p: tk.Entry(p, textvariable=self.v_exercicio, width=8, font=("Arial",10), relief="solid", bd=1))
        opcoes_mes = list(MESES.keys())
        self.v_mes_ini = tk.StringVar(value=opcoes_mes[0])
        row(c2, "Mes de Inicio:", lambda p: ttk.Combobox(p, textvariable=self.v_mes_ini, values=opcoes_mes, state="readonly", font=("Arial",10)))
        self.v_mes = tk.StringVar(value=opcoes_mes[mes_atual - 1])
        row(c2, "Mes de Fim:", lambda p: ttk.Combobox(p, textvariable=self.v_mes, values=opcoes_mes, state="readonly", font=("Arial",10)))

        c3 = card(body, "Filtros Opcionais (% = todos)")
        self.v_uo = tk.StringVar(value="%"); self.v_ne = tk.StringVar(value="%")
        self.v_nl = tk.StringVar(value="%"); self.v_cnpj = tk.StringVar(value="%")
        self.v_natureza = tk.StringVar(value="%"); self.v_contrato = tk.StringVar(value="%")
        ug_fr = tk.Frame(c3, bg=BG_CARD); ug_fr.pack(fill="x", pady=3)
        tk.Label(ug_fr, text="UG (fixo AGERSA):", font=("Arial",9), fg="#374151", bg=BG_CARD, width=22, anchor="w").pack(side="left")
        tk.Label(ug_fr, text=f"{UG_FIXA}  OK", font=("Arial",9,"bold"), fg=AZUL2, bg=BG_CARD).pack(side="left")
        row(c3, "UO:", lambda p: tk.Entry(p, textvariable=self.v_uo, font=("Arial",10), relief="solid", bd=1))
        row(c3, "NE:", lambda p: tk.Entry(p, textvariable=self.v_ne, font=("Arial",10), relief="solid", bd=1))
        row(c3, "NL:", lambda p: tk.Entry(p, textvariable=self.v_nl, font=("Arial",10), relief="solid", bd=1))
        row(c3, "CNPJ/CPF Credor:", lambda p: tk.Entry(p, textvariable=self.v_cnpj, font=("Arial",10), relief="solid", bd=1))
        row(c3, "Natureza:", lambda p: tk.Entry(p, textvariable=self.v_natureza, font=("Arial",10), relief="solid", bd=1))
        row(c3, "Contrato:", lambda p: tk.Entry(p, textvariable=self.v_contrato, font=("Arial",10), relief="solid", bd=1))

        self.btn = tk.Button(body, text=">  EXECUTAR EXTRACAO",
                             font=("Arial", 11, "bold"), fg="white", bg=AZUL2,
                             activebackground=AZUL, activeforeground="white",
                             relief="flat", padx=20, pady=10, cursor="hand2",
                             command=self._executar)
        self.btn.pack(fill="x", pady=(8, 4))

        tk.Label(body, text="Log de execucao:", font=("Arial", 8), fg="#6B7280", bg=BG).pack(anchor="w")
        self.log_box = scrolledtext.ScrolledText(body, height=10, font=("Consolas", 8),
            bg="#1E1E1E", fg="#D4D4D4", relief="flat", state="disabled", wrap="word")
        self.log_box.pack(fill="both", pady=(2, 0))

    def log(self, msg):
        ts = datetime.now().strftime("%H:%M:%S")
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"[{ts}] {msg}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")
        self.update_idletasks()

    def _executar(self):
        usuario = self.v_usuario.get().strip()
        senha = self.v_senha.get().strip()
        exercicio = self.v_exercicio.get().strip()
        mes_ini_str = self.v_mes_ini.get()
        mes_str = self.v_mes.get()

        if not usuario or not senha:
            messagebox.showwarning("Atencao", "Preencha usuario e senha."); return
        if not exercicio.isdigit():
            messagebox.showwarning("Atencao", "Exercicio invalido."); return

        mes_ini_num = int(MESES.get(mes_ini_str, "1"))
        mes_num = MESES.get(mes_str, "")
        mes_nome = mes_str.split(" - ")[-1] if " - " in mes_str else mes_str
        mes_ini_nome = mes_ini_str.split(" - ")[-1] if " - " in mes_ini_str else mes_ini_str

        if int(mes_num or 1) < mes_ini_num:
            messagebox.showwarning("Atencao", "Mes de Inicio deve ser <= Mes de Fim."); return

        params = {
            "exercicio": exercicio, "mes": mes_num, "mes_nome": mes_nome,
            "mes_ini": str(mes_ini_num), "mes_ini_nome": mes_ini_nome,
            "uo": self.v_uo.get().strip() or "%", "ug": UG_FIXA,
            "ne": self.v_ne.get().strip() or "%", "nl": self.v_nl.get().strip() or "%",
            "cnpj": self.v_cnpj.get().strip() or "%",
            "natureza": self.v_natureza.get().strip() or "%",
            "contrato": self.v_contrato.get().strip() or "%",
        }
        credenciais = {"usuario": usuario, "senha": senha}

        self.btn.configure(state="disabled", text=" Executando...")
        self.log("=" * 45)
        self.log(f"Iniciando extracao: {exercicio} / {mes_ini_nome} -> {mes_nome}")
        self.log("=" * 45)

        def run():
            try:
                arquivo = flexvision_playwright(params, credenciais, self.log)
                suffix = Path(arquivo).suffix.lower()
                if suffix in ('.xls', '.xlsx'):
                    nome_final = f"AGERSA_Execucao_{params['exercicio']}_{params['mes'].zfill(2)}_{datetime.now().strftime('%Y%m%d_%H%M')}{suffix}"
                    excel_path = str(Path(PASTA_SAIDA) / nome_final)
                    shutil.copy2(arquivo, excel_path)
                else:
                    excel_path = processar_csv(arquivo, params, self.log)
                self.log("=" * 45)
                self.log("CONCLUIDO")
                self.log(f"Arquivo: {Path(excel_path).name}")
                self.log("=" * 45)
                if sys.platform == "win32": os.startfile(excel_path)
                messagebox.showinfo("Concluido", f"Arquivo gerado!\n\n{Path(excel_path).name}\n\nPasta: {PASTA_SAIDA}")
            except Exception as e:
                self.log(f"[ERRO] {e}")
                import traceback
                self.log(traceback.format_exc())
                messagebox.showerror("Erro", str(e))
            finally:
                self.btn.configure(state="normal", text=">  EXECUTAR EXTRACAO")

        threading.Thread(target=run, daemon=True).start()

if __name__ == "__main__":
    App().mainloop()
