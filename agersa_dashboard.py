"""
AGERSA - Dashboard de Execução Orçamentária
=============================================
Execute: streamlit run agersa_dashboard.py

Instale: pip install streamlit plotly
"""

import streamlit as st
import sqlite3
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from pathlib import Path
from datetime import datetime

# ═══════════════════════════════════════════════════════
#  CONFIGURAÇÃO
# ═══════════════════════════════════════════════════════
BANCO_PATH = str(Path.home() / "Downloads" / "AGERSA" / "agersa_historico.db")

MESES = {
    1:"Janeiro", 2:"Fevereiro", 3:"Março", 4:"Abril",
    5:"Maio", 6:"Junho", 7:"Julho", 8:"Agosto",
    9:"Setembro", 10:"Outubro", 11:"Novembro", 12:"Dezembro"
}

COR_EMPENHADO  = "#1F497D"
COR_LIQUIDADO  = "#2E75B6"
COR_PAGO       = "#70AD47"
COR_A_PAGAR    = "#ED7D31"
COR_RP         = "#FF0000"

st.set_page_config(
    page_title="AGERSA — Dashboard Orçamentário",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# CSS customizado
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1F497D 0%, #2E75B6 100%);
        padding: 20px 30px;
        border-radius: 10px;
        margin-bottom: 20px;
        color: white;
    }
    .main-header h1 { color: white; margin: 0; font-size: 28px; }
    .main-header p  { color: #BDD7EE; margin: 5px 0 0 0; font-size: 14px; }

    .metric-card {
        background: white;
        border-radius: 10px;
        padding: 15px 20px;
        border-left: 4px solid #2E75B6;
        box-shadow: 0 2px 8px rgba(0,0,0,0.08);
    }
    .metric-value { font-size: 22px; font-weight: bold; color: #1F497D; }
    .metric-label { font-size: 12px; color: #666; margin-top: 4px; }
    .metric-delta { font-size: 12px; margin-top: 6px; }

    .alerta-box {
        background: #FFF3CD;
        border: 1px solid #FFCA28;
        border-radius: 8px;
        padding: 12px 16px;
        margin: 8px 0;
    }
    div[data-testid="stMetric"] {
        background: white;
        border-radius: 10px;
        padding: 15px;
        box-shadow: 0 2px 8px rgba(0,0,0,0.06);
    }
</style>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════
#  DADOS
# ═══════════════════════════════════════════════════════
@st.cache_data(ttl=300)
def carregar_dados(importacao_id=None):
    conn = sqlite3.connect(BANCO_PATH)

    importacoes = pd.read_sql(
        "SELECT * FROM vw_resumo_importacoes", conn)

    if importacao_id is None and not importacoes.empty:
        importacao_id = int(importacoes.iloc[0]['id'])

    filtro = f"WHERE importacao_id = {importacao_id}" if importacao_id else ""

    fato = pd.read_sql(f"""
        SELECT * FROM fato_execucao {filtro}
    """, conn)

    credores = pd.read_sql(f"""
        SELECT * FROM vw_top_credores {filtro}
        ORDER BY empenhado DESC
    """, conn)

    contratos = pd.read_sql(f"""
        SELECT * FROM vw_contratos {filtro}
        WHERE contrato_original != '00000000'
        ORDER BY empenhado DESC
    """, conn)

    conn.close()
    return importacoes, fato, credores, contratos, importacao_id

def formatar_moeda(valor):
    if valor >= 1_000_000:
        return f"R$ {valor/1_000_000:.1f}M"
    elif valor >= 1_000:
        return f"R$ {valor/1_000:.0f}K"
    return f"R$ {valor:,.0f}"

# ═══════════════════════════════════════════════════════
#  SIDEBAR
# ═══════════════════════════════════════════════════════
with st.sidebar:
    st.markdown("### ⚙️ Configurações")

    if not Path(BANCO_PATH).exists():
        st.error(f"Banco não encontrado:\n{BANCO_PATH}")
        st.stop()

    importacoes, fato, credores, contratos, imp_id = carregar_dados()

    if importacoes.empty:
        st.warning("Nenhuma importação encontrada.")
        st.info("Execute a extração primeiro:\n`python agersa_sistema.py`")
        st.stop()

    # Seletor de importação
    opcoes = {
        f"{r['periodo']} | {r['arquivo'][:20]}...": r['id']
        for _, r in importacoes.iterrows()
    }
    selecionado = st.selectbox(
        "Período:",
        list(opcoes.keys()),
        index=0
    )
    importacao_id = opcoes[selecionado]

    # Recarrega com importação selecionada
    _, fato, credores, contratos, _ = carregar_dados(importacao_id)

    imp_info = importacoes[importacoes['id'] == importacao_id].iloc[0]

    st.markdown("---")
    st.markdown(f"**Exercício:** {imp_info['exercicio']}")
    st.markdown(f"**Período:** {imp_info['periodo']}")
    st.markdown(f"**Registros:** {imp_info['total_linhas']:,}")
    st.markdown(f"**Atualizado:** {imp_info['data_importacao'][:16]}")

    st.markdown("---")
    if st.button("🔄 Atualizar dados"):
        st.cache_data.clear()
        st.rerun()

# ═══════════════════════════════════════════════════════
#  CABEÇALHO
# ═══════════════════════════════════════════════════════
st.markdown(f"""
<div class="main-header">
    <h1>📊 AGERSA — Dashboard Orçamentário</h1>
    <p>UG: 043500 | Exercício {imp_info['exercicio']} | Período: {imp_info['periodo']} |
    Atualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')}</p>
</div>
""", unsafe_allow_html=True)

# ═══════════════════════════════════════════════════════
#  MÉTRICAS PRINCIPAIS
# ═══════════════════════════════════════════════════════
empenhado  = float(imp_info['empenhado']  or 0)
liquidado  = float(importacoes[importacoes['id']==importacao_id]['liquidado'].iloc[0] or 0)
pago       = float(imp_info['pago']       or 0)
a_pagar    = float(imp_info['a_pagar']    or 0)
pct_exec   = float(imp_info['pct_execucao'] or 0)

rp_proc    = fato['rp_proc_inscritos'].sum()
rp_nao     = fato['rp_nao_proc_inscritos'].sum()

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.metric("💰 Empenhado",    formatar_moeda(empenhado))
with col2:
    st.metric("📋 Liquidado",    formatar_moeda(liquidado),
              f"{liquidado/empenhado*100:.1f}% do empenhado" if empenhado else "")
with col3:
    st.metric("✅ Pago",         formatar_moeda(pago),
              f"{pct_exec:.1f}% de execução")
with col4:
    st.metric("⏳ A Pagar",      formatar_moeda(a_pagar))
with col5:
    st.metric("📌 Restos a Pagar", formatar_moeda(rp_proc + rp_nao))

st.markdown("---")

# ═══════════════════════════════════════════════════════
#  LINHA 1: FUNIL + PIZZA CREDORES
# ═══════════════════════════════════════════════════════
col_a, col_b = st.columns([1, 1])

with col_a:
    st.subheader("🔽 Ciclo da Despesa")
    fig_funil = go.Figure(go.Funnel(
        y=["Empenhado", "Liquidado", "Pago"],
        x=[empenhado, liquidado, pago],
        textinfo="value+percent initial",
        texttemplate="%{label}<br>R$ %{value:,.0f}<br>(%{percentInitial})",
        marker=dict(color=[COR_EMPENHADO, COR_LIQUIDADO, COR_PAGO]),
    ))
    fig_funil.update_layout(
        margin=dict(l=10, r=10, t=10, b=10),
        height=300,
        showlegend=False
    )
    st.plotly_chart(fig_funil, use_container_width=True)

with col_b:
    st.subheader("🏢 Top 10 Credores por Empenhado")
    top10 = credores.head(10).copy()
    # Limpa nome do credor (remove CNPJ do início)
    top10['nome_credor'] = top10['credor'].str.replace(
        r'^\d{14}\s*-\s*', '', regex=True).str[:35]

    fig_bar = px.bar(
        top10[::-1],
        x='empenhado', y='nome_credor',
        orientation='h',
        color='pago',
        color_continuous_scale=[[0, COR_A_PAGAR], [1, COR_PAGO]],
        labels={'empenhado':'Empenhado (R$)', 'nome_credor':'', 'pago':'Pago (R$)'},
        text=top10[::-1]['empenhado'].apply(formatar_moeda)
    )
    fig_bar.update_layout(
        margin=dict(l=10, r=10, t=10, b=10),
        height=300,
        coloraxis_showscale=False,
        yaxis=dict(tickfont=dict(size=11))
    )
    fig_bar.update_traces(textposition='outside')
    st.plotly_chart(fig_bar, use_container_width=True)

# ═══════════════════════════════════════════════════════
#  LINHA 2: RESTOS A PAGAR + CONTRATOS
# ═══════════════════════════════════════════════════════
col_c, col_d = st.columns([1, 1])

with col_c:
    st.subheader("📌 Restos a Pagar")

    rp_dados = {
        'Tipo': ['RP Processado\nInscrito', 'RP Processado\na Pagar', 'RP Processado\nPago',
                 'RP Não Proc.\nInscrito', 'RP Não Proc.\nPago'],
        'Valor': [
            fato['rp_proc_inscritos'].sum(),
            fato['rp_proc_a_pagar'].sum(),
            fato['rp_proc_pagos'].sum(),
            fato['rp_nao_proc_inscritos'].sum(),
            fato['rp_nao_proc_pagos'].sum(),
        ]
    }
    df_rp = pd.DataFrame(rp_dados)
    df_rp = df_rp[df_rp['Valor'] > 0]

    if not df_rp.empty:
        fig_rp = px.bar(
            df_rp, x='Tipo', y='Valor',
            color='Valor',
            color_continuous_scale=[[0, '#FFF3CD'], [0.5, COR_A_PAGAR], [1, COR_RP]],
            text=df_rp['Valor'].apply(formatar_moeda)
        )
        fig_rp.update_layout(
            margin=dict(l=10, r=10, t=10, b=10),
            height=300,
            coloraxis_showscale=False,
            yaxis_title="R$",
            xaxis_title=""
        )
        fig_rp.update_traces(textposition='outside')
        st.plotly_chart(fig_rp, use_container_width=True)
    else:
        st.info("Sem restos a pagar neste período.")

with col_d:
    st.subheader("📋 Top Contratos")
    top_cont = contratos.head(10).copy()
    top_cont['contrato_label'] = top_cont['contrato_original'] + '\n' + \
        top_cont['credor'].str.replace(r'^\d{14}\s*-\s*','',regex=True).str[:25]

    if not top_cont.empty:
        fig_cont = px.bar(
            top_cont[::-1],
            x='empenhado', y='contrato_original',
            orientation='h',
            color='pago',
            color_continuous_scale=[[0, COR_A_PAGAR], [1, COR_PAGO]],
            labels={'empenhado':'Empenhado','contrato_original':'Contrato'},
            text=top_cont[::-1]['empenhado'].apply(formatar_moeda),
            hover_data=['credor', 'pago', 'a_pagar']
        )
        fig_cont.update_layout(
            margin=dict(l=10, r=10, t=10, b=10),
            height=300,
            coloraxis_showscale=False
        )
        fig_cont.update_traces(textposition='outside')
        st.plotly_chart(fig_cont, use_container_width=True)

# ═══════════════════════════════════════════════════════
#  ALERTAS AUTOMÁTICOS
# ═══════════════════════════════════════════════════════
st.markdown("---")
st.subheader("⚠️ Alertas Automáticos")

alertas = []

# Execução abaixo de 80%
if pct_exec < 80:
    alertas.append(f"🔴 Execução abaixo de 80%: {pct_exec:.1f}% pago do empenhado")

# Restos a pagar acima de 10% do empenhado
if (rp_proc + rp_nao) > empenhado * 0.10:
    pct_rp = (rp_proc + rp_nao) / empenhado * 100
    alertas.append(f"🟡 Restos a pagar elevados: {pct_rp:.1f}% do empenhado ({formatar_moeda(rp_proc + rp_nao)})")

# Credor com mais de 20% do total
if not credores.empty:
    maior_credor = credores.iloc[0]
    pct_credor = maior_credor['empenhado'] / empenhado * 100 if empenhado > 0 else 0
    if pct_credor > 20:
        nome = str(maior_credor['credor']).replace(r'^\d{14}\s*-\s*','')[:40]
        alertas.append(f"🟡 Concentração: {nome[:40]} representa {pct_credor:.1f}% do total empenhado")

# A pagar acima de 5% do empenhado
if a_pagar > empenhado * 0.05:
    alertas.append(f"🟠 Despesas a pagar elevadas: {formatar_moeda(a_pagar)} ({a_pagar/empenhado*100:.1f}% do empenhado)")

if alertas:
    cols_alerta = st.columns(min(len(alertas), 2))
    for i, alerta in enumerate(alertas):
        with cols_alerta[i % 2]:
            st.warning(alerta)
else:
    st.success("✅ Nenhum alerta identificado neste período.")

# ═══════════════════════════════════════════════════════
#  TABELAS DETALHADAS
# ═══════════════════════════════════════════════════════
st.markdown("---")
tab1, tab2, tab3 = st.tabs(["👥 Credores", "📋 Contratos", "🔍 Dados Brutos"])

with tab1:
    st.subheader(f"Todos os Credores ({len(credores)})")
    df_cred = credores.copy()
    df_cred['nome_credor'] = df_cred['credor'].str.replace(
        r'^\d{14}\s*-\s*', '', regex=True)
    df_cred['% do Total'] = (df_cred['empenhado'] / empenhado * 100).round(1)
    df_cred['Empenhado']  = df_cred['empenhado'].apply(lambda x: f"R$ {x:,.2f}")
    df_cred['Liquidado']  = df_cred['liquidado'].apply(lambda x: f"R$ {x:,.2f}")
    df_cred['Pago']       = df_cred['pago'].apply(lambda x: f"R$ {x:,.2f}")

    busca = st.text_input("🔍 Filtrar credor:", placeholder="Digite nome ou CNPJ...")
    df_show = df_cred
    if busca:
        df_show = df_cred[df_cred['credor'].str.contains(busca, case=False, na=False)]

    st.dataframe(
        df_show[['nome_credor','Empenhado','Liquidado','Pago','% do Total','qtd_empenhos']]\
            .rename(columns={'nome_credor':'Credor','qtd_empenhos':'NEs'}),
        use_container_width=True,
        height=400
    )

with tab2:
    st.subheader(f"Contratos ({len(contratos)})")
    df_cont = contratos.copy()
    df_cont['Credor'] = df_cont['credor'].str.replace(
        r'^\d{14}\s*-\s*', '', regex=True).str[:40]
    df_cont['Empenhado'] = df_cont['empenhado'].apply(lambda x: f"R$ {x:,.2f}")
    df_cont['Liquidado'] = df_cont['liquidado'].apply(lambda x: f"R$ {x:,.2f}")
    df_cont['Pago']      = df_cont['pago'].apply(lambda x: f"R$ {x:,.2f}")
    df_cont['A Pagar']   = df_cont['a_pagar'].apply(lambda x: f"R$ {x:,.2f}")

    busca_cont = st.text_input("🔍 Filtrar contrato:", placeholder="Número ou credor...")
    df_cont_show = df_cont
    if busca_cont:
        df_cont_show = df_cont[
            df_cont['contrato_original'].str.contains(busca_cont, case=False, na=False) |
            df_cont['credor'].str.contains(busca_cont, case=False, na=False)
        ]

    st.dataframe(
        df_cont_show[['contrato_original','Credor','Empenhado','Liquidado','Pago','A Pagar']]\
            .rename(columns={'contrato_original':'Contrato'}),
        use_container_width=True,
        height=400
    )

with tab3:
    st.subheader("Dados Brutos — CONSOLIDADO")
    busca_raw = st.text_input("🔍 Filtrar:", placeholder="Nota de empenho, credor, contrato...")

    colunas_exibir = ['nota_empenho','credor','contrato_siafe','contrato_original',
                      'nota_liquidacao','ordem_bancaria','despesas_empenhadas',
                      'despesas_liquidadas','despesas_pagas','despesas_a_pagar']
    df_raw = fato[colunas_exibir].copy()

    if busca_raw:
        mask = (
            fato['nota_empenho'].str.contains(busca_raw, case=False, na=False) |
            fato['credor'].str.contains(busca_raw, case=False, na=False) |
            fato['contrato_original'].str.contains(busca_raw, case=False, na=False)
        )
        df_raw = fato[mask][colunas_exibir]

    st.dataframe(df_raw, use_container_width=True, height=400)
    st.caption(f"{len(df_raw):,} registros exibidos")

# ═══════════════════════════════════════════════════════
#  RODAPÉ
# ═══════════════════════════════════════════════════════
st.markdown("---")
st.markdown(
    f"<small>AGERSA — Sistema de BI Orçamentário | "
    f"Banco: `{BANCO_PATH}` | "
    f"Atualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')}</small>",
    unsafe_allow_html=True
)
