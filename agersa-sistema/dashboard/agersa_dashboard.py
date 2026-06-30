"""
AGERSA - Dashboard Orçamentário
================================
Dashboard Streamlit conectado ao Supabase.
"""
import os
import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from pathlib import Path
from datetime import datetime

st.set_page_config(
    page_title="AGERSA — Dashboard",
    page_icon="📊",
    layout="wide"
)

# ── Conexão ──────────────────────────────────────────
DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

def get_conn():
    if DATABASE_URL:
        import psycopg2
        return psycopg2.connect(DATABASE_URL)
    raise Exception("DATABASE_URL não configurada")

# ── Carrega dados ─────────────────────────────────────
@st.cache_data(ttl=300)
def carregar():
    try:
        conn = get_conn()
        imp = pd.read_sql("""
            SELECT id, arquivo, exercicio,
                   mes_inicio || ' a ' || mes_fim AS periodo,
                   empenhado, liquidado, pago, a_pagar,
                   ROUND(pago::numeric / NULLIF(empenhado,0) * 100, 1) AS pct_execucao,
                   total_linhas
            FROM importacoes
            ORDER BY data_importacao DESC
        """, conn)

        if imp.empty:
            conn.close()
            return imp, pd.DataFrame(), pd.DataFrame()

        imp_id = int(imp.iloc[0]['id'])

        credores = pd.read_sql(f"""
            SELECT credor,
                   ROUND(SUM(despesas_empenhadas)::numeric,2) AS empenhado,
                   ROUND(SUM(despesas_liquidadas)::numeric,2) AS liquidado,
                   ROUND(SUM(despesas_pagas)::numeric,2)      AS pago,
                   COUNT(DISTINCT nota_empenho)                AS qtd_ne
            FROM fato_execucao
            WHERE importacao_id = {imp_id}
            GROUP BY credor
            ORDER BY empenhado DESC
            LIMIT 20
        """, conn)

        contratos = pd.read_sql(f"""
            SELECT contrato_original, credor,
                   ROUND(SUM(despesas_empenhadas)::numeric,2) AS empenhado,
                   ROUND(SUM(despesas_pagas)::numeric,2)      AS pago
            FROM fato_execucao
            WHERE importacao_id = {imp_id}
              AND contrato_original IS NOT NULL
              AND contrato_original != ''
              AND contrato_original != '00000000'
            GROUP BY contrato_original, credor
            ORDER BY empenhado DESC
            LIMIT 20
        """, conn)

        conn.close()
        return imp, credores, contratos

    except Exception as e:
        st.error(f"Erro ao carregar dados: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# ── Interface ─────────────────────────────────────────
st.markdown("""
<div style='background:linear-gradient(135deg,#1F497D,#2E75B6);
            padding:20px 30px;border-radius:10px;margin-bottom:20px;color:white'>
    <h1 style='color:white;margin:0'>📊 AGERSA — Dashboard Orçamentário</h1>
    <p style='color:#BDD7EE;margin:5px 0 0'>UG: 043500 | Sistema de Execução Orçamentária</p>
</div>
""", unsafe_allow_html=True)

imp, credores, contratos = carregar()

if imp.empty:
    st.warning("Nenhum dado encontrado no banco.")
    st.info(f"DATABASE_URL: {'configurada' if DATABASE_URL else 'NAO encontrada'}")
    st.stop()

r    = imp.iloc[0]
emp  = float(r['empenhado']  or 0)
liq  = float(r['liquidado']  or 0)
pago = float(r['pago']       or 0)
apag = float(r['a_pagar']    or 0)
pct  = float(r['pct_execucao'] or 0)

def fm(v):
    if v >= 1e6: return f"R$ {v/1e6:.1f}M"
    if v >= 1e3: return f"R$ {v/1e3:.0f}K"
    return f"R$ {v:,.0f}"

st.subheader(f"Exercício {r['exercicio']} | Período {r['periodo']}")
c1,c2,c3,c4 = st.columns(4)
c1.metric("💰 Empenhado",  fm(emp))
c2.metric("📋 Liquidado",  fm(liq),  f"{liq/emp*100:.1f}%" if emp else "")
c3.metric("✅ Pago",       fm(pago), f"{pct:.1f}% execução")
c4.metric("⏳ A Pagar",    fm(apag))

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("🔽 Ciclo da Despesa")
    fig = go.Figure(go.Funnel(
        y=["Empenhado","Liquidado","Pago"],
        x=[emp, liq, pago],
        texttemplate="%{label}<br>R$ %{value:,.0f}",
        marker_color=["#1F497D","#2E75B6","#70AD47"]
    ))
    fig.update_layout(height=300, margin=dict(l=0,r=0,t=0,b=0))
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("👥 Top 10 Credores")
    if not credores.empty:
        top10 = credores.head(10).copy()
        top10['nome'] = top10['credor'].str.replace(
            r'^\d{8,14}\s*-\s*','',regex=True).str[:30]
        fig2 = px.bar(top10[::-1], x='empenhado', y='nome',
                      orientation='h',
                      labels={'empenhado':'Empenhado','nome':''})
        fig2.update_layout(height=300, margin=dict(l=0,r=0,t=0,b=0))
        st.plotly_chart(fig2, use_container_width=True)

st.divider()
tab1, tab2 = st.tabs(["👥 Credores","📋 Contratos"])

with tab1:
    busca = st.text_input("🔍 Filtrar:")
    df_c = credores.copy()
    if busca:
        df_c = df_c[df_c['credor'].str.contains(busca, case=False, na=False)]
    st.dataframe(df_c, use_container_width=True, height=350)

with tab2:
    st.dataframe(contratos, use_container_width=True, height=350)

st.caption(f"Atualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')} | "
           f"Registros: {int(r['total_linhas']):,} | Fonte: FlexVision SIAFE-RJ")
