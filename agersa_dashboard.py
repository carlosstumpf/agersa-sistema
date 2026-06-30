import os
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
from datetime import datetime

st.set_page_config(page_title="AGERSA Dashboard", page_icon="📊", layout="wide")

DATABASE_URL = os.environ.get("DATABASE_URL", "").strip()

def get_conn():
    import psycopg2
    if not DATABASE_URL:
        st.error("DATABASE_URL nao configurada no Render")
        st.stop()
    return psycopg2.connect(DATABASE_URL)

@st.cache_data(ttl=300)
def carregar():
    try:
        conn = get_conn()
        imp = pd.read_sql(
            "SELECT id, exercicio, mes_inicio || ' a ' || mes_fim AS periodo,"
            " empenhado, liquidado, pago, a_pagar, total_linhas,"
            " ROUND(pago::numeric / NULLIF(empenhado,0) * 100, 1) AS pct"
            " FROM importacoes ORDER BY data_importacao DESC LIMIT 1", conn)
        if imp.empty:
            conn.close()
            return imp, pd.DataFrame(), pd.DataFrame()
        imp_id = int(imp.iloc[0]['id'])
        cred = pd.read_sql(
            f"SELECT credor,"
            f" ROUND(SUM(despesas_empenhadas)::numeric,2) AS empenhado,"
            f" ROUND(SUM(despesas_pagas)::numeric,2) AS pago,"
            f" COUNT(DISTINCT nota_empenho) AS nes"
            f" FROM fato_execucao WHERE importacao_id={imp_id}"
            f" GROUP BY credor ORDER BY empenhado DESC LIMIT 15", conn)
        cont = pd.read_sql(
            f"SELECT contrato_original, credor,"
            f" ROUND(SUM(despesas_empenhadas)::numeric,2) AS empenhado,"
            f" ROUND(SUM(despesas_pagas)::numeric,2) AS pago"
            f" FROM fato_execucao WHERE importacao_id={imp_id}"
            f" AND contrato_original NOT IN ('','00000000')"
            f" AND contrato_original IS NOT NULL"
            f" GROUP BY contrato_original, credor ORDER BY empenhado DESC LIMIT 15", conn)
        conn.close()
        return imp, cred, cont
    except Exception as e:
        st.error(f"Erro: {e}")
        return pd.DataFrame(), pd.DataFrame(), pd.DataFrame()

# Header
st.markdown("<h1 style='color:#1F497D'>AGERSA - Dashboard Orcamentario</h1>", unsafe_allow_html=True)
st.markdown("<p style='color:#666'>UG: 043500 | FlexVision SIAFE-RJ</p>", unsafe_allow_html=True)

imp, cred, cont = carregar()

if imp.empty:
    st.warning("Nenhum dado no banco. Execute a extracao primeiro.")
    st.stop()

r    = imp.iloc[0]
emp  = float(r['empenhado'] or 0)
liq  = float(r['liquidado'] or 0)
pago = float(r['pago']      or 0)
apag = float(r['a_pagar']   or 0)
pct  = float(r['pct']       or 0)

def fm(v):
    if v >= 1e6: return f"R$ {v/1e6:.1f}M"
    if v >= 1e3: return f"R$ {v/1e3:.0f}K"
    return f"R$ {v:,.0f}"

st.subheader(f"Exercicio {r['exercicio']} | Periodo {r['periodo']}")

c1,c2,c3,c4 = st.columns(4)
c1.metric("Empenhado",  fm(emp))
c2.metric("Liquidado",  fm(liq))
c3.metric("Pago",       fm(pago), f"{pct:.1f}% exec")
c4.metric("A Pagar",    fm(apag))

st.divider()

col1, col2 = st.columns(2)

with col1:
    st.subheader("Ciclo da Despesa")
    fig = go.Figure(go.Funnel(
        y=["Empenhado","Liquidado","Pago"],
        x=[emp, liq, pago],
        marker_color=["#1F497D","#2E75B6","#70AD47"]
    ))
    fig.update_layout(height=280, margin=dict(l=0,r=0,t=0,b=0))
    st.plotly_chart(fig, use_container_width=True)

with col2:
    st.subheader("Top 10 Credores")
    if not cred.empty:
        top = cred.head(10).copy()
        top['nome'] = top['credor'].str.replace(r'^\d{8,14}\s*-\s*','',regex=True).str[:30]
        fig2 = px.bar(top[::-1], x='empenhado', y='nome', orientation='h',
                      labels={'empenhado':'Empenhado','nome':''})
        fig2.update_layout(height=280, margin=dict(l=0,r=0,t=0,b=0))
        st.plotly_chart(fig2, use_container_width=True)

st.divider()
t1, t2 = st.tabs(["Credores","Contratos"])
with t1:
    st.dataframe(cred, use_container_width=True)
with t2:
    st.dataframe(cont, use_container_width=True)

st.caption(f"Atualizado: {datetime.now().strftime('%d/%m/%Y %H:%M')} | {int(r['total_linhas']):,} registros")
