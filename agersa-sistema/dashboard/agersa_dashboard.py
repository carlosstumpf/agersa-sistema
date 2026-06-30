import os
import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from datetime import datetime
st.set_page_config(page_title='AGERSA Dashboard', layout='wide')
DATABASE_URL = os.environ.get('DATABASE_URL', '').strip()
if not DATABASE_URL:
    st.error('DATABASE_URL nao encontrada no Render - configure em Environment')
    st.stop()
import psycopg2
def carregar():
    conn = psycopg2.connect(DATABASE_URL)
    imp = pd.read_sql('SELECT id, exercicio, mes_inicio || chr(32) || chr(97) || chr(32) || mes_fim AS periodo, empenhado, liquidado, pago, a_pagar, total_linhas, ROUND(pago::numeric / NULLIF(empenhado,0) * 100, 1) AS pct FROM importacoes ORDER BY data_importacao DESC LIMIT 1', conn)
    if imp.empty: conn.close(); return imp, pd.DataFrame()
    cred = pd.read_sql(f'SELECT credor, ROUND(SUM(despesas_empenhadas)::numeric,2) AS empenhado, ROUND(SUM(despesas_pagas)::numeric,2) AS pago FROM fato_execucao WHERE importacao_id={int(imp.iloc[0][chr(105)+chr(100)])} GROUP BY credor ORDER BY empenhado DESC LIMIT 15', conn)
    conn.close(); return imp, cred
st.title('AGERSA - Execucao Orcamentaria')
imp, cred = carregar()
if imp.empty: st.warning('Sem dados'); st.stop()
r=imp.iloc[0]
c1,c2,c3,c4=st.columns(4)
c1.metric('Empenhado',f'R$ {float(r.empenhado)/1e6:.1f}M')
c2.metric('Liquidado',f'R$ {float(r.liquidado)/1e6:.1f}M')
c3.metric('Pago',f'R$ {float(r.pago)/1e6:.1f}M',f'{float(r.pct):.1f}%')
c4.metric('A Pagar',f'R$ {float(r.a_pagar)/1e6:.1f}M')
st.divider()
if not cred.empty:
    top=cred.head(10).copy()
    top['nome']=top['credor'].str.replace(r'^\d+-\s*','',regex=True).str[:35]
    fig=go.Figure(go.Bar(x=top['empenhado'],y=top['nome'],orientation='h',marker_color='#2E75B6'))
    fig.update_layout(height=350,margin=dict(l=0,r=0,t=20,b=0))
    st.plotly_chart(fig,use_container_width=True)
st.dataframe(cred,use_container_width=True)
