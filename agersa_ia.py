"""
AGERSA - Analista IA
=====================
Gera resumos executivos, alertas e insights automáticos
usando a API do Claude sobre os dados do banco histórico.

USO:
  python agersa_ia.py                    → analisa extração mais recente
  python agersa_ia.py --enviar-email     → envia resumo por e-mail
  python agersa_ia.py --salvar           → salva resumo em arquivo

DEPENDENCIAS:
  pip install anthropic
"""

import sqlite3
import json
import argparse
import smtplib
import sys
from pathlib import Path
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ═══════════════════════════════════════════════════════
#  CONFIGURAÇÕES
# ═══════════════════════════════════════════════════════
BANCO_PATH   = str(Path.home() / "Downloads" / "AGERSA" / "agersa_historico.db")
PASTA_SAIDA  = Path.home() / "Downloads" / "AGERSA"

# Claude API — chave via variável de ambiente ou digitada
# Para configurar: set ANTHROPIC_API_KEY=sua_chave
# Obtenha em: https://console.anthropic.com
ANTHROPIC_API_KEY = ""  # deixe vazio para usar variável de ambiente

# E-mail (opcional)
EMAIL_ATIVO    = False
EMAIL_REMETENTE = "seu_email@gmail.com"
EMAIL_SENHA     = "sua_senha_app"
EMAIL_DESTINO   = "destino@email.com"
EMAIL_SMTP      = "smtp.gmail.com"
EMAIL_PORTA     = 587

# ═══════════════════════════════════════════════════════
#  EXTRAÇÃO DE DADOS DO BANCO
# ═══════════════════════════════════════════════════════
def carregar_dados_banco(importacao_id=None):
    """Carrega dados estruturados do banco para análise."""
    conn = sqlite3.connect(BANCO_PATH)
    conn.row_factory = sqlite3.Row

    # Última importação se não especificado
    if importacao_id is None:
        row = conn.execute(
            "SELECT id FROM importacoes ORDER BY data_importacao DESC LIMIT 1"
        ).fetchone()
        if not row:
            conn.close()
            return None
        importacao_id = row['id']

    # Metadados da importação
    imp = dict(conn.execute(
        "SELECT * FROM vw_resumo_importacoes WHERE id=?", (importacao_id,)
    ).fetchone())

    # Totais gerais
    totais = dict(conn.execute("""
        SELECT
            SUM(despesas_empenhadas)  AS empenhado,
            SUM(despesas_liquidadas)  AS liquidado,
            SUM(despesas_pagas)       AS pago,
            SUM(despesas_a_pagar)     AS a_pagar,
            SUM(rp_proc_inscritos)    AS rp_proc_inscrito,
            SUM(rp_proc_a_pagar)      AS rp_proc_a_pagar,
            SUM(rp_proc_pagos)        AS rp_proc_pago,
            SUM(rp_nao_proc_inscritos) AS rp_nao_proc,
            SUM(rp_nao_proc_pagos)    AS rp_nao_proc_pago,
            SUM(retencoes_pagas)      AS retencoes,
            COUNT(DISTINCT credor)    AS qtd_credores,
            COUNT(DISTINCT nota_empenho) AS qtd_ne,
            COUNT(DISTINCT contrato_siafe) AS qtd_contratos
        FROM fato_execucao
        WHERE importacao_id=?
    """, (importacao_id,)).fetchone())

    # Top 10 credores
    credores = [dict(r) for r in conn.execute("""
        SELECT credor, empenhado, liquidado, pago, qtd_empenhos
        FROM vw_top_credores
        WHERE importacao_id=?
        ORDER BY empenhado DESC LIMIT 10
    """, (importacao_id,)).fetchall()]

    # Top 10 contratos
    contratos = [dict(r) for r in conn.execute("""
        SELECT contrato_original, credor, empenhado, liquidado, pago, a_pagar
        FROM vw_contratos
        WHERE importacao_id=? AND contrato_original != '00000000'
        ORDER BY empenhado DESC LIMIT 10
    """, (importacao_id,)).fetchall()]

    # Histórico de importações anteriores para comparação
    historico = [dict(r) for r in conn.execute("""
        SELECT periodo, empenhado, liquidado, pago, pct_execucao
        FROM vw_resumo_importacoes
        ORDER BY data_importacao DESC LIMIT 5
    """).fetchall()]

    # Alertas automáticos baseados em regras
    alertas = gerar_alertas(totais, credores)

    conn.close()

    return {
        'importacao':  imp,
        'totais':      totais,
        'credores':    credores,
        'contratos':   contratos,
        'historico':   historico,
        'alertas':     alertas,
    }

def gerar_alertas(totais, credores):
    """Gera alertas automáticos baseados em regras."""
    alertas = []
    emp = totais.get('empenhado') or 0
    pago = totais.get('pago') or 0
    a_pagar = totais.get('a_pagar') or 0
    rp_proc = totais.get('rp_proc_inscrito') or 0
    rp_nao  = totais.get('rp_nao_proc') or 0

    pct_exec = (pago / emp * 100) if emp > 0 else 0

    if pct_exec < 70:
        alertas.append({
            'nivel': 'CRITICO',
            'mensagem': f'Execução muito baixa: {pct_exec:.1f}% do empenhado pago'
        })
    elif pct_exec < 85:
        alertas.append({
            'nivel': 'ATENCAO',
            'mensagem': f'Execução abaixo do esperado: {pct_exec:.1f}%'
        })

    rp_total = rp_proc + rp_nao
    if rp_total > emp * 0.15:
        alertas.append({
            'nivel': 'ATENCAO',
            'mensagem': f'Restos a pagar elevados: R$ {rp_total:,.2f} ({rp_total/emp*100:.1f}% do empenhado)'
        })

    if a_pagar > emp * 0.05:
        alertas.append({
            'nivel': 'ATENCAO',
            'mensagem': f'Despesas a pagar acumuladas: R$ {a_pagar:,.2f}'
        })

    if credores:
        maior = credores[0]
        pct_credor = (maior['empenhado'] / emp * 100) if emp > 0 else 0
        if pct_credor > 25:
            alertas.append({
                'nivel': 'INFO',
                'mensagem': f'Alta concentração: {maior["credor"][:40]} representa {pct_credor:.1f}% do total'
            })

    if not alertas:
        alertas.append({
            'nivel': 'OK',
            'mensagem': 'Execução dentro dos parâmetros esperados'
        })

    return alertas

# ═══════════════════════════════════════════════════════
#  ANÁLISE COM CLAUDE
# ═══════════════════════════════════════════════════════
def analisar_com_claude(dados, api_key):
    """Envia dados para Claude e recebe análise em linguagem natural."""
    try:
        import anthropic
    except ImportError:
        print("[ERRO] anthropic não instalado. Execute: pip install anthropic")
        return None

    cliente = anthropic.Anthropic(api_key=api_key)

    imp  = dados['importacao']
    tot  = dados['totais']
    emp  = tot.get('empenhado') or 0
    pago = tot.get('pago') or 0
    pct  = (pago / emp * 100) if emp > 0 else 0

    # Prepara contexto compacto para o Claude
    contexto = f"""
Você é um analista financeiro especializado em finanças públicas brasileiras.
Analise os dados de execução orçamentária da AGERSA (Agência Reguladora de
Saneamento e Energia do Estado do Rio de Janeiro) e gere um relatório executivo.

DADOS DO PERÍODO:
- Exercício: {imp.get('exercicio')}
- Período: {imp.get('periodo')}
- Total de registros: {imp.get('total_linhas'):,}

CICLO DA DESPESA (R$):
- Empenhado:  R$ {emp:,.2f}
- Liquidado:  R$ {(tot.get('liquidado') or 0):,.2f}
- Pago:       R$ {pago:,.2f}
- A Pagar:    R$ {(tot.get('a_pagar') or 0):,.2f}
- % Execução: {pct:.1f}%

RESTOS A PAGAR:
- RP Processados Inscritos: R$ {(tot.get('rp_proc_inscrito') or 0):,.2f}
- RP Processados a Pagar:   R$ {(tot.get('rp_proc_a_pagar') or 0):,.2f}
- RP Processados Pagos:     R$ {(tot.get('rp_proc_pago') or 0):,.2f}
- RP Não Processados:       R$ {(tot.get('rp_nao_proc') or 0):,.2f}

CONTADORES:
- Credores distintos: {int(tot.get('qtd_credores') or 0)}
- Notas de Empenho:   {int(tot.get('qtd_ne') or 0)}
- Contratos:          {int(tot.get('qtd_contratos') or 0)}

TOP 5 CREDORES:
{chr(10).join([
    f"  {i+1}. {c['credor'][:50]} — Empenhado: R$ {c['empenhado']:,.2f} | Pago: R$ {c['pago']:,.2f}"
    for i, c in enumerate(dados['credores'][:5])
])}

TOP 5 CONTRATOS:
{chr(10).join([
    f"  {i+1}. {c['contrato_original']} ({c['credor'][:30]}) — R$ {c['empenhado']:,.2f}"
    for i, c in enumerate(dados['contratos'][:5])
])}

ALERTAS IDENTIFICADOS:
{chr(10).join([f"  [{a['nivel']}] {a['mensagem']}" for a in dados['alertas']])}

INSTRUÇÃO:
Gere um relatório executivo em português com:

1. **RESUMO EXECUTIVO** (3-4 linhas): situação geral da execução orçamentária
2. **ANÁLISE DO CICLO DA DESPESA**: interprete os números de empenho, liquidação e pagamento
3. **RESTOS A PAGAR**: avalie o volume e implicações
4. **DESTAQUES**: principais credores e contratos, observe concentrações
5. **PONTOS DE ATENÇÃO**: riscos ou irregularidades identificadas
6. **RECOMENDAÇÕES**: 2-3 ações sugeridas para a gestão

Use linguagem técnica mas acessível para gestores públicos.
Seja objetivo e direto. Máximo 600 palavras.
"""

    print("Consultando Claude API...")
    resposta = cliente.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1000,
        messages=[{"role": "user", "content": contexto}]
    )

    return resposta.content[0].text

# ═══════════════════════════════════════════════════════
#  FORMATAÇÃO DO RELATÓRIO
# ═══════════════════════════════════════════════════════
def formatar_relatorio(dados, analise_ia=None):
    """Monta o relatório final completo."""
    imp  = dados['importacao']
    tot  = dados['totais']
    emp  = tot.get('empenhado') or 0
    pago = tot.get('pago') or 0
    pct  = (pago / emp * 100) if emp > 0 else 0
    agora = datetime.now().strftime('%d/%m/%Y %H:%M')

    separador = "=" * 65

    relatorio = f"""
{separador}
  AGERSA — RELATÓRIO EXECUTIVO DE EXECUÇÃO ORÇAMENTÁRIA
  UG: 043500 | Período: {imp.get('periodo')} | Gerado: {agora}
{separador}

── INDICADORES PRINCIPAIS ──────────────────────────────────

  Empenhado:      R$ {emp:>18,.2f}
  Liquidado:      R$ {(tot.get('liquidado') or 0):>18,.2f}
  Pago:           R$ {pago:>18,.2f}  ({pct:.1f}% de execução)
  A Pagar:        R$ {(tot.get('a_pagar') or 0):>18,.2f}

  RP Proc. Inscrito: R$ {(tot.get('rp_proc_inscrito') or 0):>15,.2f}
  RP Não Proc.:      R$ {(tot.get('rp_nao_proc') or 0):>15,.2f}

  Credores:    {int(tot.get('qtd_credores') or 0):>4}
  Empenhos:    {int(tot.get('qtd_ne') or 0):>4}
  Contratos:   {int(tot.get('qtd_contratos') or 0):>4}

── ALERTAS ─────────────────────────────────────────────────
"""
    for a in dados['alertas']:
        emoji = {'CRITICO':'🔴','ATENCAO':'🟡','INFO':'🔵','OK':'✅'}.get(a['nivel'],'⚪')
        relatorio += f"\n  {emoji} [{a['nivel']}] {a['mensagem']}"

    relatorio += f"""

── TOP 5 CREDORES ──────────────────────────────────────────
"""
    for i, c in enumerate(dados['credores'][:5], 1):
        nome = c['credor'][:50]
        pct_c = (c['empenhado'] / emp * 100) if emp > 0 else 0
        relatorio += f"\n  {i}. {nome}"
        relatorio += f"\n     Empenhado: R$ {c['empenhado']:,.2f} ({pct_c:.1f}%) | Pago: R$ {c['pago']:,.2f}"

    relatorio += f"""

── TOP 5 CONTRATOS ─────────────────────────────────────────
"""
    for i, c in enumerate(dados['contratos'][:5], 1):
        credor_nome = c['credor'][:35]
        relatorio += f"\n  {i}. Contrato {c['contrato_original']} — {credor_nome}"
        relatorio += f"\n     R$ {c['empenhado']:,.2f} empenhado | R$ {c['pago']:,.2f} pago"

    if analise_ia:
        relatorio += f"""

{separador}
  ANÁLISE INTELIGENTE (Claude AI)
{separador}

{analise_ia}
"""

    relatorio += f"""

{separador}
  Fonte: FlexVision SIAFE-RJ | AGERSA UG 043500
  Banco: {BANCO_PATH}
{separador}
"""
    return relatorio

# ═══════════════════════════════════════════════════════
#  ENVIO POR E-MAIL
# ═══════════════════════════════════════════════════════
def enviar_email(assunto, corpo):
    if not EMAIL_ATIVO:
        return
    try:
        msg = MIMEMultipart()
        msg['From']    = EMAIL_REMETENTE
        msg['To']      = EMAIL_DESTINO
        msg['Subject'] = assunto
        msg.attach(MIMEText(corpo, 'plain', 'utf-8'))
        with smtplib.SMTP(EMAIL_SMTP, EMAIL_PORTA) as s:
            s.starttls()
            s.login(EMAIL_REMETENTE, EMAIL_SENHA)
            s.send_message(msg)
        print(f"✓ E-mail enviado para {EMAIL_DESTINO}")
    except Exception as e:
        print(f"[AVISO] E-mail: {e}")

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser(description='AGERSA — Analista IA')
    parser.add_argument('--importacao', type=int, help='ID da importação (padrão: mais recente)')
    parser.add_argument('--enviar-email', action='store_true', help='Envia relatório por e-mail')
    parser.add_argument('--salvar',       action='store_true', help='Salva relatório em arquivo')
    parser.add_argument('--sem-ia',       action='store_true', help='Gera só relatório numérico, sem Claude')
    args = parser.parse_args()

    # Verifica banco
    if not Path(BANCO_PATH).exists():
        print(f"[ERRO] Banco não encontrado: {BANCO_PATH}")
        print("Execute primeiro: python agersa_banco.py")
        sys.exit(1)

    print("Carregando dados do banco...")
    dados = carregar_dados_banco(args.importacao)
    if not dados:
        print("[ERRO] Nenhuma importação encontrada no banco.")
        sys.exit(1)

    imp = dados['importacao']
    print(f"  → Período: {imp.get('periodo')} | Exercício: {imp.get('exercicio')}")

    # Análise com IA
    analise_ia = None
    if not args.sem_ia:
        api_key = (ANTHROPIC_API_KEY or
                   __import__('os').environ.get('ANTHROPIC_API_KEY', ''))
        if not api_key:
            print("\n[INFO] Chave da API Claude não configurada.")
            print("  Para usar IA: set ANTHROPIC_API_KEY=sua_chave")
            print("  Gerando relatório sem IA...\n")
        else:
            analise_ia = analisar_com_claude(dados, api_key)

    # Monta relatório
    relatorio = formatar_relatorio(dados, analise_ia)
    print(relatorio)

    # Salva em arquivo
    if args.salvar or True:  # sempre salva
        nome = f"relatorio_{datetime.now().strftime('%Y%m%d_%H%M')}.txt"
        caminho = PASTA_SAIDA / nome
        caminho.parent.mkdir(parents=True, exist_ok=True)
        caminho.write_text(relatorio, encoding='utf-8')
        print(f"✓ Relatório salvo: {caminho}")

    # Envia por e-mail
    if args.enviar_email:
        assunto = (f"[AGERSA] Relatório Orçamentário — "
                   f"{imp.get('exercicio')} Período {imp.get('periodo')} — "
                   f"{datetime.now().strftime('%d/%m/%Y')}")
        enviar_email(assunto, relatorio)

    return relatorio

if __name__ == "__main__":
    main()
