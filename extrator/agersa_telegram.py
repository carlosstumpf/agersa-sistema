"""
AGERSA - Alertas Telegram
==========================
Envia resumo automático e alertas para o Telegram
após cada extração do FlexVision.

CONFIGURAÇÃO (fazer uma vez):
  1. No Telegram, procure @BotFather
  2. Digite /newbot e siga as instruções
  3. Copie o TOKEN gerado
  4. Procure @userinfobot e envie /start
  5. Copie seu CHAT_ID
  6. Preencha TELEGRAM_TOKEN e TELEGRAM_CHAT_ID abaixo

USO:
  python agersa_telegram.py              → envia resumo da extração mais recente
  python agersa_telegram.py --teste      → envia mensagem de teste
  python agersa_telegram.py --alerta     → envia só alertas críticos
"""

import sqlite3
import asyncio
import argparse
import sys
from pathlib import Path
from datetime import datetime

# ═══════════════════════════════════════════════════════
#  CONFIGURAÇÕES — PREENCHA AQUI
# ═══════════════════════════════════════════════════════
TELEGRAM_TOKEN   = ""   # Ex: "7123456789:AAFxxxxx..."
TELEGRAM_CHAT_ID = ""   # Ex: "123456789"

BANCO_PATH  = str(Path.home() / "Downloads" / "AGERSA" / "agersa_historico.db")

# ═══════════════════════════════════════════════════════
#  ENVIO DE MENSAGEM
# ═══════════════════════════════════════════════════════
async def enviar_mensagem(token, chat_id, texto):
    """Envia mensagem via Telegram Bot API."""
    try:
        from telegram import Bot
        bot = Bot(token=token)
        # Telegram tem limite de 4096 chars por mensagem
        if len(texto) > 4000:
            texto = texto[:3990] + "\n...(truncado)"
        await bot.send_message(
            chat_id=chat_id,
            text=texto,
            parse_mode='HTML'
        )
        print("✓ Mensagem Telegram enviada.")
        return True
    except Exception as e:
        print(f"[ERRO] Telegram: {e}")
        return False

def enviar(token, chat_id, texto):
    """Wrapper síncrono para envio."""
    return asyncio.run(enviar_mensagem(token, chat_id, texto))

# ═══════════════════════════════════════════════════════
#  DADOS DO BANCO
# ═══════════════════════════════════════════════════════
def carregar_resumo():
    """Carrega resumo da última extração do banco."""
    if not Path(BANCO_PATH).exists():
        return None

    conn = sqlite3.connect(BANCO_PATH)
    conn.row_factory = sqlite3.Row

    imp = conn.execute(
        "SELECT * FROM vw_resumo_importacoes LIMIT 1"
    ).fetchone()

    if not imp:
        conn.close()
        return None

    imp = dict(imp)
    imp_id = imp['id']

    # Totais
    totais = dict(conn.execute("""
        SELECT
            SUM(despesas_empenhadas)    AS empenhado,
            SUM(despesas_liquidadas)    AS liquidado,
            SUM(despesas_pagas)         AS pago,
            SUM(despesas_a_pagar)       AS a_pagar,
            SUM(rp_proc_inscritos)      AS rp_proc,
            SUM(rp_nao_proc_inscritos)  AS rp_nao_proc,
            COUNT(DISTINCT credor)      AS credores,
            COUNT(DISTINCT nota_empenho) AS empenhos,
            COUNT(DISTINCT contrato_siafe) AS contratos
        FROM fato_execucao WHERE importacao_id=?
    """, (imp_id,)).fetchone())

    # Top 5 credores
    credores = conn.execute("""
        SELECT credor, empenhado, pago
        FROM vw_top_credores
        WHERE importacao_id=?
        ORDER BY empenhado DESC LIMIT 5
    """, (imp_id,)).fetchall()

    conn.close()
    return imp, totais, list(credores)

# ═══════════════════════════════════════════════════════
#  FORMATAÇÃO DAS MENSAGENS
# ═══════════════════════════════════════════════════════
def formatar_moeda(valor):
    if not valor:
        return "R$ 0,00"
    v = float(valor)
    if v >= 1_000_000:
        return f"R$ {v/1_000_000:.1f}M"
    elif v >= 1_000:
        return f"R$ {v/1_000:.0f}K"
    return f"R$ {v:,.2f}"

def mensagem_resumo(imp, totais, credores):
    """Mensagem completa de resumo."""
    emp  = float(totais.get('empenhado') or 0)
    pago = float(totais.get('pago') or 0)
    liq  = float(totais.get('liquidado') or 0)
    a_pagar = float(totais.get('a_pagar') or 0)
    rp   = float(totais.get('rp_proc') or 0) + float(totais.get('rp_nao_proc') or 0)
    pct  = (pago / emp * 100) if emp > 0 else 0

    # Emoji de status
    if pct >= 90:
        status = "🟢"
    elif pct >= 75:
        status = "🟡"
    else:
        status = "🔴"

    agora = datetime.now().strftime('%d/%m/%Y %H:%M')

    msg = f"""📊 <b>AGERSA — Execução Orçamentária</b>
🏛 UG: 043500 | Período: {imp.get('periodo')} | {imp.get('exercicio')}
🕐 Atualizado: {agora}

{status} <b>Execução: {pct:.1f}%</b>

💰 <b>Ciclo da Despesa:</b>
  • Empenhado:  {formatar_moeda(emp)}
  • Liquidado:  {formatar_moeda(liq)}
  • Pago:       {formatar_moeda(pago)}
  • A Pagar:    {formatar_moeda(a_pagar)}

📌 <b>Restos a Pagar:</b> {formatar_moeda(rp)}

📋 <b>Contadores:</b>
  • Credores:  {int(totais.get('credores') or 0)}
  • Empenhos:  {int(totais.get('empenhos') or 0)}
  • Contratos: {int(totais.get('contratos') or 0)}

👥 <b>Top 5 Credores:</b>"""

    for i, c in enumerate(credores, 1):
        nome = str(c['credor'])
        # Remove CNPJ do início se houver
        import re
        nome = re.sub(r'^\d{8,14}\s*-\s*', '', nome)[:35]
        msg += f"\n  {i}. {nome}"
        msg += f"\n     {formatar_moeda(c['empenhado'])} emp | {formatar_moeda(c['pago'])} pago"

    return msg

def mensagem_alertas(totais):
    """Gera mensagem só com alertas críticos."""
    emp    = float(totais.get('empenhado') or 0)
    pago   = float(totais.get('pago') or 0)
    rp     = float(totais.get('rp_proc') or 0) + float(totais.get('rp_nao_proc') or 0)
    a_pagar = float(totais.get('a_pagar') or 0)
    pct    = (pago / emp * 100) if emp > 0 else 0

    alertas = []

    if pct < 70:
        alertas.append(f"🔴 Execução crítica: {pct:.1f}% pago")
    elif pct < 85:
        alertas.append(f"🟡 Execução baixa: {pct:.1f}% pago")

    if rp > emp * 0.15:
        alertas.append(f"🟡 RP elevado: {formatar_moeda(rp)} ({rp/emp*100:.1f}% do empenhado)")

    if a_pagar > emp * 0.05:
        alertas.append(f"🟠 A Pagar acumulado: {formatar_moeda(a_pagar)}")

    if not alertas:
        return None  # Sem alertas críticos

    msg = "⚠️ <b>AGERSA — Alertas Automáticos</b>\n"
    msg += f"📅 {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n"
    for a in alertas:
        msg += f"{a}\n"
    return msg

def mensagem_teste():
    """Mensagem simples de teste da conexão."""
    return (f"✅ <b>AGERSA Bot conectado!</b>\n"
            f"🕐 {datetime.now().strftime('%d/%m/%Y %H:%M')}\n\n"
            f"Sistema de alertas configurado com sucesso.\n"
            f"Você receberá notificações automáticas após cada extração.")

# ═══════════════════════════════════════════════════════
#  FUNÇÃO PRINCIPAL (usada pelo agendador)
# ═══════════════════════════════════════════════════════
def notificar(token=None, chat_id=None, apenas_alertas=False):
    """
    Função chamada pelo agendador após cada extração.
    Retorna True se enviou com sucesso.
    """
    token   = token   or TELEGRAM_TOKEN
    chat_id = chat_id or TELEGRAM_CHAT_ID

    if not token or not chat_id:
        print("[AVISO] Telegram não configurado — TOKEN e CHAT_ID necessários.")
        return False

    resultado = carregar_resumo()
    if not resultado:
        print("[AVISO] Nenhum dado no banco para notificar.")
        return False

    imp, totais, credores = resultado

    if apenas_alertas:
        msg = mensagem_alertas(totais)
        if not msg:
            print("  Sem alertas críticos — mensagem não enviada.")
            return True
    else:
        msg = mensagem_resumo(imp, totais, credores)

    return enviar(token, chat_id, msg)

# ═══════════════════════════════════════════════════════
#  MAIN
# ═══════════════════════════════════════════════════════
def main():
    parser = argparse.ArgumentParser(description='AGERSA — Alertas Telegram')
    parser.add_argument('--teste',  action='store_true', help='Envia mensagem de teste')
    parser.add_argument('--alerta', action='store_true', help='Envia só alertas críticos')
    parser.add_argument('--token',   help='Token do bot (opcional)')
    parser.add_argument('--chat-id', help='Chat ID (opcional)')
    args = parser.parse_args()

    token   = args.token   or TELEGRAM_TOKEN
    chat_id = args.chat_id or TELEGRAM_CHAT_ID

    if not token or not chat_id:
        print("""
[CONFIGURAÇÃO NECESSÁRIA]

1. Abra o Telegram e procure: @BotFather
2. Digite: /newbot
3. Escolha um nome e username para o bot
4. Copie o TOKEN que aparecer

5. Procure: @userinfobot
6. Digite: /start
7. Copie o número do seu ID (Chat ID)

8. Preencha no arquivo agersa_telegram.py:
   TELEGRAM_TOKEN   = "seu_token_aqui"
   TELEGRAM_CHAT_ID = "seu_chat_id_aqui"

Ou passe como argumento:
   python agersa_telegram.py --token SEU_TOKEN --chat-id SEU_CHAT_ID
        """)
        sys.exit(1)

    if args.teste:
        print("Enviando mensagem de teste...")
        enviar(token, chat_id, mensagem_teste())
        return

    if args.alerta:
        print("Verificando alertas críticos...")
        notificar(token, chat_id, apenas_alertas=True)
        return

    # Padrão: envia resumo completo
    print("Enviando resumo completo...")
    notificar(token, chat_id, apenas_alertas=False)

if __name__ == "__main__":
    main()
