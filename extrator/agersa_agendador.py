"""
AGERSA - Agendador Automático
==============================
Roda a extração automaticamente no horário configurado.
Salva log em arquivo e envia e-mail em caso de falha.

COMO USAR:
  1. Configure as variáveis abaixo (HORARIO, EMAIL, etc.)
  2. Execute: python agersa_agendador.py
  3. Deixe rodando em segundo plano

PARA RODAR AO INICIAR O WINDOWS AUTOMATICAMENTE:
  - Win+R → shell:startup
  - Crie um atalho para: python agersa_agendador.py
"""

import schedule
import time
import subprocess
import sys
import logging
import smtplib
import os
from pathlib import Path
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# ═══════════════════════════════════════════════════════
#  CONFIGURAÇÕES — EDITE AQUI
# ═══════════════════════════════════════════════════════

# Pasta onde está o script principal
PASTA_SCRIPTS = Path(__file__).parent

# Script principal de extração
SCRIPT_EXTRACAO = str(PASTA_SCRIPTS / "agersa_sistema.py")

# Python a usar
PYTHON = sys.executable

# Pasta de logs
PASTA_LOGS = Path.home() / "Downloads" / "AGERSA" / "logs"
PASTA_LOGS.mkdir(parents=True, exist_ok=True)

# ── Horários de execução ──────────────────────────────
# Formato: "HH:MM"
HORARIOS = [
    "07:00",   # roda de manhã cedo
    # "12:00", # descomente para rodar ao meio-dia também
    # "18:00", # descomente para rodar à tarde
]

# Dias da semana para rodar (0=segunda, 6=domingo)
# None = todos os dias
# [0,1,2,3,4] = segunda a sexta
DIAS_SEMANA = [0, 1, 2, 3, 4]  # segunda a sexta

# ── Notificação por e-mail (opcional) ────────────────
NOTIFICAR_EMAIL = False  # True para ativar
EMAIL_REMETENTE = "seu_email@gmail.com"
EMAIL_SENHA     = "sua_senha_app"      # senha de app Gmail
EMAIL_DESTINO   = "destino@email.com"
SMTP_HOST       = "smtp.gmail.com"
SMTP_PORT       = 587

# ── Credenciais FlexVision ───────────────────────────
# Deixe vazio para digitar na interface
# Ou preencha para rodar sem interação (rotina automática)
USUARIO_FLEXVISION = ""
SENHA_FLEXVISION   = ""

# ═══════════════════════════════════════════════════════
#  LOGGING
# ═══════════════════════════════════════════════════════

def configurar_log():
    log_file = PASTA_LOGS / f"agendador_{datetime.now().strftime('%Y%m')}.log"
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s [%(levelname)s] %(message)s',
        handlers=[
            logging.FileHandler(str(log_file), encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    return logging.getLogger("AGERSA")

logger = configurar_log()

# ═══════════════════════════════════════════════════════
#  NOTIFICAÇÃO E-MAIL
# ═══════════════════════════════════════════════════════

def enviar_email(assunto, corpo):
    if not NOTIFICAR_EMAIL:
        return
    try:
        msg = MIMEMultipart()
        msg['From']    = EMAIL_REMETENTE
        msg['To']      = EMAIL_DESTINO
        msg['Subject'] = assunto
        msg.attach(MIMEText(corpo, 'plain', 'utf-8'))

        with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
            s.starttls()
            s.login(EMAIL_REMETENTE, EMAIL_SENHA)
            s.send_message(msg)
        logger.info(f"E-mail enviado: {assunto}")
    except Exception as e:
        logger.error(f"Erro ao enviar e-mail: {e}")

# ═══════════════════════════════════════════════════════
#  EXECUÇÃO DA EXTRAÇÃO
# ═══════════════════════════════════════════════════════

def executar_extracao():
    """Roda o script de extração e registra o resultado."""

    agora = datetime.now()

    # Verifica dia da semana
    if DIAS_SEMANA and agora.weekday() not in DIAS_SEMANA:
        logger.info(f"Dia {agora.strftime('%A')} — extração não agendada.")
        return

    logger.info("=" * 55)
    logger.info(f"INICIANDO EXTRAÇÃO AUTOMÁTICA — {agora.strftime('%d/%m/%Y %H:%M')}")
    logger.info("=" * 55)

    log_extracao = PASTA_LOGS / f"extracao_{agora.strftime('%Y%m%d_%H%M')}.log"

    try:
        # Executa o script principal
        resultado = subprocess.run(
            [PYTHON, SCRIPT_EXTRACAO],
            capture_output=True,
            text=True,
            encoding='utf-8',
            timeout=1800  # 30 minutos máximo
        )

        # Salva log da extração
        with open(str(log_extracao), 'w', encoding='utf-8') as f:
            f.write(f"=== EXTRAÇÃO {agora.strftime('%d/%m/%Y %H:%M')} ===\n\n")
            f.write("=== STDOUT ===\n")
            f.write(resultado.stdout or "")
            f.write("\n=== STDERR ===\n")
            f.write(resultado.stderr or "")

        if resultado.returncode == 0:
            logger.info("✓ Extração concluída com sucesso.")

            # 1. Banco histórico
            try:
                banco_script = str(Path(__file__).parent / "agersa_banco.py")
                subprocess.run([PYTHON, banco_script], timeout=120)
                logger.info("✓ Banco histórico atualizado.")
            except Exception as e:
                logger.warning(f"  Banco: {e}")

            # 2. Relatório IA
            try:
                ia_script = str(Path(__file__).parent / "agersa_ia.py")
                subprocess.run([PYTHON, ia_script, "--sem-ia"], timeout=120)
                logger.info("✓ Relatório IA gerado.")
            except Exception as e:
                logger.warning(f"  IA: {e}")

            # 3. Telegram
            try:
                tg_script = str(Path(__file__).parent / "agersa_telegram.py")
                subprocess.run([PYTHON, tg_script], timeout=30)
                logger.info("✓ Telegram notificado.")
            except Exception as e:
                logger.warning(f"  Telegram: {e}")
            logger.info(f"  Log: {log_extracao}")

            # E-mail de sucesso (opcional)
            enviar_email(
                f"[AGERSA] Extração concluída — {agora.strftime('%d/%m/%Y')}",
                f"Extração automática concluída com sucesso.\n"
                f"Horário: {agora.strftime('%d/%m/%Y %H:%M')}\n"
                f"Log: {log_extracao}\n\n"
                f"Saída:\n{resultado.stdout[-2000:]}"
            )
        else:
            logger.error(f"✗ Extração falhou (código {resultado.returncode}).")
            logger.error(f"  Log: {log_extracao}")

            # E-mail de falha
            enviar_email(
                f"[AGERSA] ⚠ FALHA na extração — {agora.strftime('%d/%m/%Y')}",
                f"A extração automática falhou.\n"
                f"Horário: {agora.strftime('%d/%m/%Y %H:%M')}\n"
                f"Código de saída: {resultado.returncode}\n\n"
                f"Erro:\n{resultado.stderr[-2000:]}\n\n"
                f"Log completo: {log_extracao}"
            )

    except subprocess.TimeoutExpired:
        logger.error("✗ Extração excedeu 30 minutos — cancelada.")
        enviar_email(
            f"[AGERSA] ⚠ TIMEOUT na extração — {agora.strftime('%d/%m/%Y')}",
            "A extração demorou mais de 30 minutos e foi cancelada automaticamente."
        )

    except Exception as e:
        logger.error(f"✗ Erro inesperado: {e}")
        enviar_email(
            f"[AGERSA] ⚠ ERRO na extração — {agora.strftime('%d/%m/%Y')}",
            f"Erro inesperado: {e}"
        )

    logger.info("=" * 55)

# ═══════════════════════════════════════════════════════
#  AGENDAMENTO
# ═══════════════════════════════════════════════════════

def configurar_agenda():
    """Configura os horários de execução."""
    for horario in HORARIOS:
        schedule.every().day.at(horario).do(executar_extracao)
        logger.info(f"  Agendado: todo dia às {horario}")

    logger.info(f"  Dias: {'seg-sex' if DIAS_SEMANA == [0,1,2,3,4] else str(DIAS_SEMANA)}")
    logger.info(f"  Script: {SCRIPT_EXTRACAO}")
    logger.info(f"  Logs: {PASTA_LOGS}")

def main():
    logger.info("=" * 55)
    logger.info("  AGERSA — Agendador Automático")
    logger.info("=" * 55)

    if not Path(SCRIPT_EXTRACAO).exists():
        logger.error(f"Script não encontrado: {SCRIPT_EXTRACAO}")
        logger.error("Coloque o agendador na mesma pasta do agersa_sistema.py")
        sys.exit(1)

    configurar_agenda()

    proxima = schedule.next_run()
    logger.info(f"  Próxima execução: {proxima.strftime('%d/%m/%Y %H:%M') if proxima else 'N/A'}")
    logger.info("  Aguardando... (Ctrl+C para parar)")
    logger.info("=" * 55)

    # Loop principal
    while True:
        schedule.run_pending()
        time.sleep(30)  # verifica a cada 30 segundos

if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        logger.info("Agendador encerrado pelo usuário.")
