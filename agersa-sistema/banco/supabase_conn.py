"""
AGERSA - Conexão Supabase
==========================
Gerencia a conexão com o banco Postgres na cloud.
"""

import os
from pathlib import Path

# ── Carrega .env se existir ──────────────────────────
env_path = Path(__file__).parent.parent / ".env"
if env_path.exists():
    for line in env_path.read_text(encoding='utf-8').splitlines():
        line = line.strip()
        if line and not line.startswith('#') and '=' in line:
            k, v = line.split('=', 1)
            os.environ.setdefault(k.strip(), v.strip())

# ── Configurações ────────────────────────────────────
SUPABASE_URL = os.environ.get('SUPABASE_URL', '')
SUPABASE_KEY = os.environ.get('SUPABASE_KEY', '')
DATABASE_URL = os.environ.get('DATABASE_URL', '')

def get_conn():
    """Retorna conexão psycopg2 com o Supabase."""
    import psycopg2
    import urllib.parse

    if not DATABASE_URL:
        raise ValueError("DATABASE_URL não configurada no .env")

    return psycopg2.connect(DATABASE_URL)

def get_supabase():
    """Retorna cliente Supabase."""
    from supabase import create_client
    if not SUPABASE_URL or not SUPABASE_KEY:
        raise ValueError("SUPABASE_URL e SUPABASE_KEY não configuradas no .env")
    return create_client(SUPABASE_URL, SUPABASE_KEY)

def testar_conexao():
    """Testa conexão com o Supabase."""
    try:
        conn = get_conn()
        cur = conn.cursor()
        cur.execute("SELECT COUNT(*) FROM importacoes")
        total = cur.fetchone()[0]
        conn.close()
        print(f"✓ Supabase conectado — {total} importações no banco")
        return True
    except Exception as e:
        print(f"✗ Erro Supabase: {e}")
        return False

if __name__ == "__main__":
    testar_conexao()
