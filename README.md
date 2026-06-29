# AGERSA — Sistema de Extração Orçamentária

Sistema automatizado de extração, análise e visualização de dados orçamentários da AGERSA (UG 043500) via FlexVision SIAFE-RJ.

## Arquitetura

```
PC Casa (extração automática 07:00)
  └── agersa_sistema.py     → Playwright → FlexVision
  └── agersa_banco.py       → Envia dados → Supabase

Cloud (acesso de qualquer lugar)
  ├── Supabase              → Banco Postgres
  ├── Render                → Dashboard online
  └── Telegram              → Alertas no celular
```

## Módulos

| Arquivo | Função |
|---|---|
| `extrator/agersa_sistema.py` | Extração automática FlexVision |
| `banco/agersa_banco.py` | Banco histórico (SQLite local + Supabase cloud) |
| `dashboard/agersa_dashboard.py` | Dashboard Streamlit |
| `ia/agersa_ia.py` | Relatório executivo com IA |
| `extrator/agersa_agendador.py` | Agendador automático |
| `extrator/agersa_telegram.py` | Alertas Telegram |

## Instalação (PC de casa)

```bash
pip install -r requirements.txt
playwright install chromium
```

## Configuração

Copie `.env.exemplo` para `.env` e preencha:

```
SUPABASE_URL=sua_url
SUPABASE_KEY=sua_chave
TELEGRAM_TOKEN=seu_token
TELEGRAM_CHAT_ID=seu_chat_id
ANTHROPIC_API_KEY=sua_chave (opcional)
```

## Execução

```bash
# Extração manual
python extrator/agersa_sistema.py

# Agendador automático (roda todo dia às 07:00)
python extrator/agersa_agendador.py

# Dashboard local
streamlit run dashboard/agersa_dashboard.py
```

## Deploy (Render)

O dashboard é deployado automaticamente via GitHub Actions quando há push na branch `main`.

URL de acesso: `https://agersa-dashboard.onrender.com`
