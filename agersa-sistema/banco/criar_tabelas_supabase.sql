-- ================================================
-- AGERSA - Criação das Tabelas no Supabase
-- Execute este SQL no Supabase SQL Editor
-- ================================================

-- Tabela de importações
CREATE TABLE IF NOT EXISTS importacoes (
    id              SERIAL PRIMARY KEY,
    arquivo         TEXT NOT NULL,
    data_importacao TIMESTAMPTZ DEFAULT NOW(),
    exercicio       INTEGER,
    mes_inicio      INTEGER,
    mes_fim         INTEGER,
    total_linhas    INTEGER,
    empenhado       NUMERIC(18,2),
    liquidado       NUMERIC(18,2),
    pago            NUMERIC(18,2),
    a_pagar         NUMERIC(18,2),
    rp_processado   NUMERIC(18,2),
    rp_nao_proc     NUMERIC(18,2),
    contratos       INTEGER,
    credores        INTEGER,
    notas_empenho   INTEGER
);

-- Tabela de execução orçamentária
CREATE TABLE IF NOT EXISTS fato_execucao (
    id                      SERIAL PRIMARY KEY,
    importacao_id           INTEGER REFERENCES importacoes(id),
    exercicio               INTEGER,
    mes_referencia          INTEGER,

    -- Dimensões
    unidade_orcamentaria    TEXT,
    unidade_gestora         TEXT,
    credor                  TEXT,
    nota_empenho            TEXT,
    historico               TEXT,
    processo_sei            TEXT,
    nota_liquidacao         TEXT,
    nf                      TEXT,
    programacao_desembolso  TEXT,
    ordem_bancaria          TEXT,
    contrato_siafe          TEXT,
    contrato_original       TEXT,

    -- Valores
    despesas_empenhadas     NUMERIC(18,2) DEFAULT 0,
    desp_emp_a_liquidar     NUMERIC(18,2) DEFAULT 0,
    despesas_liquidadas     NUMERIC(18,2) DEFAULT 0,
    desp_liq_a_pagar        NUMERIC(18,2) DEFAULT 0,
    despesas_a_pagar        NUMERIC(18,2) DEFAULT 0,
    despesas_pagas          NUMERIC(18,2) DEFAULT 0,
    rp_nao_proc_inscritos   NUMERIC(18,2) DEFAULT 0,
    rp_nao_proc_pagos       NUMERIC(18,2) DEFAULT 0,
    rp_proc_pagos           NUMERIC(18,2) DEFAULT 0,
    rp_proc_a_pagar         NUMERIC(18,2) DEFAULT 0,
    rp_proc_inscritos       NUMERIC(18,2) DEFAULT 0,
    rp_pagos                NUMERIC(18,2) DEFAULT 0,
    retencoes_pagas         NUMERIC(18,2) DEFAULT 0
);

-- Índices para performance
CREATE INDEX IF NOT EXISTS idx_fato_importacao  ON fato_execucao(importacao_id);
CREATE INDEX IF NOT EXISTS idx_fato_exercicio   ON fato_execucao(exercicio);
CREATE INDEX IF NOT EXISTS idx_fato_credor      ON fato_execucao(credor);
CREATE INDEX IF NOT EXISTS idx_fato_contrato    ON fato_execucao(contrato_siafe);

-- View: resumo por importação
CREATE OR REPLACE VIEW vw_resumo_importacoes AS
SELECT
    i.id,
    i.arquivo,
    i.data_importacao,
    i.exercicio,
    i.mes_inicio || ' a ' || i.mes_fim AS periodo,
    i.total_linhas,
    i.empenhado,
    i.liquidado,
    i.pago,
    i.a_pagar,
    ROUND(i.pago / NULLIF(i.empenhado, 0) * 100, 1) AS pct_execucao
FROM importacoes i
ORDER BY i.data_importacao DESC;

-- View: top credores
CREATE OR REPLACE VIEW vw_top_credores AS
SELECT
    importacao_id,
    credor,
    ROUND(SUM(despesas_empenhadas)::numeric, 2) AS empenhado,
    ROUND(SUM(despesas_liquidadas)::numeric, 2) AS liquidado,
    ROUND(SUM(despesas_pagas)::numeric, 2)      AS pago,
    COUNT(DISTINCT nota_empenho)                 AS qtd_empenhos
FROM fato_execucao
GROUP BY importacao_id, credor
ORDER BY empenhado DESC;

-- View: contratos
CREATE OR REPLACE VIEW vw_contratos AS
SELECT
    importacao_id,
    contrato_siafe,
    contrato_original,
    credor,
    ROUND(SUM(despesas_empenhadas)::numeric, 2) AS empenhado,
    ROUND(SUM(despesas_liquidadas)::numeric, 2) AS liquidado,
    ROUND(SUM(despesas_pagas)::numeric, 2)      AS pago,
    ROUND(SUM(despesas_a_pagar)::numeric, 2)    AS a_pagar
FROM fato_execucao
WHERE contrato_siafe IS NOT NULL AND contrato_siafe != ''
GROUP BY importacao_id, contrato_siafe, contrato_original, credor
ORDER BY empenhado DESC;

-- Habilitar Row Level Security (RLS) — acesso público leitura
ALTER TABLE importacoes   ENABLE ROW LEVEL SECURITY;
ALTER TABLE fato_execucao ENABLE ROW LEVEL SECURITY;

CREATE POLICY "Leitura publica" ON importacoes
    FOR SELECT USING (true);

CREATE POLICY "Leitura publica" ON fato_execucao
    FOR SELECT USING (true);

CREATE POLICY "Insercao autenticada" ON importacoes
    FOR INSERT WITH CHECK (true);

CREATE POLICY "Insercao autenticada" ON fato_execucao
    FOR INSERT WITH CHECK (true);

-- Mensagem de confirmação
SELECT 'Tabelas criadas com sucesso!' AS status;
