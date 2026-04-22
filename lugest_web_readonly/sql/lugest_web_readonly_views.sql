-- Views opcionais para simplificar a camada de consulta do portal web.
-- Podem ser usadas para reduzir a exposição direta das tabelas de origem.

CREATE OR REPLACE VIEW vw_web_orders AS
SELECT
    e.numero,
    e.cliente_codigo,
    COALESCE(c.nome, e.cliente_codigo, '-') AS cliente_nome,
    e.estado,
    e.numero_orcamento,
    e.data_entrega,
    e.data_criacao,
    e.tempo_estimado,
    e.ano,
    COALESCE(p.pecas_total, 0) AS pecas_total,
    COALESCE(pl.blocos_total, 0) AS blocos_total
FROM encomendas e
LEFT JOIN clientes c ON c.codigo = e.cliente_codigo
LEFT JOIN (
    SELECT encomenda_numero, COUNT(*) AS pecas_total
    FROM pecas
    GROUP BY encomenda_numero
) p ON p.encomenda_numero = e.numero
LEFT JOIN (
    SELECT encomenda_numero, COUNT(*) AS blocos_total
    FROM plano
    GROUP BY encomenda_numero
) pl ON pl.encomenda_numero = e.numero;

CREATE OR REPLACE VIEW vw_web_planning AS
SELECT
    bloco_id,
    encomenda_numero,
    material,
    espessura,
    data_planeada,
    inicio,
    duracao_min,
    ano
FROM plano;

CREATE OR REPLACE VIEW vw_web_products AS
SELECT
    codigo,
    descricao,
    categoria,
    subcat,
    tipo,
    unid,
    qty,
    alerta,
    p_compra,
    atualizado_em
FROM produtos;

CREATE OR REPLACE VIEW vw_web_materials AS
SELECT
    id,
    lote_fornecedor,
    formato,
    material,
    espessura,
    quantidade,
    reservado,
    localizacao,
    atualizado_em
FROM materiais;

CREATE OR REPLACE VIEW vw_web_quotes AS
SELECT
    o.numero,
    o.data,
    o.estado,
    o.cliente_codigo,
    COALESCE(c.nome, o.cliente_codigo, '-') AS cliente_nome,
    o.subtotal,
    o.total,
    o.numero_encomenda,
    o.ano
FROM orcamentos o
LEFT JOIN clientes c ON c.codigo = o.cliente_codigo;
