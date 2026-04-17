USE lugest;

CREATE TABLE IF NOT EXISTS conjuntos_modelo (
    codigo VARCHAR(40) PRIMARY KEY,
    descricao VARCHAR(150) NOT NULL,
    notas TEXT NULL,
    ativo BOOLEAN NULL,
    created_at DATETIME NULL,
    updated_at DATETIME NULL
);

CREATE TABLE IF NOT EXISTS conjuntos_modelo_itens (
    id INT AUTO_INCREMENT PRIMARY KEY,
    conjunto_codigo VARCHAR(40) NOT NULL,
    linha_ordem INT NULL,
    tipo_item VARCHAR(30) NULL,
    ref_externa VARCHAR(100) NULL,
    descricao TEXT NULL,
    material VARCHAR(100) NULL,
    espessura VARCHAR(20) NULL,
    operacao VARCHAR(150) NULL,
    produto_codigo VARCHAR(20) NULL,
    produto_unid VARCHAR(20) NULL,
    qtd DECIMAL(10,2) NULL,
    tempo_peca_min DECIMAL(10,2) NULL,
    preco_unit DECIMAL(10,4) NULL,
    desenho_path VARCHAR(512) NULL,
    INDEX idx_conjuntos_itens_codigo_ord (conjunto_codigo, linha_ordem),
    FOREIGN KEY (conjunto_codigo) REFERENCES conjuntos_modelo(codigo) ON DELETE CASCADE
);

CREATE TABLE IF NOT EXISTS encomenda_montagem_itens (
    id INT AUTO_INCREMENT PRIMARY KEY,
    encomenda_numero VARCHAR(30) NOT NULL,
    linha_ordem INT NULL,
    tipo_item VARCHAR(30) NULL,
    descricao TEXT NULL,
    produto_codigo VARCHAR(20) NULL,
    produto_unid VARCHAR(20) NULL,
    qtd_planeada DECIMAL(10,2) NULL,
    qtd_consumida DECIMAL(10,2) NULL,
    preco_unit DECIMAL(10,4) NULL,
    conjunto_codigo VARCHAR(40) NULL,
    conjunto_nome VARCHAR(150) NULL,
    grupo_uuid VARCHAR(60) NULL,
    estado VARCHAR(30) NULL,
    obs TEXT NULL,
    created_at DATETIME NULL,
    consumed_at DATETIME NULL,
    consumed_by VARCHAR(120) NULL,
    INDEX idx_enc_montagem_num_ord (encomenda_numero, linha_ordem),
    INDEX idx_enc_montagem_estado (estado),
    FOREIGN KEY (encomenda_numero) REFERENCES encomendas(numero) ON DELETE CASCADE
);

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'orcamento_linhas' AND COLUMN_NAME = 'tempo_peca_min'
);
SET @sql := IF(@c = 0, 'ALTER TABLE orcamento_linhas ADD COLUMN tempo_peca_min DECIMAL(10,2) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'orcamento_linhas' AND COLUMN_NAME = 'tipo_item'
);
SET @sql := IF(@c = 0, 'ALTER TABLE orcamento_linhas ADD COLUMN tipo_item VARCHAR(30) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'orcamento_linhas' AND COLUMN_NAME = 'produto_codigo'
);
SET @sql := IF(@c = 0, 'ALTER TABLE orcamento_linhas ADD COLUMN produto_codigo VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'orcamento_linhas' AND COLUMN_NAME = 'produto_unid'
);
SET @sql := IF(@c = 0, 'ALTER TABLE orcamento_linhas ADD COLUMN produto_unid VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'orcamento_linhas' AND COLUMN_NAME = 'conjunto_codigo'
);
SET @sql := IF(@c = 0, 'ALTER TABLE orcamento_linhas ADD COLUMN conjunto_codigo VARCHAR(40) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'orcamento_linhas' AND COLUMN_NAME = 'conjunto_nome'
);
SET @sql := IF(@c = 0, 'ALTER TABLE orcamento_linhas ADD COLUMN conjunto_nome VARCHAR(150) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'orcamento_linhas' AND COLUMN_NAME = 'grupo_uuid'
);
SET @sql := IF(@c = 0, 'ALTER TABLE orcamento_linhas ADD COLUMN grupo_uuid VARCHAR(60) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'orcamento_linhas' AND COLUMN_NAME = 'qtd_base'
);
SET @sql := IF(@c = 0, 'ALTER TABLE orcamento_linhas ADD COLUMN qtd_base DECIMAL(10,2) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo' AND COLUMN_NAME = 'notas'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo ADD COLUMN notas TEXT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo' AND COLUMN_NAME = 'ativo'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo ADD COLUMN ativo BOOLEAN NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo' AND COLUMN_NAME = 'created_at'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo ADD COLUMN created_at DATETIME NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo' AND COLUMN_NAME = 'updated_at'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo ADD COLUMN updated_at DATETIME NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'linha_ordem'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN linha_ordem INT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'tipo_item'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN tipo_item VARCHAR(30) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'ref_externa'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN ref_externa VARCHAR(100) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'descricao'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN descricao TEXT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'material'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN material VARCHAR(100) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'espessura'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN espessura VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'operacao'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN operacao VARCHAR(150) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'produto_codigo'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN produto_codigo VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'produto_unid'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN produto_unid VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'qtd'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN qtd DECIMAL(10,2) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'tempo_peca_min'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN tempo_peca_min DECIMAL(10,2) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'preco_unit'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN preco_unit DECIMAL(10,4) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND COLUMN_NAME = 'desenho_path'
);
SET @sql := IF(@c = 0, 'ALTER TABLE conjuntos_modelo_itens ADD COLUMN desenho_path VARCHAR(512) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'conjuntos_modelo_itens' AND INDEX_NAME = 'idx_conjuntos_itens_codigo_ord'
);
SET @sql := IF(@i = 0, 'ALTER TABLE conjuntos_modelo_itens ADD INDEX idx_conjuntos_itens_codigo_ord (conjunto_codigo, linha_ordem)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'linha_ordem'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN linha_ordem INT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'tipo_item'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN tipo_item VARCHAR(30) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'descricao'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN descricao TEXT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'produto_codigo'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN produto_codigo VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'produto_unid'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN produto_unid VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'qtd_planeada'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN qtd_planeada DECIMAL(10,2) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'qtd_consumida'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN qtd_consumida DECIMAL(10,2) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'preco_unit'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN preco_unit DECIMAL(10,4) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'conjunto_codigo'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN conjunto_codigo VARCHAR(40) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'conjunto_nome'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN conjunto_nome VARCHAR(150) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'grupo_uuid'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN grupo_uuid VARCHAR(60) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'estado'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN estado VARCHAR(30) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'obs'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN obs TEXT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'created_at'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN created_at DATETIME NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'consumed_at'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN consumed_at DATETIME NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND COLUMN_NAME = 'consumed_by'
);
SET @sql := IF(@c = 0, 'ALTER TABLE encomenda_montagem_itens ADD COLUMN consumed_by VARCHAR(120) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND INDEX_NAME = 'idx_enc_montagem_num_ord'
);
SET @sql := IF(@i = 0, 'ALTER TABLE encomenda_montagem_itens ADD INDEX idx_enc_montagem_num_ord (encomenda_numero, linha_ordem)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'encomenda_montagem_itens' AND INDEX_NAME = 'idx_enc_montagem_estado'
);
SET @sql := IF(@i = 0, 'ALTER TABLE encomenda_montagem_itens ADD INDEX idx_enc_montagem_estado (estado)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;
