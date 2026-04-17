USE lugest;

CREATE TABLE IF NOT EXISTS notas_encomenda (
    numero VARCHAR(30) PRIMARY KEY,
    fornecedor_id VARCHAR(20) NULL,
    data_entrega DATE NULL,
    estado VARCHAR(50) NULL,
    total DECIMAL(12,2) NULL
);

CREATE TABLE IF NOT EXISTS notas_encomenda_linhas (
    id INT AUTO_INCREMENT PRIMARY KEY,
    ne_numero VARCHAR(30) NULL,
    ref_material VARCHAR(20) NULL,
    qtd DECIMAL(10,2) NULL,
    preco DECIMAL(10,2) NULL,
    total DECIMAL(12,2) NULL,
    entregue BOOLEAN NULL,
    lote_fornecedor VARCHAR(100) NULL
);

CREATE TABLE IF NOT EXISTS notas_encomenda_entregas (
    id INT AUTO_INCREMENT PRIMARY KEY,
    ne_numero VARCHAR(30) NOT NULL,
    data_registo DATETIME NULL,
    data_entrega DATE NULL,
    data_documento DATE NULL,
    guia VARCHAR(60) NULL,
    fatura VARCHAR(60) NULL,
    obs TEXT NULL
);

CREATE TABLE IF NOT EXISTS notas_encomenda_linha_entregas (
    id INT AUTO_INCREMENT PRIMARY KEY,
    ne_numero VARCHAR(30) NOT NULL,
    linha_ordem INT NOT NULL,
    data_registo DATETIME NULL,
    data_entrega DATE NULL,
    data_documento DATE NULL,
    guia VARCHAR(60) NULL,
    fatura VARCHAR(60) NULL,
    obs TEXT NULL,
    qtd DECIMAL(10,2) NULL
);

CREATE TABLE IF NOT EXISTS notas_encomenda_documentos (
    id INT AUTO_INCREMENT PRIMARY KEY,
    ne_numero VARCHAR(30) NOT NULL,
    data_registo DATETIME NULL,
    tipo VARCHAR(40) NULL,
    titulo VARCHAR(150) NULL,
    caminho VARCHAR(512) NULL,
    guia VARCHAR(60) NULL,
    fatura VARCHAR(60) NULL,
    data_documento DATE NULL,
    obs TEXT NULL
);

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'contacto'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN contacto VARCHAR(80) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'obs'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN obs TEXT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'local_descarga'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN local_descarga VARCHAR(255) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'meio_transporte'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN meio_transporte VARCHAR(100) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'oculta'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN oculta BOOLEAN NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'is_draft'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN is_draft BOOLEAN NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'data_ultima_entrega'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN data_ultima_entrega DATE NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'guia_ultima'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN guia_ultima VARCHAR(60) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'fatura_ultima'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN fatura_ultima VARCHAR(60) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'fatura_caminho_ultima'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN fatura_caminho_ultima VARCHAR(512) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'data_doc_ultima'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN data_doc_ultima DATE NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'origem_cotacao'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN origem_cotacao VARCHAR(30) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda' AND COLUMN_NAME = 'ne_geradas'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda ADD COLUMN ne_geradas TEXT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'linha_ordem'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN linha_ordem INT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'descricao'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN descricao VARCHAR(255) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'fornecedor_linha'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN fornecedor_linha VARCHAR(150) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'origem'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN origem VARCHAR(50) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'unid'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN unid VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'qtd_entregue'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN qtd_entregue DECIMAL(10,2) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'material'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN material VARCHAR(100) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'espessura'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN espessura VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'comprimento'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN comprimento DECIMAL(10,2) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'largura'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN largura DECIMAL(10,2) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'metros'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN metros DECIMAL(10,2) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'localizacao'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN localizacao VARCHAR(100) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'peso_unid'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN peso_unid DECIMAL(10,3) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'p_compra'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN p_compra DECIMAL(10,4) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'formato'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN formato VARCHAR(50) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'stock_in'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN stock_in BOOLEAN NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'guia_entrega'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN guia_entrega VARCHAR(60) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'fatura_entrega'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN fatura_entrega VARCHAR(60) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'data_doc_entrega'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN data_doc_entrega DATE NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'data_entrega_real'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN data_entrega_real DATE NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND COLUMN_NAME = 'obs_entrega'
);
SET @sql := IF(@c = 0, 'ALTER TABLE notas_encomenda_linhas ADD COLUMN obs_entrega TEXT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linhas' AND INDEX_NAME = 'idx_ne_linhas_num_ord'
);
SET @sql := IF(@i = 0, 'ALTER TABLE notas_encomenda_linhas ADD INDEX idx_ne_linhas_num_ord (ne_numero, linha_ordem)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_entregas' AND INDEX_NAME = 'idx_ne_entregas_num'
);
SET @sql := IF(@i = 0, 'ALTER TABLE notas_encomenda_entregas ADD INDEX idx_ne_entregas_num (ne_numero)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_linha_entregas' AND INDEX_NAME = 'idx_ne_linha_entregas_num_ord'
);
SET @sql := IF(@i = 0, 'ALTER TABLE notas_encomenda_linha_entregas ADD INDEX idx_ne_linha_entregas_num_ord (ne_numero, linha_ordem)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'notas_encomenda_documentos' AND INDEX_NAME = 'idx_ne_docs_num'
);
SET @sql := IF(@i = 0, 'ALTER TABLE notas_encomenda_documentos ADD INDEX idx_ne_docs_num (ne_numero)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @fk := (
    SELECT COUNT(*)
    FROM information_schema.KEY_COLUMN_USAGE
    WHERE TABLE_SCHEMA = DATABASE()
      AND TABLE_NAME = 'notas_encomenda'
      AND COLUMN_NAME = 'fornecedor_id'
      AND REFERENCED_TABLE_NAME = 'fornecedores'
      AND REFERENCED_COLUMN_NAME = 'id'
);
SET @sql := IF(@fk = 0, 'ALTER TABLE notas_encomenda ADD CONSTRAINT fk_notas_encomenda_fornecedor FOREIGN KEY (fornecedor_id) REFERENCES fornecedores(id)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @fk := (
    SELECT COUNT(*)
    FROM information_schema.KEY_COLUMN_USAGE
    WHERE TABLE_SCHEMA = DATABASE()
      AND TABLE_NAME = 'notas_encomenda_linhas'
      AND COLUMN_NAME = 'ne_numero'
      AND REFERENCED_TABLE_NAME = 'notas_encomenda'
      AND REFERENCED_COLUMN_NAME = 'numero'
);
SET @sql := IF(@fk = 0, 'ALTER TABLE notas_encomenda_linhas ADD CONSTRAINT fk_notas_encomenda_linhas_ne FOREIGN KEY (ne_numero) REFERENCES notas_encomenda(numero) ON DELETE CASCADE', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @fk := (
    SELECT COUNT(*)
    FROM information_schema.KEY_COLUMN_USAGE
    WHERE TABLE_SCHEMA = DATABASE()
      AND TABLE_NAME = 'notas_encomenda_entregas'
      AND COLUMN_NAME = 'ne_numero'
      AND REFERENCED_TABLE_NAME = 'notas_encomenda'
      AND REFERENCED_COLUMN_NAME = 'numero'
);
SET @sql := IF(@fk = 0, 'ALTER TABLE notas_encomenda_entregas ADD CONSTRAINT fk_ne_entregas_ne FOREIGN KEY (ne_numero) REFERENCES notas_encomenda(numero) ON DELETE CASCADE', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @fk := (
    SELECT COUNT(*)
    FROM information_schema.KEY_COLUMN_USAGE
    WHERE TABLE_SCHEMA = DATABASE()
      AND TABLE_NAME = 'notas_encomenda_linha_entregas'
      AND COLUMN_NAME = 'ne_numero'
      AND REFERENCED_TABLE_NAME = 'notas_encomenda'
      AND REFERENCED_COLUMN_NAME = 'numero'
);
SET @sql := IF(@fk = 0, 'ALTER TABLE notas_encomenda_linha_entregas ADD CONSTRAINT fk_ne_linha_entregas_ne FOREIGN KEY (ne_numero) REFERENCES notas_encomenda(numero) ON DELETE CASCADE', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @fk := (
    SELECT COUNT(*)
    FROM information_schema.KEY_COLUMN_USAGE
    WHERE TABLE_SCHEMA = DATABASE()
      AND TABLE_NAME = 'notas_encomenda_documentos'
      AND COLUMN_NAME = 'ne_numero'
      AND REFERENCED_TABLE_NAME = 'notas_encomenda'
      AND REFERENCED_COLUMN_NAME = 'numero'
);
SET @sql := IF(@fk = 0, 'ALTER TABLE notas_encomenda_documentos ADD CONSTRAINT fk_ne_documentos_ne FOREIGN KEY (ne_numero) REFERENCES notas_encomenda(numero) ON DELETE CASCADE', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;
