USE lugest;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'fornecedores' AND COLUMN_NAME = 'codigo_postal'
);
SET @sql := IF(@c = 0, 'ALTER TABLE fornecedores ADD COLUMN codigo_postal VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'fornecedores' AND COLUMN_NAME = 'localidade'
);
SET @sql := IF(@c = 0, 'ALTER TABLE fornecedores ADD COLUMN localidade VARCHAR(120) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'fornecedores' AND COLUMN_NAME = 'pais'
);
SET @sql := IF(@c = 0, 'ALTER TABLE fornecedores ADD COLUMN pais VARCHAR(80) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'fornecedores' AND COLUMN_NAME = 'cond_pagamento'
);
SET @sql := IF(@c = 0, 'ALTER TABLE fornecedores ADD COLUMN cond_pagamento VARCHAR(120) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'fornecedores' AND COLUMN_NAME = 'prazo_entrega_dias'
);
SET @sql := IF(@c = 0, 'ALTER TABLE fornecedores ADD COLUMN prazo_entrega_dias INT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'fornecedores' AND COLUMN_NAME = 'website'
);
SET @sql := IF(@c = 0, 'ALTER TABLE fornecedores ADD COLUMN website VARCHAR(255) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'fornecedores' AND COLUMN_NAME = 'obs'
);
SET @sql := IF(@c = 0, 'ALTER TABLE fornecedores ADD COLUMN obs TEXT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'fornecedores' AND INDEX_NAME = 'idx_fornecedores_nome'
);
SET @sql := IF(@i = 0, 'ALTER TABLE fornecedores ADD INDEX idx_fornecedores_nome (nome)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'fornecedores' AND INDEX_NAME = 'idx_fornecedores_nif'
);
SET @sql := IF(@i = 0, 'ALTER TABLE fornecedores ADD INDEX idx_fornecedores_nif (nif)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;
