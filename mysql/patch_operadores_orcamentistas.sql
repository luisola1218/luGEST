USE lugest;

CREATE TABLE IF NOT EXISTS operadores (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nome VARCHAR(120) NOT NULL,
    UNIQUE KEY uq_operadores_nome (nome)
);

CREATE TABLE IF NOT EXISTS orcamentistas (
    id INT AUTO_INCREMENT PRIMARY KEY,
    nome VARCHAR(120) NOT NULL,
    UNIQUE KEY uq_orcamentistas_nome (nome)
);

SET @sql := IF (
    EXISTS (
        SELECT 1
        FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = 'orcamentos'
          AND COLUMN_NAME = 'executado_por'
    ),
    'SELECT 1',
    'ALTER TABLE orcamentos ADD COLUMN executado_por VARCHAR(120) NULL'
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @sql := IF (
    EXISTS (
        SELECT 1
        FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = 'orcamentos'
          AND COLUMN_NAME = 'nota_transporte'
    ),
    'SELECT 1',
    'ALTER TABLE orcamentos ADD COLUMN nota_transporte TEXT NULL'
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @sql := IF (
    EXISTS (
        SELECT 1
        FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = 'orcamentos'
          AND COLUMN_NAME = 'notas_pdf'
    ),
    'SELECT 1',
    'ALTER TABLE orcamentos ADD COLUMN notas_pdf TEXT NULL'
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;
