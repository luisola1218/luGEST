USE lugest;

CREATE TABLE IF NOT EXISTS at_series (
    id INT AUTO_INCREMENT PRIMARY KEY,
    doc_type VARCHAR(10) NOT NULL,
    serie_id VARCHAR(40) NOT NULL,
    inicio_sequencia INT NOT NULL DEFAULT 1,
    next_seq INT NOT NULL DEFAULT 1,
    data_inicio_prevista DATE NULL,
    validation_code VARCHAR(40) NULL,
    status VARCHAR(20) NULL,
    last_error TEXT NULL,
    last_sent_payload_hash VARCHAR(64) NULL,
    updated_at DATETIME NULL,
    UNIQUE KEY uq_at_series_doc_serie (doc_type, serie_id)
);

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'expedicoes' AND COLUMN_NAME = 'serie_id'
);
SET @sql := IF(@c = 0, 'ALTER TABLE expedicoes ADD COLUMN serie_id VARCHAR(40) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'expedicoes' AND COLUMN_NAME = 'seq_num'
);
SET @sql := IF(@c = 0, 'ALTER TABLE expedicoes ADD COLUMN seq_num INT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'expedicoes' AND COLUMN_NAME = 'at_validation_code'
);
SET @sql := IF(@c = 0, 'ALTER TABLE expedicoes ADD COLUMN at_validation_code VARCHAR(40) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'expedicoes' AND COLUMN_NAME = 'atcud'
);
SET @sql := IF(@c = 0, 'ALTER TABLE expedicoes ADD COLUMN atcud VARCHAR(120) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'at_series' AND COLUMN_NAME = 'doc_type'
);
SET @sql := IF(@c = 0, 'ALTER TABLE at_series ADD COLUMN doc_type VARCHAR(10) NOT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'at_series' AND COLUMN_NAME = 'serie_id'
);
SET @sql := IF(@c = 0, 'ALTER TABLE at_series ADD COLUMN serie_id VARCHAR(40) NOT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'at_series' AND COLUMN_NAME = 'inicio_sequencia'
);
SET @sql := IF(@c = 0, 'ALTER TABLE at_series ADD COLUMN inicio_sequencia INT NOT NULL DEFAULT 1', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'at_series' AND COLUMN_NAME = 'next_seq'
);
SET @sql := IF(@c = 0, 'ALTER TABLE at_series ADD COLUMN next_seq INT NOT NULL DEFAULT 1', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'at_series' AND COLUMN_NAME = 'data_inicio_prevista'
);
SET @sql := IF(@c = 0, 'ALTER TABLE at_series ADD COLUMN data_inicio_prevista DATE NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'at_series' AND COLUMN_NAME = 'validation_code'
);
SET @sql := IF(@c = 0, 'ALTER TABLE at_series ADD COLUMN validation_code VARCHAR(40) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'at_series' AND COLUMN_NAME = 'status'
);
SET @sql := IF(@c = 0, 'ALTER TABLE at_series ADD COLUMN status VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'at_series' AND COLUMN_NAME = 'last_error'
);
SET @sql := IF(@c = 0, 'ALTER TABLE at_series ADD COLUMN last_error TEXT NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'at_series' AND COLUMN_NAME = 'last_sent_payload_hash'
);
SET @sql := IF(@c = 0, 'ALTER TABLE at_series ADD COLUMN last_sent_payload_hash VARCHAR(64) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'at_series' AND COLUMN_NAME = 'updated_at'
);
SET @sql := IF(@c = 0, 'ALTER TABLE at_series ADD COLUMN updated_at DATETIME NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'expedicoes' AND INDEX_NAME = 'idx_expedicoes_serie_seq'
);
SET @sql := IF(@i = 0, 'ALTER TABLE expedicoes ADD INDEX idx_expedicoes_serie_seq (serie_id, seq_num)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'expedicoes' AND INDEX_NAME = 'idx_expedicoes_atcud'
);
SET @sql := IF(@i = 0, 'ALTER TABLE expedicoes ADD INDEX idx_expedicoes_atcud (atcud)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @i := (
    SELECT COUNT(*)
    FROM information_schema.STATISTICS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'at_series' AND INDEX_NAME = 'idx_at_series_doc_serie'
);
SET @sql := IF(@i = 0, 'ALTER TABLE at_series ADD INDEX idx_at_series_doc_serie (doc_type, serie_id)', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;
