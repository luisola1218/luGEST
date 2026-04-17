USE lugest;

SET @sql := IF (
    EXISTS (
        SELECT 1
        FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = 'encomendas'
          AND COLUMN_NAME = 'data_entrega'
    ),
    'SELECT 1',
    'ALTER TABLE encomendas ADD COLUMN data_entrega DATE NULL'
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @sql := IF (
    EXISTS (
        SELECT 1
        FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = 'encomendas'
          AND COLUMN_NAME = 'tempo_estimado'
    ),
    'SELECT 1',
    'ALTER TABLE encomendas ADD COLUMN tempo_estimado DECIMAL(10,2) NULL'
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @sql := IF (
    EXISTS (
        SELECT 1
        FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = 'encomendas'
          AND COLUMN_NAME = 'cativar'
    ),
    'SELECT 1',
    'ALTER TABLE encomendas ADD COLUMN cativar BOOLEAN NULL'
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;

SET @sql := IF (
    EXISTS (
        SELECT 1
        FROM information_schema.COLUMNS
        WHERE TABLE_SCHEMA = DATABASE()
          AND TABLE_NAME = 'encomendas'
          AND COLUMN_NAME = 'observacoes'
    ),
    'SELECT 1',
    'ALTER TABLE encomendas ADD COLUMN observacoes TEXT NULL'
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;
