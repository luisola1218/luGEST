USE lugest;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'op_paragens' AND COLUMN_NAME = 'fechada_at'
);
SET @sql := IF(@c = 0, 'ALTER TABLE op_paragens ADD COLUMN fechada_at DATETIME NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'op_paragens' AND COLUMN_NAME = 'origem'
);
SET @sql := IF(@c = 0, 'ALTER TABLE op_paragens ADD COLUMN origem VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

SET @c := (
    SELECT COUNT(*)
    FROM information_schema.COLUMNS
    WHERE TABLE_SCHEMA = DATABASE() AND TABLE_NAME = 'op_paragens' AND COLUMN_NAME = 'estado'
);
SET @sql := IF(@c = 0, 'ALTER TABLE op_paragens ADD COLUMN estado VARCHAR(20) NULL', 'SELECT 1');
PREPARE stmt FROM @sql; EXECUTE stmt; DEALLOCATE PREPARE stmt;

UPDATE op_paragens
SET origem = CASE
        WHEN COALESCE(NULLIF(origem, ''), '') <> '' THEN origem
        WHEN LOWER(COALESCE(causa, '')) LIKE '%avaria%' THEN 'AVARIA'
        ELSE NULL
    END,
    estado = CASE
        WHEN COALESCE(duracao_min, 0) > 0 THEN 'FECHADA'
        ELSE COALESCE(NULLIF(estado, ''), 'ABERTA')
    END
WHERE 1=1;
