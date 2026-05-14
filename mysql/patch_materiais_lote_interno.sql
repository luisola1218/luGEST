SET @lugest_col_exists := (
  SELECT COUNT(*)
  FROM information_schema.COLUMNS
  WHERE TABLE_SCHEMA = DATABASE()
    AND TABLE_NAME = 'materiais'
    AND COLUMN_NAME = 'lote_interno'
);

SET @lugest_sql := IF(
  @lugest_col_exists = 0,
  'ALTER TABLE `materiais` ADD COLUMN `lote_interno` VARCHAR(100) NULL AFTER `id`',
  'SELECT ''Coluna lote_interno ja existe'' AS estado'
);

PREPARE lugest_stmt FROM @lugest_sql;
EXECUTE lugest_stmt;
DEALLOCATE PREPARE lugest_stmt;
