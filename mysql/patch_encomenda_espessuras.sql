USE lugest;

CREATE TABLE IF NOT EXISTS encomenda_espessuras (
    id INT AUTO_INCREMENT PRIMARY KEY,
    encomenda_numero VARCHAR(30) NOT NULL,
    material VARCHAR(100) NOT NULL,
    espessura VARCHAR(20) NOT NULL,
    tempo_min DECIMAL(10,2),
    estado VARCHAR(50),
    inicio_producao DATETIME,
    fim_producao DATETIME,
    tempo_producao_min DECIMAL(10,2),
    lote_baixa VARCHAR(100),
    UNIQUE KEY uq_enc_esp (encomenda_numero, material, espessura),
    INDEX idx_enc_esp_num (encomenda_numero)
);

SET @fk_exists := (
    SELECT COUNT(*)
    FROM information_schema.TABLE_CONSTRAINTS
    WHERE TABLE_SCHEMA = DATABASE()
      AND TABLE_NAME = 'encomenda_espessuras'
      AND CONSTRAINT_TYPE = 'FOREIGN KEY'
      AND CONSTRAINT_NAME = 'fk_enc_esp_encomenda'
);

SET @sql := IF(
    @fk_exists > 0,
    'SELECT 1',
    'ALTER TABLE encomenda_espessuras ADD CONSTRAINT fk_enc_esp_encomenda FOREIGN KEY (encomenda_numero) REFERENCES encomendas(numero) ON DELETE CASCADE'
);
PREPARE stmt FROM @sql;
EXECUTE stmt;
DEALLOCATE PREPARE stmt;
