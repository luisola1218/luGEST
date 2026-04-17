USE lugest;

CREATE TABLE IF NOT EXISTS produtos_mov (
    id INT AUTO_INCREMENT PRIMARY KEY,
    data DATETIME NULL,
    tipo VARCHAR(40) NULL,
    operador VARCHAR(120) NULL,
    codigo VARCHAR(20) NULL,
    descricao VARCHAR(255) NULL,
    qtd DECIMAL(10,2) NULL,
    antes DECIMAL(10,2) NULL,
    depois DECIMAL(10,2) NULL,
    obs TEXT NULL,
    origem VARCHAR(80) NULL,
    ref_doc VARCHAR(50) NULL,
    INDEX idx_prod_mov_data (data),
    INDEX idx_prod_mov_operador (operador),
    INDEX idx_prod_mov_codigo (codigo)
);
