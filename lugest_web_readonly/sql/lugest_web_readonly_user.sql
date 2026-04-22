-- Exemplo para criar um utilizador MySQL apenas de leitura para o portal web.
-- Ajustar password, host e nome da base antes de executar.

CREATE USER IF NOT EXISTS 'lugest_web_reader'@'localhost' IDENTIFIED BY 'ALTERAR_PASSWORD_FORTE';

GRANT SELECT ON `lugest`.* TO 'lugest_web_reader'@'localhost';

FLUSH PRIVILEGES;
