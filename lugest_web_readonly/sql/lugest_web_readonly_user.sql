-- Exemplo para criar um utilizador MySQL apenas de leitura para o portal web.
-- Ajustar password, host e nome da base antes de executar.

CREATE USER IF NOT EXISTS 'lugest_web_reader'@'localhost' IDENTIFIED BY 'ALTERAR_PASSWORD_FORTE';

GRANT SELECT ON `lugest`.* TO 'lugest_web_reader'@'localhost';

FLUSH PRIVILEGES;
A tabela ficou bem, porem se reparares no cartao do planeamento de operação temos fontes cortadas "trocar operação " e no cartões inferiores blocos ativos encomendas carga semanal blocos fechados, tb vê com muita dificuldade. o que sugiro e manter o cartao pendentes e quadro semanal como está reduzindo só 15% de altura dos cartões e moldando tudo o resto eu acho que assim corrigiria o problema.