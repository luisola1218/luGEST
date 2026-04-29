-- =====================================================
-- LuGEST CURRENT FULL SCHEMA
-- Gerado a partir da base MySQL atual
-- =====================================================

-- Opcional: para reiniciar tudo, descomenta a linha seguinte.
-- DROP DATABASE IF EXISTS `lugest`;
CREATE DATABASE IF NOT EXISTS `lugest` CHARACTER SET utf8 COLLATE utf8_general_ci;
USE `lugest`;
SET NAMES utf8mb4;
SET FOREIGN_KEY_CHECKS=0;

-- =====================================================
-- TABELAS
-- =====================================================

CREATE TABLE IF NOT EXISTS `app_config` (
  `ckey` varchar(80) NOT NULL,
  `cvalue` longtext,
  `updated_at` datetime DEFAULT NULL,
  PRIMARY KEY (`ckey`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `app_state` (
  `id` tinyint(4) NOT NULL,
  `data_json` longtext COLLATE utf8mb4_unicode_ci NOT NULL,
  `updated_at` timestamp NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4 COLLATE=utf8mb4_unicode_ci;

CREATE TABLE IF NOT EXISTS `at_series` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `doc_type` varchar(10) NOT NULL,
  `serie_id` varchar(40) NOT NULL,
  `inicio_sequencia` int(11) NOT NULL DEFAULT '1',
  `next_seq` int(11) NOT NULL DEFAULT '1',
  `data_inicio_prevista` date DEFAULT NULL,
  `validation_code` varchar(40) DEFAULT NULL,
  `status` varchar(20) DEFAULT NULL,
  `last_error` text,
  `last_sent_payload_hash` varchar(64) DEFAULT NULL,
  `updated_at` datetime DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `uq_at_series_doc_serie` (`doc_type`,`serie_id`),
  KEY `idx_at_series_doc_serie` (`doc_type`,`serie_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `clientes` (
  `codigo` varchar(20) NOT NULL,
  `nome` varchar(150) DEFAULT NULL,
  `nif` varchar(20) DEFAULT NULL,
  `morada` varchar(255) DEFAULT NULL,
  `contacto` varchar(50) DEFAULT NULL,
  `email` varchar(150) DEFAULT NULL,
  PRIMARY KEY (`codigo`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `conjuntos_modelo` (
  `codigo` varchar(40) NOT NULL,
  `descricao` varchar(150) NOT NULL,
  `notas` text,
  `ativo` tinyint(1) DEFAULT NULL,
  `created_at` datetime DEFAULT NULL,
  `updated_at` datetime DEFAULT NULL,
  PRIMARY KEY (`codigo`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `expedicoes` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `numero` varchar(30) DEFAULT NULL,
  `tipo` varchar(30) DEFAULT NULL,
  `encomenda_numero` varchar(30) DEFAULT NULL,
  `cliente_codigo` varchar(20) DEFAULT NULL,
  `cliente_nome` varchar(150) DEFAULT NULL,
  `destinatario` varchar(150) DEFAULT NULL,
  `dest_nif` varchar(20) DEFAULT NULL,
  `dest_morada` varchar(255) DEFAULT NULL,
  `local_carga` varchar(255) DEFAULT NULL,
  `local_descarga` varchar(255) DEFAULT NULL,
  `data_emissao` datetime DEFAULT NULL,
  `data_transporte` datetime DEFAULT NULL,
  `matricula` varchar(30) DEFAULT NULL,
  `transportador` varchar(150) DEFAULT NULL,
  `estado` varchar(50) DEFAULT NULL,
  `observacoes` text,
  `created_by` varchar(80) DEFAULT NULL,
  `anulada` tinyint(1) DEFAULT NULL,
  `anulada_motivo` text,
  `codigo_at` varchar(80) DEFAULT NULL,
  `emitente_nome` varchar(150) DEFAULT NULL,
  `emitente_nif` varchar(20) DEFAULT NULL,
  `emitente_morada` varchar(255) DEFAULT NULL,
  `serie_id` varchar(40) DEFAULT NULL,
  `seq_num` int(11) DEFAULT NULL,
  `at_validation_code` varchar(40) DEFAULT NULL,
  `atcud` varchar(120) DEFAULT NULL,
  `ano` int(11) DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `numero` (`numero`),
  UNIQUE KEY `uq_expedicoes_numero` (`numero`),
  KEY `idx_expedicoes_encomenda_numero` (`encomenda_numero`),
  KEY `idx_expedicoes_cliente_codigo` (`cliente_codigo`),
  KEY `idx_expedicoes_serie_seq` (`serie_id`,`seq_num`),
  KEY `idx_expedicoes_atcud` (`atcud`),
  KEY `idx_expedicoes_ano` (`ano`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `fornecedores` (
  `id` varchar(20) NOT NULL,
  `nome` varchar(150) DEFAULT NULL,
  `nif` varchar(20) DEFAULT NULL,
  `morada` varchar(255) DEFAULT NULL,
  `contacto` varchar(50) DEFAULT NULL,
  `email` varchar(150) DEFAULT NULL,
  `codigo_postal` varchar(20) DEFAULT NULL,
  `localidade` varchar(120) DEFAULT NULL,
  `pais` varchar(80) DEFAULT NULL,
  `cond_pagamento` varchar(120) DEFAULT NULL,
  `prazo_entrega_dias` int(11) DEFAULT NULL,
  `website` varchar(255) DEFAULT NULL,
  `obs` text,
  PRIMARY KEY (`id`),
  KEY `idx_fornecedores_nome` (`nome`),
  KEY `idx_fornecedores_nif` (`nif`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `materiais` (
  `id` varchar(20) NOT NULL,
  `lote_fornecedor` varchar(100) DEFAULT NULL,
  `formato` varchar(50) DEFAULT NULL,
  `material` varchar(100) DEFAULT NULL,
  `espessura` varchar(20) DEFAULT NULL,
  `comprimento` decimal(10,2) DEFAULT NULL,
  `largura` decimal(10,2) DEFAULT NULL,
  `metros` decimal(10,2) DEFAULT NULL,
  `peso_unid` decimal(10,3) DEFAULT NULL,
  `p_compra` decimal(10,6) DEFAULT NULL,
  `quantidade` decimal(10,2) DEFAULT NULL,
  `reservado` decimal(10,2) DEFAULT NULL,
  `tipo` varchar(50) DEFAULT NULL,
  `localizacao` varchar(100) DEFAULT NULL,
  `is_sobra` tinyint(1) DEFAULT NULL,
  `atualizado_em` datetime DEFAULT NULL,
  `preco_unid` decimal(12,4) DEFAULT NULL,
  `logistic_status` varchar(30) DEFAULT NULL,
  `quality_status` varchar(40) DEFAULT NULL,
  `quality_blocked` tinyint(1) DEFAULT NULL,
  `inspection_status` varchar(40) DEFAULT NULL,
  `inspection_defect` varchar(255) DEFAULT NULL,
  `inspection_decision` varchar(255) DEFAULT NULL,
  `inspection_at` datetime DEFAULT NULL,
  `inspection_by` varchar(120) DEFAULT NULL,
  `inspection_note_number` varchar(30) DEFAULT NULL,
  `inspection_supplier_id` varchar(20) DEFAULT NULL,
  `inspection_supplier_name` varchar(150) DEFAULT NULL,
  `inspection_guia` varchar(60) DEFAULT NULL,
  `inspection_fatura` varchar(60) DEFAULT NULL,
  `quality_nc_id` varchar(30) DEFAULT NULL,
  `supplier_claim_id` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `ne_linhas_historico` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `created_at` datetime NOT NULL,
  `evento` varchar(30) NOT NULL,
  `origem_menu` varchar(80) DEFAULT NULL,
  `utilizador` varchar(80) DEFAULT NULL,
  `guia_numero` varchar(30) DEFAULT NULL,
  `produto_codigo` varchar(20) DEFAULT NULL,
  `descricao` varchar(255) DEFAULT NULL,
  `qtd` decimal(10,2) DEFAULT NULL,
  `unid` varchar(20) DEFAULT NULL,
  `destinatario` varchar(150) DEFAULT NULL,
  `observacoes` text,
  `payload_json` longtext,
  PRIMARY KEY (`id`),
  KEY `idx_ne_linhas_hist_created_at` (`created_at`),
  KEY `idx_ne_linhas_hist_evento` (`evento`),
  KEY `idx_ne_linhas_hist_guia` (`guia_numero`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `op_eventos` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `created_at` datetime NOT NULL,
  `evento` varchar(30) NOT NULL,
  `encomenda_numero` varchar(30) DEFAULT NULL,
  `peca_id` varchar(30) DEFAULT NULL,
  `ref_interna` varchar(60) DEFAULT NULL,
  `material` varchar(100) DEFAULT NULL,
  `espessura` varchar(20) DEFAULT NULL,
  `operacao` varchar(80) DEFAULT NULL,
  `operador` varchar(80) DEFAULT NULL,
  `qtd_ok` decimal(10,2) DEFAULT NULL,
  `qtd_nok` decimal(10,2) DEFAULT NULL,
  `info` text,
  PRIMARY KEY (`id`),
  KEY `idx_op_eventos_created_at` (`created_at`),
  KEY `idx_op_eventos_evento` (`evento`),
  KEY `idx_op_eventos_enc` (`encomenda_numero`),
  KEY `idx_op_eventos_peca` (`peca_id`),
  KEY `idx_op_eventos_operador` (`operador`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `op_paragens` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `created_at` datetime NOT NULL,
  `encomenda_numero` varchar(30) DEFAULT NULL,
  `peca_id` varchar(30) DEFAULT NULL,
  `ref_interna` varchar(60) DEFAULT NULL,
  `material` varchar(100) DEFAULT NULL,
  `espessura` varchar(20) DEFAULT NULL,
  `operador` varchar(80) DEFAULT NULL,
  `causa` varchar(120) DEFAULT NULL,
  `detalhe` text,
  `duracao_min` decimal(10,2) DEFAULT NULL,
  `fechada_at` datetime DEFAULT NULL,
  `origem` varchar(20) DEFAULT NULL,
  `estado` varchar(20) DEFAULT NULL,
  `grupo_id` varchar(80) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_op_paragens_created_at` (`created_at`),
  KEY `idx_op_paragens_causa` (`causa`),
  KEY `idx_op_paragens_enc` (`encomenda_numero`),
  KEY `idx_op_paragens_peca` (`peca_id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `operadores` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `nome` varchar(120) NOT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `uq_operadores_nome` (`nome`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `orc_referencias_historico` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `ref_externa` varchar(100) NOT NULL,
  `ref_interna` varchar(50) DEFAULT NULL,
  `descricao` text,
  `material` varchar(100) DEFAULT NULL,
  `espessura` varchar(20) DEFAULT NULL,
  `preco_unit` decimal(10,2) DEFAULT NULL,
  `operacao` varchar(150) DEFAULT NULL,
  `desenho_path` varchar(512) DEFAULT NULL,
  `updated_at` datetime NOT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `uq_orc_ref_externa` (`ref_externa`),
  KEY `idx_orc_ref_interna` (`ref_interna`),
  KEY `idx_orc_ref_updated_at` (`updated_at`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `orcamentistas` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `nome` varchar(120) NOT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `uq_orcamentistas_nome` (`nome`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `peca_operacoes_execucao` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `encomenda_numero` varchar(30) NOT NULL,
  `peca_id` varchar(30) NOT NULL,
  `operacao` varchar(80) NOT NULL,
  `estado` varchar(20) NOT NULL DEFAULT 'Livre',
  `operador_atual` varchar(120) DEFAULT NULL,
  `inicio_ts` datetime DEFAULT NULL,
  `fim_ts` datetime DEFAULT NULL,
  `ok_qty` decimal(10,2) DEFAULT NULL,
  `nok_qty` decimal(10,2) DEFAULT NULL,
  `qual_qty` decimal(10,2) DEFAULT NULL,
  `updated_at` datetime DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `uq_poe_enc_peca_operacao` (`encomenda_numero`,`peca_id`,`operacao`),
  KEY `idx_poe_enc` (`encomenda_numero`),
  KEY `idx_poe_estado` (`estado`),
  KEY `idx_poe_operador` (`operador_atual`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `plano` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `bloco_id` varchar(60) NOT NULL,
  `encomenda_numero` varchar(30) NOT NULL,
  `material` varchar(100) DEFAULT NULL,
  `espessura` varchar(20) DEFAULT NULL,
  `data_planeada` date NOT NULL,
  `inicio` varchar(8) NOT NULL,
  `duracao_min` decimal(10,2) NOT NULL,
  `color` varchar(20) DEFAULT NULL,
  `chapa` varchar(120) DEFAULT NULL,
  `ano` int(11) DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `uq_plano_bloco` (`bloco_id`),
  KEY `idx_plano_data_inicio` (`data_planeada`,`inicio`),
  KEY `idx_plano_enc` (`encomenda_numero`),
  KEY `idx_plano_ano` (`ano`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `plano_hist` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `bloco_id` varchar(60) NOT NULL,
  `encomenda_numero` varchar(30) NOT NULL,
  `material` varchar(100) DEFAULT NULL,
  `espessura` varchar(20) DEFAULT NULL,
  `data_planeada` date NOT NULL,
  `inicio` varchar(8) NOT NULL,
  `duracao_min` decimal(10,2) NOT NULL,
  `color` varchar(20) DEFAULT NULL,
  `chapa` varchar(120) DEFAULT NULL,
  `movido_em` datetime DEFAULT NULL,
  `estado_final` varchar(50) DEFAULT NULL,
  `tempo_planeado_min` decimal(10,2) DEFAULT NULL,
  `tempo_real_min` decimal(10,2) DEFAULT NULL,
  `ano` int(11) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_plano_hist_enc` (`encomenda_numero`),
  KEY `idx_plano_hist_data` (`data_planeada`,`inicio`),
  KEY `idx_plano_hist_ano` (`ano`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `produtos` (
  `codigo` varchar(20) NOT NULL,
  `descricao` varchar(255) DEFAULT NULL,
  `categoria` varchar(100) DEFAULT NULL,
  `subcat` varchar(100) DEFAULT NULL,
  `tipo` varchar(100) DEFAULT NULL,
  `unid` varchar(20) DEFAULT NULL,
  `qty` decimal(10,2) DEFAULT NULL,
  `alerta` decimal(10,2) DEFAULT NULL,
  `p_compra` decimal(10,2) DEFAULT NULL,
  `atualizado_em` datetime DEFAULT NULL,
  `logistic_status` varchar(30) DEFAULT NULL,
  `quality_status` varchar(40) DEFAULT NULL,
  `quality_blocked` tinyint(1) DEFAULT NULL,
  `inspection_defect` varchar(255) DEFAULT NULL,
  `inspection_decision` varchar(255) DEFAULT NULL,
  `inspection_note_number` varchar(30) DEFAULT NULL,
  `quality_nc_id` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`codigo`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `produtos_mov` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `data` datetime DEFAULT NULL,
  `tipo` varchar(40) DEFAULT NULL,
  `operador` varchar(120) DEFAULT NULL,
  `codigo` varchar(20) DEFAULT NULL,
  `descricao` varchar(255) DEFAULT NULL,
  `qtd` decimal(10,2) DEFAULT NULL,
  `antes` decimal(10,2) DEFAULT NULL,
  `depois` decimal(10,2) DEFAULT NULL,
  `obs` text,
  `origem` varchar(80) DEFAULT NULL,
  `ref_doc` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_prod_mov_data` (`data`),
  KEY `idx_prod_mov_operador` (`operador`),
  KEY `idx_prod_mov_codigo` (`codigo`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `stock_log` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `data` datetime DEFAULT NULL,
  `acao` varchar(50) DEFAULT NULL,
  `detalhes` text,
  PRIMARY KEY (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `users` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `username` varchar(50) DEFAULT NULL,
  `password` varchar(255) DEFAULT NULL,
  `role` varchar(50) DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `username` (`username`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `encomendas` (
  `numero` varchar(30) NOT NULL,
  `cliente_codigo` varchar(20) DEFAULT NULL,
  `nota_cliente` varchar(255) DEFAULT NULL,
  `data_criacao` datetime DEFAULT NULL,
  `estado` varchar(50) DEFAULT NULL,
  `numero_orcamento` varchar(30) DEFAULT NULL,
  `data_entrega` date DEFAULT NULL,
  `tempo_estimado` decimal(10,2) DEFAULT NULL,
  `cativar` tinyint(1) DEFAULT NULL,
  `observacoes` text,
  `ano` int(11) DEFAULT NULL,
  PRIMARY KEY (`numero`),
  KEY `cliente_codigo` (`cliente_codigo`),
  KEY `idx_encomendas_estado` (`estado`),
  KEY `idx_encomendas_ano` (`ano`),
  CONSTRAINT `encomendas_ibfk_1` FOREIGN KEY (`cliente_codigo`) REFERENCES `clientes` (`codigo`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `orcamentos` (
  `numero` varchar(30) NOT NULL,
  `data` datetime DEFAULT NULL,
  `estado` varchar(50) DEFAULT NULL,
  `cliente_codigo` varchar(20) DEFAULT NULL,
  `iva_perc` decimal(5,2) DEFAULT NULL,
  `subtotal` decimal(12,2) DEFAULT NULL,
  `total` decimal(12,2) DEFAULT NULL,
  `numero_encomenda` varchar(30) DEFAULT NULL,
  `nota_cliente` text,
  `executado_por` varchar(120) DEFAULT NULL,
  `nota_transporte` text,
  `notas_pdf` text,
  `ano` int(11) DEFAULT NULL,
  `preco_transporte` decimal(12,2) DEFAULT NULL,
  PRIMARY KEY (`numero`),
  KEY `cliente_codigo` (`cliente_codigo`),
  KEY `idx_orcamentos_estado` (`estado`),
  KEY `idx_orcamentos_ano` (`ano`),
  CONSTRAINT `orcamentos_ibfk_1` FOREIGN KEY (`cliente_codigo`) REFERENCES `clientes` (`codigo`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `conjuntos_modelo_itens` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `conjunto_codigo` varchar(40) NOT NULL,
  `linha_ordem` int(11) DEFAULT NULL,
  `tipo_item` varchar(30) DEFAULT NULL,
  `ref_externa` varchar(100) DEFAULT NULL,
  `descricao` text,
  `material` varchar(100) DEFAULT NULL,
  `espessura` varchar(20) DEFAULT NULL,
  `operacao` varchar(150) DEFAULT NULL,
  `produto_codigo` varchar(20) DEFAULT NULL,
  `produto_unid` varchar(20) DEFAULT NULL,
  `qtd` decimal(10,2) DEFAULT NULL,
  `tempo_peca_min` decimal(10,2) DEFAULT NULL,
  `preco_unit` decimal(10,4) DEFAULT NULL,
  `desenho_path` varchar(512) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_conjuntos_itens_codigo_ord` (`conjunto_codigo`,`linha_ordem`),
  CONSTRAINT `conjuntos_modelo_itens_ibfk_1` FOREIGN KEY (`conjunto_codigo`) REFERENCES `conjuntos_modelo` (`codigo`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `expedicao_linhas` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `expedicao_numero` varchar(30) DEFAULT NULL,
  `encomenda_numero` varchar(30) DEFAULT NULL,
  `peca_id` varchar(30) DEFAULT NULL,
  `ref_interna` varchar(50) DEFAULT NULL,
  `ref_externa` varchar(100) DEFAULT NULL,
  `descricao` varchar(255) DEFAULT NULL,
  `qtd` decimal(10,2) DEFAULT NULL,
  `unid` varchar(20) DEFAULT NULL,
  `peso` decimal(10,3) DEFAULT NULL,
  `manual` tinyint(1) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_expedicao_linhas_expedicao_numero` (`expedicao_numero`),
  KEY `idx_expedicao_linhas_encomenda_numero` (`encomenda_numero`),
  KEY `idx_expedicao_linhas_peca_id` (`peca_id`),
  CONSTRAINT `fk_expedicao_linhas_expedicoes` FOREIGN KEY (`expedicao_numero`) REFERENCES `expedicoes` (`numero`) ON DELETE CASCADE,
  CONSTRAINT `expedicao_linhas_ibfk_1` FOREIGN KEY (`expedicao_numero`) REFERENCES `expedicoes` (`numero`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `notas_encomenda` (
  `numero` varchar(30) NOT NULL,
  `fornecedor_id` varchar(20) DEFAULT NULL,
  `data_entrega` date DEFAULT NULL,
  `estado` varchar(50) DEFAULT NULL,
  `total` decimal(12,2) DEFAULT NULL,
  `contacto` varchar(80) DEFAULT NULL,
  `obs` text,
  `local_descarga` varchar(255) DEFAULT NULL,
  `meio_transporte` varchar(100) DEFAULT NULL,
  `oculta` tinyint(1) DEFAULT NULL,
  `is_draft` tinyint(1) DEFAULT NULL,
  `data_ultima_entrega` date DEFAULT NULL,
  `guia_ultima` varchar(60) DEFAULT NULL,
  `fatura_ultima` varchar(60) DEFAULT NULL,
  `data_doc_ultima` date DEFAULT NULL,
  `origem_cotacao` varchar(30) DEFAULT NULL,
  `ne_geradas` text,
  `fatura_caminho_ultima` varchar(512) DEFAULT NULL,
  `ano` int(11) DEFAULT NULL,
  PRIMARY KEY (`numero`),
  KEY `fornecedor_id` (`fornecedor_id`),
  KEY `idx_ne_ano` (`ano`),
  CONSTRAINT `notas_encomenda_ibfk_1` FOREIGN KEY (`fornecedor_id`) REFERENCES `fornecedores` (`id`)
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `encomenda_espessuras` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `encomenda_numero` varchar(30) NOT NULL,
  `material` varchar(100) NOT NULL,
  `espessura` varchar(20) NOT NULL,
  `tempo_min` decimal(10,2) DEFAULT NULL,
  `estado` varchar(50) DEFAULT NULL,
  `inicio_producao` datetime DEFAULT NULL,
  `fim_producao` datetime DEFAULT NULL,
  `tempo_producao_min` decimal(10,2) DEFAULT NULL,
  `lote_baixa` varchar(100) DEFAULT NULL,
  PRIMARY KEY (`id`),
  UNIQUE KEY `uq_enc_esp` (`encomenda_numero`,`material`,`espessura`),
  KEY `idx_enc_esp_num` (`encomenda_numero`),
  CONSTRAINT `fk_enc_esp_encomenda` FOREIGN KEY (`encomenda_numero`) REFERENCES `encomendas` (`numero`) ON DELETE CASCADE,
  CONSTRAINT `encomenda_espessuras_ibfk_1` FOREIGN KEY (`encomenda_numero`) REFERENCES `encomendas` (`numero`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `encomenda_montagem_itens` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `encomenda_numero` varchar(30) NOT NULL,
  `linha_ordem` int(11) DEFAULT NULL,
  `tipo_item` varchar(30) DEFAULT NULL,
  `descricao` text,
  `produto_codigo` varchar(20) DEFAULT NULL,
  `produto_unid` varchar(20) DEFAULT NULL,
  `qtd_planeada` decimal(10,2) DEFAULT NULL,
  `qtd_consumida` decimal(10,2) DEFAULT NULL,
  `preco_unit` decimal(10,4) DEFAULT NULL,
  `conjunto_codigo` varchar(40) DEFAULT NULL,
  `conjunto_nome` varchar(150) DEFAULT NULL,
  `grupo_uuid` varchar(60) DEFAULT NULL,
  `estado` varchar(30) DEFAULT NULL,
  `obs` text,
  `created_at` datetime DEFAULT NULL,
  `consumed_at` datetime DEFAULT NULL,
  `consumed_by` varchar(120) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_enc_montagem_num_ord` (`encomenda_numero`,`linha_ordem`),
  KEY `idx_enc_montagem_estado` (`estado`),
  CONSTRAINT `encomenda_montagem_itens_ibfk_1` FOREIGN KEY (`encomenda_numero`) REFERENCES `encomendas` (`numero`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `encomenda_reservas` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `encomenda_numero` varchar(30) NOT NULL,
  `material_id` varchar(30) DEFAULT NULL,
  `material` varchar(100) NOT NULL,
  `espessura` varchar(20) NOT NULL,
  `quantidade` decimal(10,2) NOT NULL,
  `created_at` datetime DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_enc_res_num` (`encomenda_numero`),
  KEY `idx_enc_res_mat_esp` (`material`,`espessura`),
  CONSTRAINT `encomenda_reservas_ibfk_1` FOREIGN KEY (`encomenda_numero`) REFERENCES `encomendas` (`numero`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `pecas` (
  `id` varchar(30) NOT NULL,
  `encomenda_numero` varchar(30) DEFAULT NULL,
  `ref_interna` varchar(50) DEFAULT NULL,
  `ref_externa` varchar(100) DEFAULT NULL,
  `material` varchar(100) DEFAULT NULL,
  `espessura` varchar(20) DEFAULT NULL,
  `quantidade_pedida` decimal(10,2) DEFAULT NULL,
  `operacoes` varchar(150) DEFAULT NULL,
  `of_codigo` varchar(30) DEFAULT NULL,
  `opp_codigo` varchar(30) DEFAULT NULL,
  `estado` varchar(50) DEFAULT NULL,
  `produzido_ok` decimal(10,2) DEFAULT NULL,
  `produzido_nok` decimal(10,2) DEFAULT NULL,
  `inicio_producao` datetime DEFAULT NULL,
  `fim_producao` datetime DEFAULT NULL,
  `tempo_producao_min` decimal(10,2) DEFAULT NULL,
  `lote_baixa` varchar(100) DEFAULT NULL,
  `observacoes` text,
  `desenho_path` varchar(512) DEFAULT NULL,
  `operacoes_fluxo_json` longtext,
  `qtd_expedida` decimal(10,2) DEFAULT NULL,
  `hist_json` longtext,
  PRIMARY KEY (`id`),
  KEY `idx_pecas_encomenda_numero` (`encomenda_numero`),
  KEY `idx_pecas_estado` (`estado`),
  KEY `idx_pecas_ref_interna` (`ref_interna`),
  CONSTRAINT `pecas_ibfk_1` FOREIGN KEY (`encomenda_numero`) REFERENCES `encomendas` (`numero`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `orcamento_linhas` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `orcamento_numero` varchar(30) DEFAULT NULL,
  `ref_interna` varchar(50) DEFAULT NULL,
  `ref_externa` varchar(100) DEFAULT NULL,
  `material` varchar(100) DEFAULT NULL,
  `espessura` varchar(20) DEFAULT NULL,
  `operacao` varchar(150) DEFAULT NULL,
  `qtd` decimal(10,2) DEFAULT NULL,
  `preco_unit` decimal(10,2) DEFAULT NULL,
  `total` decimal(12,2) DEFAULT NULL,
  `descricao` text,
  `of_codigo` varchar(30) DEFAULT NULL,
  `desenho_path` varchar(512) DEFAULT NULL,
  `tempo_peca_min` decimal(10,2) DEFAULT NULL,
  `tipo_item` varchar(30) DEFAULT NULL,
  `produto_codigo` varchar(20) DEFAULT NULL,
  `produto_unid` varchar(20) DEFAULT NULL,
  `conjunto_codigo` varchar(40) DEFAULT NULL,
  `conjunto_nome` varchar(150) DEFAULT NULL,
  `grupo_uuid` varchar(60) DEFAULT NULL,
  `qtd_base` decimal(10,2) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_orcamento_linhas_orcamento_numero` (`orcamento_numero`),
  CONSTRAINT `orcamento_linhas_ibfk_1` FOREIGN KEY (`orcamento_numero`) REFERENCES `orcamentos` (`numero`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `notas_encomenda_documentos` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `ne_numero` varchar(30) NOT NULL,
  `data_registo` datetime DEFAULT NULL,
  `tipo` varchar(40) DEFAULT NULL,
  `titulo` varchar(150) DEFAULT NULL,
  `caminho` varchar(512) DEFAULT NULL,
  `guia` varchar(60) DEFAULT NULL,
  `fatura` varchar(60) DEFAULT NULL,
  `data_documento` date DEFAULT NULL,
  `obs` text,
  `data_entrega` date DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_ne_docs_num` (`ne_numero`),
  CONSTRAINT `fk_ne_documentos_ne` FOREIGN KEY (`ne_numero`) REFERENCES `notas_encomenda` (`numero`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `notas_encomenda_entregas` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `ne_numero` varchar(30) NOT NULL,
  `data_registo` datetime DEFAULT NULL,
  `data_entrega` date DEFAULT NULL,
  `data_documento` date DEFAULT NULL,
  `guia` varchar(60) DEFAULT NULL,
  `fatura` varchar(60) DEFAULT NULL,
  `obs` text,
  PRIMARY KEY (`id`),
  KEY `idx_ne_entregas_num` (`ne_numero`),
  CONSTRAINT `fk_ne_entregas_ne` FOREIGN KEY (`ne_numero`) REFERENCES `notas_encomenda` (`numero`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `notas_encomenda_linha_entregas` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `ne_numero` varchar(30) NOT NULL,
  `linha_ordem` int(11) NOT NULL,
  `data_registo` datetime DEFAULT NULL,
  `data_entrega` date DEFAULT NULL,
  `data_documento` date DEFAULT NULL,
  `guia` varchar(60) DEFAULT NULL,
  `fatura` varchar(60) DEFAULT NULL,
  `obs` text,
  `qtd` decimal(10,2) DEFAULT NULL,
  `lote_fornecedor` varchar(100) DEFAULT NULL,
  `localizacao` varchar(100) DEFAULT NULL,
  `entrega_total` tinyint(1) DEFAULT NULL,
  `stock_ref` varchar(30) DEFAULT NULL,
  `logistic_status` varchar(30) DEFAULT NULL,
  `inspection_status` varchar(40) DEFAULT NULL,
  `inspection_defect` varchar(255) DEFAULT NULL,
  `inspection_decision` varchar(255) DEFAULT NULL,
  `quality_status` varchar(40) DEFAULT NULL,
  `quality_nc_id` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_ne_linha_entregas_num_ord` (`ne_numero`,`linha_ordem`),
  CONSTRAINT `fk_ne_linha_entregas_ne` FOREIGN KEY (`ne_numero`) REFERENCES `notas_encomenda` (`numero`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

CREATE TABLE IF NOT EXISTS `notas_encomenda_linhas` (
  `id` int(11) NOT NULL AUTO_INCREMENT,
  `ne_numero` varchar(30) DEFAULT NULL,
  `ref_material` varchar(20) DEFAULT NULL,
  `qtd` decimal(10,2) DEFAULT NULL,
  `preco` decimal(10,2) DEFAULT NULL,
  `total` decimal(12,2) DEFAULT NULL,
  `entregue` tinyint(1) DEFAULT NULL,
  `lote_fornecedor` varchar(100) DEFAULT NULL,
  `linha_ordem` int(11) DEFAULT NULL,
  `descricao` varchar(255) DEFAULT NULL,
  `fornecedor_linha` varchar(150) DEFAULT NULL,
  `origem` varchar(50) DEFAULT NULL,
  `unid` varchar(20) DEFAULT NULL,
  `qtd_entregue` decimal(10,2) DEFAULT NULL,
  `material` varchar(100) DEFAULT NULL,
  `espessura` varchar(20) DEFAULT NULL,
  `comprimento` decimal(10,2) DEFAULT NULL,
  `largura` decimal(10,2) DEFAULT NULL,
  `metros` decimal(10,2) DEFAULT NULL,
  `localizacao` varchar(100) DEFAULT NULL,
  `peso_unid` decimal(10,3) DEFAULT NULL,
  `p_compra` decimal(10,4) DEFAULT NULL,
  `formato` varchar(50) DEFAULT NULL,
  `stock_in` tinyint(1) DEFAULT NULL,
  `guia_entrega` varchar(60) DEFAULT NULL,
  `fatura_entrega` varchar(60) DEFAULT NULL,
  `data_doc_entrega` date DEFAULT NULL,
  `data_entrega_real` date DEFAULT NULL,
  `obs_entrega` text,
  `desconto` decimal(6,2) DEFAULT NULL,
  `iva` decimal(6,2) DEFAULT NULL,
  `logistic_status` varchar(30) DEFAULT NULL,
  `inspection_status` varchar(40) DEFAULT NULL,
  `inspection_defect` varchar(255) DEFAULT NULL,
  `inspection_decision` varchar(255) DEFAULT NULL,
  `quality_status` varchar(40) DEFAULT NULL,
  `quality_nc_id` varchar(30) DEFAULT NULL,
  PRIMARY KEY (`id`),
  KEY `idx_notas_encomenda_linhas_ne_numero` (`ne_numero`),
  KEY `idx_ne_linhas_num_ord` (`ne_numero`,`linha_ordem`),
  CONSTRAINT `notas_encomenda_linhas_ibfk_1` FOREIGN KEY (`ne_numero`) REFERENCES `notas_encomenda` (`numero`) ON DELETE CASCADE
) ENGINE=InnoDB DEFAULT CHARSET=utf8;

SET FOREIGN_KEY_CHECKS=1;

-- Fim do schema atual consolidado.
