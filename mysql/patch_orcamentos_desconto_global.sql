ALTER TABLE `orcamentos`
  ADD COLUMN `desconto_perc` DECIMAL(6,2) NULL;

ALTER TABLE `orcamentos`
  ADD COLUMN `desconto_valor` DECIMAL(12,2) NULL;

ALTER TABLE `orcamentos`
  ADD COLUMN `subtotal_bruto` DECIMAL(12,2) NULL;
