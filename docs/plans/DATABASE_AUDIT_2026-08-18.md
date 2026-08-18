# Auditoria da base de dados — 2026-08-18

## Âmbito

A auditoria foi executada contra a base configurada, exclusivamente em transação
`READ ONLY`, com o script `scripts/audit_database_readonly.py`.

## Resultado

- esquema mínimo validado com sucesso;
- todas as tabelas usam InnoDB;
- zero origens de faturação duplicadas por orçamento, encomenda ou serviço;
- zero identificadores/documentos fiscais duplicados;
- zero pagamentos órfãos de fatura;
- zero ligações órfãs entre serviços diretos e faturação;
- zero totais negativos em serviços, faturas ou pagamentos;
- zero documentos de linhas de serviço com JSON inválido.

Não foi aplicada qualquer correção aos dados.

## Correção efetuada no projeto

O ficheiro oficial `mysql/lugest.sql` não incluía a tabela
`servicos_diretos` nem a coluna `faturacao_registos.servico_numero`, embora o
runtime já as criasse. O schema oficial e os dois validadores de instalação
foram atualizados. Assim, instalações novas deixam de depender do arranque da
aplicação para completar a estrutura.

## Melhorias recomendadas

### Prioridade alta antes da API mobile

1. Criar uma API HTTPS; nunca permitir acesso MySQL direto pelo telemóvel.
2. Adicionar `company_id`, `version` e `sync_updated_at` às entidades móveis.
3. Normalizar linhas, eventos e anexos de serviços, mantendo `linhas_json`
   apenas durante a transição.
4. Adicionar índices únicos após verificação prévia para `documento_id`,
   `legal_invoice_no`, `(serie_id, seq_num)` e `pagamento_id`.

### Migração de charset

Foram encontradas 52 tabelas ainda em `utf8`/`utf8mb3`, apesar de InnoDB estar
correto. A conversão para `utf8mb4` deve ser feita numa janela de manutenção:

1. backup completo e teste de restauro;
2. inventário de índices que possam exceder o limite do servidor;
3. ensaio da conversão numa cópia da base;
4. conversão por grupos de tabelas;
5. validação funcional, fiscal e de integridade;
6. só depois execução em produção.

Não se recomenda uma conversão automática durante o arranque do ERP porque
pode bloquear tabelas e aumentar o risco operacional.
