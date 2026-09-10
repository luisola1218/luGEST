# Qualidade

## Onde alterar

| Comportamento | Implementacao |
| --- | --- |
| Criar, editar, fechar e remover NC; duplicados e consultas | `application/nonconformities.py` |
| Catalogo de documentos e respetiva auditoria | `application/documents.py` |
| Gravacao dos catalogos com auditoria e recuperacao local | `infrastructure/legacy_catalog_repository.py` |

`api.py` e a entrada publica dos casos de uso. As capacidades de armazenamento,
ficheiros, identificadores e utilizador sao ligadas em
`lugest_qt/services/quality_composition.py`.

Os servicos trabalham com copias dos registos e validam antes de publicar.
O repositorio verifica o estado local esperado, grava de forma bloqueante e
recupera o catalogo e a auditoria perante falha imediata. O adaptador antigo
`LegacyNonconformityRepository` reutiliza este mesmo mecanismo.

## Verificacao

- `scripts/verify_quality_nonconformities.py`: ciclo de NC, duplicados,
  quantidades, conflitos locais e recuperacao de falhas.
- `scripts/verify_quality_documents.py`: CRUD, validacao, isolamento de dados
  e recuperacao de falhas de auditoria/gravacao.
- `scripts/verify_database_rollback.py verify_quality_nc_flow`: SQL com
  releitura e rollback.
- `scripts/verify_database_rollback.py verify_quality_document_flow`: inclui
  importacao de ficheiro em armazenamento temporario, releitura e rollback.

## Fronteira atual

Rececao, conciliacao, libertacao e consultas usam casos de uso do modulo.
O mixin quality.py mantem as assinaturas de compatibilidade; relatórios ainda
usam quality_reports.py e a persistencia passa pelo runtime partilhado.
O armazenamento de ficheiros ainda usa o adaptador partilhado: a recuperacao
do registo nao remove automaticamente um ficheiro copiado antes de uma falha
de gravacao. A remocao de metadados tambem nao apaga o ficheiro, que pode ter
outras referencias. Nao assumir uma transacao unica entre SQL e ficheiros.


## Libertacao de material por NC

`application/material_release.py` prepara materiais, NC e notas de compra em
copias. `stock_policy.py` concentra a politica de quarentena, tambem usada
pelas fachadas de compatibilidade. A composicao fornece a sincronizacao de
linhas de compra por uma capacidade explicita.
`infrastructure/legacy_release_repository.py` verifica o estado esperado,
prepara stock_log/audit_log antes de publicar e pede uma gravacao bloqueante.
Em erro imediato recupera os cinco catalogos, incluindo campos antes ausentes.

Teste: `scripts/verify_quality_release.py` cobre falhas, repeticao sem duplicar
stock e alteracoes concorrentes locais. O fluxo SQL `verify_quality_nc_flow`
cobre tambem libertacao e releitura (55 tabelas inalteradas apos rollback).
A cobertura foi alargada a rececoes associadas e reconciliacao apos releitura
em `verify_quality_reception_flow`.


## Rececao e conciliacao

- `application/receptions.py`: avalia um movimento, prepara NC e devolucao;
  repositorio local de NC recolhe alteracoes e auditorias para a mesma gravacao.
- `application/delivery_movements.py`: projecao de movimentos, totais de qualidade
  e libertacao de pendentes. Quantidades explicitas prevalecem sobre estimativas
  antigas de NC; NC sem movimento so permite inferencia numa rececao unica.
- `application/queries.py`: resumo, integridade e selecao de entidades.
- `infrastructure/legacy_reception_repository.py`: verifica catalogos esperados,
  prepara logs, publica uma vez e recupera oito catalogos em erro imediato.
- `infrastructure/reception_metadata.py`: conserva decisoes quantitativas,
  identificadores de movimentos e metadados de devolucao no runtime_state SQL.
  A reaplicacao verifica referencia e quantidade do movimento; nao cria linhas.

Testes: `verify_quality_receptions.py` e fluxo protegido
`verify_database_rollback.py verify_quality_reception_flow`. Incluem parcial,
rejeicao, devolucao, isolamento de notas, valores nao finitos, falha de gravacao,
metadados, libertacao associada, conciliacao e consultas apos releitura.
