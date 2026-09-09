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

Rececao, movimentos e libertacao de stock ainda pertencem ao adaptador
historico em `lugest_qt/services/bridge_mixins/quality.py`.
O armazenamento de ficheiros ainda usa o adaptador partilhado: a recuperacao
do registo nao remove automaticamente um ficheiro copiado antes de uma falha
de gravacao. A remocao de metadados tambem nao apaga o ficheiro, que pode ter
outras referencias. Nao assumir uma transacao unica entre SQL e ficheiros.
