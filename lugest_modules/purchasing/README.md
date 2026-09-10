# Compras

| Comportamento | Implementacao |
| --- | --- |
| Rascunho, aprovacao, envio, remocao e conversao por fornecedor | `application/note_lifecycle.py` |
| Guardar nota e catalogos de precos numa operacao | `application/note_commands.py` |
| Normalizar linhas e inferir materiais | `application/line_normalization.py` |
| Precos e sincronizacao das linhas | `application/pricing.py`, `application/material_lines.py` |
| Estados derivados das entregas | `application/note_status.py` |
| Gravacao preparada das notas e sequencia local | `infrastructure/legacy_note_repository.py` |
| Datas e referencias ausentes da tabela historica | `infrastructure/note_metadata.py` |

O ponto de composicao e `lugest_qt/services/purchasing_composition.py`.
Os metodos publicos antigos delegam nestes casos de uso. A marcacao de envio
apenas regista o estado; nao envia mensagens externas.

A conversao valida todos os fornecedores antes de reservar numeros e prepara
as notas filhas e a nota de origem antes de uma gravacao bloqueante unica.
Uma cotacao ja convertida e rejeitada, evitando duplicacao de encomendas.
Falhas imediatas recuperam as notas e a sequencia local. A reserva de numeros
SQL continua no adaptador historico e pode deixar intervalos.

Os metadados sao gravados no payload `runtime_state` existente, na mesma
transacao que as tabelas relacionais. Nao e necessaria uma alteracao de schema.
O carregamento aplica apenas os campos permitidos a notas que ainda existem.
Datas perdidas antes desta correcao nao sao reconstruidas nem inventadas.

A gravacao completa usa `LegacyPurchaseRepository` para notas, produtos,
materiais e conjuntos. Os calculos decorrem em copias; os conjuntos sao
recalculados uma vez depois de atualizar os precos. Apenas o resultado final e
publicado. Os adaptadores antigos reutilizam as mesmas regras numericas.

## Verificacao

- `scripts/verify_purchase_commands.py`: preparacao dos quatro catalogos,
  falhas de calculo/escrita, uma gravacao e conflitos locais.
- `scripts/verify_purchase_lines.py`: materiais, produtos, inferencia e valores
  nao finitos.
- `scripts/verify_purchase_lifecycle.py`: validacoes, estados, duplicacao,
  falhas de gravacao, conflitos locais e metadados.
- `scripts/verify_database_rollback.py verify_purchase_lifecycle_flow`:
  rascunho, aprovacao, envio e remocao, com releitura apos cada operacao.
- `scripts/verify_database_rollback.py verify_purchase_flow`: conversao,
  documentos e rececao de produtos e materia-prima.

## Fronteira ainda historica

Sugestoes e historico de fornecedores, rececao e documentos ainda dependem
dos adaptadores antigos. A geometria e algumas regras de valor continuam
ligadas por capacidades explicitas na composicao. Este modulo
nao representa a conclusao da migracao de compras ou do ERP.
