# Fase seguinte: clientes e testes integrados

## Resultado implementado

- Casos de uso de clientes separados em `lugest_core/clients.py`, com contrato
  de repositorio explicito e sem imports de interface, infraestrutura ou `main`.
- Repositorio de compatibilidade em `lugest_infra/legacy/clients_repository.py`,
  com callbacks declarados para leitura, escrita, codigo seguinte e referencias.
- O adaptador Qt conserva as cinco assinaturas existentes e apenas compoe o
  servico. Nao foi alterada a pagina de clientes.
- Upserts sincronos atualizam o cache e o snapshot apenas depois de sucesso.
  Falhas imediatas da gravacao do dataset restauram a colecao de clientes.
  Isto nao introduz transacoes distribuidas nem garante gravacao apos aceitar
  uma tarefa na fila assincrona.

## Testes

- O checkpoint passou novamente as 19 verificacoes seguras existentes.
- Novo teste integrado: pagina Qt -> backend real -> servico -> repositorio,
  com persistencia simulada e sem importar `main.py`.
- Cobertura: criar/editar/remover, pesquisa sem acentos, seletores de encomenda,
  validacao do nome, referencias por encomenda/orcamento, falha de escrita,
  falha do upsert direto, independencia do resultado e despacho assincrono.
- A verificacao dos 487 contratos e a compilacao continuam ativas.
- Resultado final da porta segura: 271 ficheiros compilados e 21 verificacoes
  aprovadas, incluindo os novos testes de clientes e substituicao de snapshots.

## Tentativa de MySQL integrado

A configuracao local corresponde ao servidor e base mostrados pelo utilizador.
A ligacao ao servidor foi validada com uma consulta simples. A tentativa de
criar uma base nova, com nome aleatorio `lugest_stress_<uuid>`, foi recusada
com erro MySQL 1044 (permissao insuficiente).

Nessa tentativa nenhuma base foi criada e nenhum fluxo com escrita arrancou.
O utilizador autorizou depois explicitamente os testes na base atual. Foram
executados pelo runner de rollback descrito abaixo.

`scripts/verify_isolated_database.py` prepara uma base descartavel e dados
sinteticos, restringe as ligacoes a essa base, isola os caminhos de runtime e
remove apenas a base que criou. Nao pertence a `-SafeOnly`. Requer uma conta
de testes com permissao para criar/remover a sua base no servidor de staging.
As credenciais da aplicacao em producao nao devem ser alargadas para esse fim.

## Testes autorizados na base atual

`scripts/verify_database_rollback.py` executou um fluxo por processo na base
configurada. As 55 tabelas foram confirmadas como InnoDB. As ligacoes foram
encaminhadas para uma unica transacao, as confirmacoes de gravacao suprimidas,
as alteracoes de estrutura bloqueadas e os dados revertidos no fim. Caminhos
de runtime e ficheiros partilhados foram redirecionados para pastas temporarias.

| Fluxo | Resultado | Duracao | Conteudo das 55 tabelas apos rollback |
| --- | --- | --- | --- |
| Compras | Aprovado | 18,1 s | Igual |
| Conjuntos e montagem | Aprovado | 38,0 s | Igual |
| Ordem de fabrico | Aprovado | 18,6 s | Igual |
| Planeamento | Aprovado | 69,3 s | Igual |
| Faturacao | Aprovado | 18,3 s | Igual |

A igualdade foi verificada por contagem e SHA-256 do conteudo de cada tabela,
antes e depois de cada execucao. Os hashes nao incluem metadados como o proximo
AUTO_INCREMENT: no MySQL podem ficar intervalos na numeracao mesmo apos rollback.
Este modo serializa as ligacoes e desativa a fila de gravacao; nao valida
concorrencia real entre postos nem a durabilidade da fila assincrona.

## Bugs encontrados e corrigidos

- Geracao de UUID em conjuntos, faturacao, servicos e dialogo laser dependia de
  `main.uuid`, que nao existe. Esses modulos importam agora `uuid` diretamente;
  as dependencias indiretas equivalentes de `datetime` tambem foram removidas.
- Reparar referencias de cliente durante a gravacao de um orcamento pode
  substituir o snapshot do backend. O orcamento era publicado na copia antiga,
  causando "Orcamento nao encontrado". A gravacao usa agora o snapshot atual.
- Importar uma peca de um conjunto pode substituir o snapshot da encomenda.
  Itens de montagem e fichas do conjunto eram depois escritos numa referencia
  antiga. A importacao volta a obter a encomenda atual entre essas operacoes.
- O teste de ordem de fabrico tambem foi corrigido para preparar os produtos no
  snapshot atual, apos guardar o conjunto.

`verify_quote_snapshot_save.py` reproduz as substituicoes de snapshot sem MySQL
e cobre criacao/edicao de orcamentos e importacao de pecas, montagem e fichas.
O teste de contratos impede reintroduzir `desktop_main.uuid/datetime`.

A validacao do pacote Windows continua pendente. O sistema comercial de licencas
nao foi ativado.
