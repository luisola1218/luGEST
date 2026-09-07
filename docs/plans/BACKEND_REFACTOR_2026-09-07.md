# Reorganizacao do backend — 2026-09-07

## Ponto de recuperacao

O estado anterior foi guardado no commit `87c56b9` e enviado para `luGEST/main`
antes de qualquer alteracao desta fase. O trabalho seguinte vive na branch
`codex/backend-organization`.

## Resultado

- `main_bridge.py`: de 19 741 para 166 linhas; contem composicao e inicializacao.
- 485 metodos extraidos para 32 areas com nomes funcionais. Os nove adaptadores
  anteriores continuam com a mesma precedencia perante estes metodos.
- A API dos 487 metodos originais e verificada por contrato, incluindo properties
  e staticmethods. O construtor recebe agora um argumento opcional `runtime`.
- `LegacyRuntime` declara os onze componentes historicos recebidos no arranque;
  um teste constroi o backend sem importar `main.py` nem ligar ao MySQL.
- Custos de operacoes vivem em `lugest_core/operation_costing.py`, com entradas
  e normalizadores explicitos, sem UI, ficheiros ou MySQL.
- Comparacao e combinacao de snapshots vivem em `lugest_core/snapshots.py`.
- A configuracao usa `lugest_infra/config/repository.py`, com interfaces de
  armazenamento e fabrica de ligacoes explicitas; o adaptador conserva a API.
- [Guia de manutencao](../architecture/BACKEND_GUIDE.md), indice de metodos e
  `scripts/backend_map.py` permitem encontrar a implementacao efetiva de uma
  chamada e as suas dependencias sem arrancar a aplicacao.

## Correcoes associadas

- Linhas historicas sem identificador deixam de ser duplicadas pela combinacao
  de snapshots. Remocoes locais sao respeitadas e linhas locais inalteradas
  deixam de ressuscitar registos removidos remotamente.
- O cache de configuracao devolve copias profundas: editar um dicionario numa
  pagina nao altera silenciosamente a configuracao em cache noutras paginas.
- O repositorio fecha ligacoes em sucesso e erro; tenta rollback de falhas SQL
  e devolve diagnostico de falhas parciais. Leitura nao cria tabelas.

Nao foi alterada a politica de conflito entre alteracoes na mesma linha com ID:
continua a prevalecer a linha local inteira. Nao foram adicionadas garantias
de transacao entre postos ou redesenhado o sistema de gravacao assincrona.

## Evidencias

- Porta segura: 265 ficheiros Python compilados, dependencias consistentes,
  19 verificacoes aprovadas.
- Comparacao AST confirmou a preservacao das implementacoes extraidas. As
  alteracoes posteriores intencionais restringiram-se ao construtor, aos cinco
  adaptadores de snapshots, aos dois de configuracao e aos tres de custos.
- Comparacao diferencial do motor de custos com o codigo do checkpoint:
  150 de 150 casos identicos, variando modo, quantidades, tarifas e confirmacoes.
- Testes novos cobrem contratos, referencias globais dos callbacks, construcao
  injetada, combinacao concorrente, independencia dos dados, matriz de falhas
  da configuracao e calculos de custos.
- Catalogo de produtos: verificacao adicional aprovada com 24 casos.
- Auditoria de seguranca: zero ocorrencias altas ou medias; dois avisos baixos
  existentes na verificacao (`lugest.env` local e cache Python regeneravel).
- `git diff --check` sem erros.

## Limites

A organizacao facilita encontrar e alterar funcionalidades. Os adaptadores
ainda partilham estado e muitos continuam dependentes do runtime historico;
nao se deve apresenta-los como modulos totalmente independentes.

Nao foi gerado instalador nem executado teste com escrita na base do cliente.
Os fluxos integrados de producao, stock e faturacao e o pacote Windows precisam
de validacao em staging antes de distribuir esta branch.
