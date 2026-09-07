# Reorganizacao do Lugest — 2026-09-07

## Alteracoes entregues

- O antigo `runtime_pages.py`, com cerca de 25 800 linhas, passa a fachada de
  compatibilidade. As implementacoes estao separadas em orcamentos, encomendas,
  operador, planeamento, expedicao, transportes, OPP e extensao de compras.
- As classes base e as extensoes historicas mantem identidade e heranca.
  A extracao verificou equivalencia AST das 36 definicoes movidas, preservando
  corpos, assinaturas e decoradores.
- `page_registry.py` centraliza as 20 paginas e respetivas dependencias.
  Importar a janela e criar as factories ja nao carrega todas as paginas.
  Cada modulo e importado na primeira abertura. Nao foi medido um ganho
  percentual de tempo nem de memoria numa instalacao de cliente.
- Os 15 metodos de versao/atualizacao do backend vivem agora num adaptador
  dedicado. Downloads falhados limpam staging e downloads HTTP diretos sao
  rejeitados antes de criar ficheiros temporarios.
- O cache de consultas deixa de partilhar listas/dicionarios internos com a
  UI e com o produtor. Usa relogio monotonico, expiracao no limite do TTL,
  descarte LRU e no maximo 128 entradas. O runtime pode ser injetado em testes.
- A verificacao do projeto fixa a diretoria de trabalho na raiz e restaura-a
  no fim. A compilacao deixa de percorrer ambientes, builds e caches excluidos.
- As fronteiras de arquitetura passam a rejeitar tambem imports diretos de
  Qt/Tk no dominio e na infraestrutura.

## Validacao

`powershell -NoProfile -ExecutionPolicy Bypass -File scripts/verify_project.ps1 -SafeOnly`

- 224 ficheiros Python compilados; dependencias instaladas consistentes.
- 15 verificacoes aprovadas: seguranca, arquitetura, lock, armazenamento,
  atualizacoes, diagnostico, fundacao de licencas, schema, migracoes, calculo
  laser, nesting, controlos de orcamentos, parceiros, cache e modulos de paginas.
- Os testes novos nao abrem MySQL. Verificam carregamento tardio, despachos
  das factories, imports historicos e referencias globais dos callbacks.
- Auditoria de seguranca: nenhuma ocorrencia alta ou media. Os avisos baixos
  da execucao foram o `lugest.env` local e cache Python regenerado pelos testes.
- `git diff --check` sem erros.

As alteracoes locais anteriores foram preservadas. Esta passagem nao fez
instalacao, publicacao, commit nem implementou bloqueio comercial.

## Limites e continuacao

A modularizacao melhora a manutencao, mas nao prova ausencia de bugs em todos
os fluxos. `main_bridge.py` ainda tem cerca de 19 700 linhas e `quotes_page.py`
cerca de 10 900; dialogos de montagem e edicao de linhas sao as proximas
extracoes por responsabilidade. `module_context` continua como divida legacy.

Antes de venda geral, validar numa base de staging os fluxos que escrevem
dados e testar o executavel empacotado, instalacao, atualizacao e recuperacao
num Windows limpo. A compatibilidade do empacotamento foi revista no spec,
mas nao foi produzido um novo executavel nesta passagem.

A integracao de licencas deve reutilizar `lugest_core/licensing` e
`lugest_infra/licensing`, com verificacao nos servicos e composicao no arranque.
O registo de paginas nao substitui autorizacao. Os planos, postos, modulos e
regras offline pertencem ao passo comercial seguinte.
