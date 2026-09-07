# Estrutura do projeto e regras de evolução

A evolucao atual e um [monolito modular por negocio](MODULAR_MONOLITH.md).
`lugest_modules/clients`, `lugest_modules/quotes` e `lugest_modules/inventory` ja contem componentes com
contratos explicitos. As camadas e adaptadores descritos abaixo ainda suportam
as areas historicas; a migracao integral continua pendente.

Estado em 2026-09-07: a aplicacao tem separacao em `lugest_core`,
`lugest_infra`, `lugest_desktop/legacy` e `lugest_qt`. As implementacoes das
paginas foram retiradas de `runtime_pages.py`; o ficheiro conserva uma fachada
de compatibilidade com resolucao por procura. `main_bridge.py` foi reduzido a
composicao e inicializacao; 485 metodos foram separados em 32 areas. A pagina
de orcamentos continua grande. Consultar o [guia do backend](BACKEND_GUIDE.md)
para localizar implementacoes e conhecer os limites da integracao legacy.

Atualizacao incremental: a janela principal deve importar paginas pelos modulos
proprios em `lugest_qt/ui/pages/*_page.py`, atraves do registo central e lazy em
`lugest_qt/ui/page_registry.py`. O backend Qt deve ser importado por
`lugest_qt.services.legacy_backend`, deixando `main_bridge.py` como detalhe de
compatibilidade legacy.

## Estrutura implementada

```text
lugest_qt/
  app.py                         composicao e arranque
  ui/main_window.py              navegacao e ciclo de vida
  ui/page_registry.py            registo de paginas e dependencias
  ui/pages/*_page.py             pontos de entrada por area
  ui/pages/*_workspace.py        paginas base e extensoes da mesma area
  ui/pages/runtime_common.py     primitivas de widgets e tabelas
  ui/pages/runtime_support.py    dialogos e helpers partilhados
  ui/pages/runtime_pages.py      compatibilidade, sem implementacoes
  services/legacy_backend.py     entrada publica do adaptador
  services/bridge_mixins/        adaptadores por responsabilidade
  services/runtime_service.py    consultas com cache limitado
lugest_core/                     dominio, sem UI nem infraestrutura
lugest_infra/                    caminhos, configuracao, storage, PDF
lugest_desktop/legacy/           compatibilidade historica
mysql/                          schema e migracoes
scripts/                        verificacao, manutencao e release
```

## Regras de limpeza

- `generated/`, `backups/`, `build/`, `dist/` e caches nao entram no Git.
- Documentacao vive em `docs/`.
- Exemplos de ambiente vivem em `config/examples/`.
- Runtime local fica fora do Git: sequencias, trial, estado de UI e similares.
- O codigo novo deve entrar primeiro em `lugest_core`, `lugest_infra` ou em
  mixins pequenos de `lugest_qt/services/bridge_mixins`.
- Novas chamadas externas ao backend Qt devem importar `LegacyBackend` de
  `lugest_qt.services.legacy_backend`, nao de `main_bridge.py`.

## Direção das dependências

```text
lugest_qt  --------------------->  lugest_core
    |                                  ^
    +---------->  lugest_infra  -------+
    |
    +---------->  lugest_desktop/legacy (compatibilidade temporária)
```

- `lugest_core` não importa Qt, desktop legacy ou infraestrutura.
- `lugest_infra` não importa Qt nem desktop legacy.
- `lugest_qt` coordena interface e adaptadores, sem colocar regras de negócio em
  widgets novos.
- `lugest_desktop/legacy` é uma fronteira de compatibilidade, não o destino de
  novas funcionalidades.
- `scripts/verify_architecture_boundaries.py` impede regressões nestas regras.
- O registo das 20 paginas explicita os argumentos de construcao; criar as
  factories nao importa as paginas. O PyInstaller inclui os modulos atraves
  de `collect_submodules('lugest_qt')` no spec existente.
- Imports antigos conservam as classes originais, incluindo a diferenca entre
  a pagina base e a extensao historica. Nao trocar uma pela outra ao migrar.
- `RuntimeService` aceita um runtime injetado para testes sem MySQL. O cache
  usa relogio monotonico, copias profundas, TTL e limite LRU de 128 entradas.
  O limite e por entradas, nao por bytes; payloads grandes devem ser medidos.

## Organização de runtime

- Diagnóstico e crash logging vivem em `lugest_infra/diagnostics`; a UI apenas
  emite eventos.
- `lugest_infra/app_paths.py` resolve de forma centralizada os caminhos de
  runtime. Configuracao, logs e cache ficam em `%LOCALAPPDATA%\luGEST-data`,
  fora da pasta do executavel; `LUGEST_USER_DATA_DIR` permite um override
  controlado em testes e instalacoes especiais.
- Escritas JSON locais usam `lugest_infra/config/AtomicJsonStore`: substituicao
  atomica, tamanho limitado e copia de recuperacao. Uma falha total ao guardar
  configuracao deixa de ser silenciosa.
- Tokens comerciais vivem em `lugest_infra/licensing`, separados da base do
  cliente e de `lugest_qt_config.json`.
- O build comercial não inclui Qt WebEngine/QML: o módulo de Transportes usa o
  fallback já existente para abrir rotas no navegador do posto. Isto reduz o
  pacote sem retirar funções operacionais.
- Regras e formato de licença vivem em `lugest_core/licensing`, sem dependência
  da interface ou do MySQL.
- Regras do manifesto de atualizacao vivem em `lugest_core/updates`; a UI e os
  scripts apenas coordenam download e instalacao. ZIP e bootstrap exigem hashes
  SHA-256 completos e referencias remotas HTTPS.
- Configurações locais, segredos e estado não devem tornar-se módulos de domínio.

## Qualidade antes de alterar estrutura

Executar primeiro:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\verify_project.ps1 -SafeOnly
```

Esta porta de qualidade não escreve na base configurada. Os testes funcionais
completos só podem correr numa base de testes ou quando a base remota foi
explicitamente autorizada.

## Proximos refactors recomendados

- Dividir a pagina de orcamentos por editor, catalogo e apresentacao de linhas.
- Extrair as restantes regras dos adaptadores para servicos de dominio, seguindo
  os exemplos de `operation_costing.py`, `snapshots.py` e do repositorio de
  configuracao; os adaptadores ainda partilham estado legacy.
- Trocar `module_context.py` por dependencias explicitas.
- Renomear gradualmente `*_rooting.py` para nomes claros como `*_ui.py` ou
  `*_routing.py`, mantendo shims temporarios.
- Extrair gradualmente os restantes JSON de runtime para stores atomicos
  dedicados, mantendo MySQL apenas quando a configuracao tiver de ser comum a
  varios postos.

## Preparacao para licencas

Manter a politica em `lugest_core/licensing` e armazenamento/verificacao local
em `lugest_infra/licensing`. A futura integracao deve compor estes componentes
no arranque e aplicar autorizacao nos servicos, alem da visibilidade de menus.
O registo de paginas descreve navegacao; nao constitui uma barreira de licenca.
Esta reorganizacao nao ativa bloqueios nem altera planos comerciais.

## Regressao da modularizacao

`verify_page_modules.py` valida carregamento tardio, classes das 20 paginas,
dependencias dos callbacks e imports de compatibilidade sem ligar ao MySQL.
`verify_runtime_cache.py` cobre isolamento de dados aninhados, expiracao,
invalidacao, atualizacao forcada e limite LRU. Ambos integram `-SafeOnly`.
