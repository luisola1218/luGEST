# Estrutura do projeto e regras de evolução

Estado atual: a aplicacao ja tem separacao parcial em `lugest_core`,
`lugest_infra`, `lugest_desktop/legacy` e `lugest_qt`, mas ainda existem dois
ficheiros muito grandes que concentram demasiado comportamento:
`lugest_qt/ui/pages/runtime_pages.py` e `lugest_qt/services/main_bridge.py`.

Atualizacao incremental: a janela principal deve importar paginas pelos modulos
proprios em `lugest_qt/ui/pages/*_page.py`, mesmo quando a implementacao ainda
e uma fachada para `runtime_pages.py`. O backend Qt deve ser importado por
`lugest_qt.services.legacy_backend`, deixando `main_bridge.py` como detalhe de
compatibilidade legacy.

## Alvo modular

```text
lugest/
  app/
    qt_app.py
    main_window.py
  modules/
    clientes/
    orcamentos/
    encomendas/
    compras/
    stock/
    operador/
    planeamento/
    faturacao/
    transportes/
    laser/
  core/
    cad/
    laser/
    pricing/
    compliance/
  infra/
    db/
    storage/
    pdf/
    config/
  legacy/
    desktop_tk/
    compat/
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

## Organização de runtime

- Diagnóstico e crash logging vivem em `lugest_infra/diagnostics`; a UI apenas
  emite eventos.
- Tokens comerciais vivem em `lugest_infra/licensing`, separados da base do
  cliente e de `lugest_qt_config.json`.
- O build comercial não inclui Qt WebEngine/QML: o módulo de Transportes usa o
  fallback já existente para abrir rotas no navegador do posto. Isto reduz o
  pacote sem retirar funções operacionais.
- Regras e formato de licença vivem em `lugest_core/licensing`, sem dependência
  da interface ou do MySQL.
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

- Dividir `runtime_pages.py` por paginas reais.
- Dividir `main_bridge.py` por servicos de dominio.
- Trocar `module_context.py` por dependencias explicitas.
- Renomear gradualmente `*_rooting.py` para nomes claros como `*_ui.py` ou
  `*_routing.py`, mantendo shims temporarios.
- Centralizar runtime JSON em MySQL `app_config`, deixando ficheiros JSON so
  para modo local, demo ou fallback controlado.
