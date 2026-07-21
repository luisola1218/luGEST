# Estrutura recomendada do projeto

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

## Proximos refactors recomendados

- Dividir `runtime_pages.py` por paginas reais.
- Dividir `main_bridge.py` por servicos de dominio.
- Trocar `module_context.py` por dependencias explicitas.
- Renomear gradualmente `*_rooting.py` para nomes claros como `*_ui.py` ou
  `*_routing.py`, mantendo shims temporarios.
- Centralizar runtime JSON em MySQL `app_config`, deixando ficheiros JSON so
  para modo local, demo ou fallback controlado.
