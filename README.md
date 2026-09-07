# luGEST

ERP industrial desktop para orcamentos, encomendas, producao, planeamento,
operador, expedicao, faturacao, compras e stock.

## Stack

- Python + PySide6 para a aplicacao desktop
- MySQL para persistencia central
- ReportLab/Pillow para PDF e imagem

## Arranque rapido

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements-qt.txt
.\.venv\Scripts\python.exe -m lugest_qt.app
```

Para reproduzir exatamente o ambiente já validado para build/release, usar
`requirements-qt.lock.txt`. O ficheiro `requirements-qt.txt` mantém intervalos
compatíveis para desenvolvimento; o lock só é atualizado depois da porta de
qualidade completa.

Para criar o ambiente local:

```text
config/examples/lugest.env.example -> lugest.env
```

Campos minimos:

```env
LUGEST_DB_HOST=127.0.0.1
LUGEST_DB_PORT=3306
LUGEST_DB_USER=lugest_user
LUGEST_DB_PASS=trocar-password
LUGEST_DB_NAME=lugest
```

## Estrutura principal

```text
lugest_qt/app.py                 entrada desktop atual (python -m lugest_qt.app)
main.py                         runtime historico usado pelo adaptador legacy
lugest_qt/                      UI Qt, paginas e bridge da aplicacao
impulse_mobile_app/             app Flutter LuGEST Field para servicos no terreno
lugest_core/                    logica de dominio reutilizavel
lugest_infra/                   infraestrutura, storage e PDF
lugest_desktop/legacy/          codigo historico ainda usado pelo Qt
mysql/                          schema, patches e tooling MySQL
scripts/                        verificacoes, seeds, release e manutencao
docs/                           manuais, instalacao, planos e arquitetura
config/examples/                exemplos de configuracao ambiente
generated/                      artefactos gerados, fora do Git
backups/                        copias locais, fora do Git
```

## Documentacao

- [Indice da documentacao](docs/README.md)
- [Manual de operacao](docs/manual/MANUAL_OPERACAO_LUISGEST_PROFISSIONAL.md)
- [Guia de arranque Qt local](docs/install/GUIA_ARRANQUE_QT_LOCAL.md)
- [Plano de faturacao](docs/plans/FATURACAO_PLAN.md)
- [Arquitetura mobile de servicos](docs/plans/MOBILE_SERVICES_ARCHITECTURE.md)
- [Auditoria da base de dados](docs/plans/DATABASE_AUDIT_2026-08-18.md)
- [Auditoria técnica do sistema](docs/plans/SYSTEM_AUDIT_2026-08-18.md)
- [Plano de conjuntos e montagem](docs/plans/CONJUNTOS_MONTAGEM_PLAN.md)
- [Estrutura recomendada](docs/architecture/PROJECT_STRUCTURE.md)
- [Guia de manutencao do backend](docs/architecture/BACKEND_GUIDE.md)
- [Base de dados MySQL](mysql/README.md)

## Verificacoes

```powershell
.\scripts\verify_project.ps1 -SafeOnly
.\.venv\Scripts\python.exe scripts\verify_laser_quote_engine.py
.\.venv\Scripts\python.exe scripts\verify_conjuntos_montagem_flow.py
.\.venv\Scripts\python.exe scripts\verify_purchase_flow.py
powershell -ExecutionPolicy Bypass -File scripts\verify_project.ps1
```

`-SafeOnly` compila o projeto e executa verificações isoladas de dependências,
segurança, licenciamento, laser, nesting e controlos Qt sem escrever na base de
dados. Para uma auditoria explícita da base em modo só de leitura, acrescenta
`-ReadOnlyDatabaseAudit`. Os fluxos completos só devem ser executados numa base
de testes ou com autorização consciente através de `-AllowRemoteDatabase`.

## Notas de manutencao

- A app de producao usa MySQL; JSONs de runtime devem ser tratados como estado
  local ou fallback, nao como fonte de verdade.
- A raiz ainda contem alguns wrappers de compatibilidade para modulos legacy.
  Devem desaparecer quando nao houver imports externos dependentes deles.
- As paginas vivem nos respetivos modulos e sao carregadas quando abertas,
  atraves de `lugest_qt/ui/page_registry.py`. `runtime_pages.py` conserva apenas
  imports de compatibilidade. A janela principal nao importa todas as paginas.
- `main_bridge.py` compoe os adaptadores por area. Para localizar uma alteracao,
  usar `python scripts/backend_map.py nome_do_metodo` e o guia do backend.
  Os adaptadores preservam a integracao legacy; custos de operacoes, combinacao
  de snapshots e persistencia de configuracao ja sao componentes independentes.
- O licenciamento comercial tem uma base assinada e testável em
  `lugest_core/licensing`, mas ainda não bloqueia o produto: os planos, módulos,
  postos e tolerância offline têm de ser decididos antes da integração.
