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
main.py                         entrada desktop atual
lugest_qt/                      UI Qt, paginas e bridge da aplicacao
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
- [Guia de instalacao total](docs/install/GUIA_INSTALACAO_TOTAL.md)
- [Plano de faturacao](docs/plans/FATURACAO_PLAN.md)
- [Plano de conjuntos e montagem](docs/plans/CONJUNTOS_MONTAGEM_PLAN.md)
- [Estrutura recomendada](docs/architecture/PROJECT_STRUCTURE.md)
- [Base de dados MySQL](mysql/README.md)

## Verificacoes

```powershell
.\.venv\Scripts\python.exe scripts\verify_laser_quote_engine.py
.\.venv\Scripts\python.exe scripts\verify_conjuntos_montagem_flow.py
.\.venv\Scripts\python.exe scripts\verify_purchase_flow.py
powershell -ExecutionPolicy Bypass -File scripts\verify_project.ps1
```

## Notas de manutencao

- A app de producao usa MySQL; JSONs de runtime devem ser tratados como estado
  local ou fallback, nao como fonte de verdade.
- A raiz ainda contem alguns wrappers de compatibilidade para modulos legacy.
  Devem desaparecer quando nao houver imports externos dependentes deles.
- `runtime_pages.py` e `main_bridge.py` sao os proximos candidatos a divisao
  por modulo: orcamentos, compras, stock, operador, faturacao e laser.
