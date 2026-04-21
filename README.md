# luGEST

ERP industrial desktop com foco em orçamentos, encomendas, produção, planeamento por operação, expedição, faturação e apoio mobile.

O projeto está organizado em torno da app desktop principal em Python/Qt, com persistência em MySQL, geração de PDFs, armazenamento partilhado e dois módulos opcionais:
- `impulse_mobile_api`: API de leitura para o Pulse / acompanhamento mobile
- `impulse_mobile_app`: app Flutter que consome a API

## Stack

- Python
- PySide6
- MySQL
- ReportLab
- Pillow
- Flutter, no módulo mobile

## Arranque rápido local

Na raiz do projeto:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements-qt.txt
```

Depois:

```powershell
.\.venv\Scripts\Activate.ps1
python main.py
```

Também podes usar:

```powershell
py main.py
```

No Windows, o arranque recomendado continua a ser:

```powershell
.\arrancar_lugest_qt.bat
```

## Configuração mínima

Copiar:

```text
lugest.env.example -> lugest.env
```

Campos mínimos:

```env
LUGEST_DB_HOST=127.0.0.1
LUGEST_DB_PORT=3306
LUGEST_DB_USER=lugest_user
LUGEST_DB_PASS=trocar-password
LUGEST_DB_NAME=lugest
```

Em ambientes com vários postos, é recomendado definir também:

```env
LUGEST_SHARED_STORAGE_ROOT=\\\\SERVIDOR\\LuGEST\\Ficheiros
```

Notas:
- o desktop arranca em Qt
- a app exige ligação ativa a MySQL
- usa um utilizador MySQL dedicado em vez de `root`

## Estrutura principal

```text
main.py                         app desktop principal
plan_actions.py                 lógica de planeamento e PDF do plano
encomendas_actions.py           gestão de encomendas
ne_expedicao_actions.py         expedição e guias
lugest_qt/                      frontend Qt e bridge com o backend
mysql/                          schema, patches e scripts MySQL
scripts/                        seeds, verificações e utilitários
generated/                      ficheiros gerados
impulse_mobile_api/             API mobile
impulse_mobile_app/             app Flutter
```

## Módulos principais

- `Orçamentos`: criação de propostas e conversão em encomenda
- `Encomendas`: detalhe por materiais, espessuras, peças e operações
- `Planeamento`: quadro semanal por operação/recurso, auto-planeamento e PDFs
- `Operador`: execução de operações, tempos e estados de produção
- `Expedição`: preparação de guias e controlo de saída
- `Faturação`: documentos, PDFs e exportação fiscal
- `Impulse`: leitura operacional via API e app mobile

## Base de dados

A aplicação foi pensada para trabalhar com MySQL.

Documentação útil:
- [mysql/README.md](c:/Users/engenharia/VSCodeProjects/teste/mysql/README.md)
- [GUIA_INSTALACAO_TOTAL.md](c:/Users/engenharia/VSCodeProjects/teste/GUIA_INSTALACAO_TOTAL.md)

Na pasta `mysql/` tens:
- schema principal
- patches incrementais
- scripts de backup, restore e validação

## Verificações e testes

Existem vários scripts de verificação na pasta `scripts/`.

Exemplos:

```powershell
python scripts/verify_planning_flow.py
python scripts/verify_core_flows.py
python scripts/verify_internal_full_cycle.py
```

São úteis para validar regressões em:
- planeamento
- expedição
- faturação
- integridade de dados

## Documentação complementar

- [GUIA_ARRANQUE_QT_LOCAL.md](c:/Users/engenharia/VSCodeProjects/teste/GUIA_ARRANQUE_QT_LOCAL.md)
- [GUIA_INSTALACAO_TOTAL.md](c:/Users/engenharia/VSCodeProjects/teste/GUIA_INSTALACAO_TOTAL.md)
- [GUIA_INSTALACAO_OUTRO_PC.md](c:/Users/engenharia/VSCodeProjects/teste/GUIA_INSTALACAO_OUTRO_PC.md)
- [GUIA_MUITO_SIMPLES.md](c:/Users/engenharia/VSCodeProjects/teste/GUIA_MUITO_SIMPLES.md)
- [MANUAL_OPERACAO_LUISGEST_PROFISSIONAL.md](c:/Users/engenharia/VSCodeProjects/teste/MANUAL_OPERACAO_LUISGEST_PROFISSIONAL.md)
- [impulse_mobile_api/README.md](c:/Users/engenharia/VSCodeProjects/teste/impulse_mobile_api/README.md)
- [impulse_mobile_app/README.md](c:/Users/engenharia/VSCodeProjects/teste/impulse_mobile_app/README.md)

## Fluxo recomendado de desenvolvimento

1. Criar e ativar `.venv`
2. Configurar `lugest.env`
3. Validar MySQL
4. Arrancar a app desktop
5. Executar os scripts de verificação da área alterada

## Observações

- o projeto pode ter dados reais, ficheiros gerados e worktree suja; evita operações destrutivas
- parte da lógica histórica ainda vive em módulos legacy, enquanto a UI nova está em `lugest_qt/`
- o planeamento e os PDFs operacionais dependem muito da coerência entre operações, tempos e recursos atribuídos nas encomendas
