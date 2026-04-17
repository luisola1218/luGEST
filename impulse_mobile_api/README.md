# Lugest Impulse Mobile API

API separada, apenas de leitura, para expor o `Production Pulse` e o estado das encomendas a uma app móvel.

Esta pasta nao altera a aplicacao desktop. Reaproveita a mesma ligacao MySQL e a mesma logica base do Pulse para evitar divergencias entre desktop e mobile.

## Estrutura

- `app/main.py`: entrada FastAPI
- `app/config.py`: leitura do `.env`
- `app/security.py`: autenticacao simples por token assinado
- `app/services/pulse_runtime.py`: ponte para `main.py` e `menu_rooting.py`

## Variaveis

Copiar `.env.example` para `.env` e ajustar:

- `LUGEST_DB_HOST`
- `LUGEST_DB_PORT`
- `LUGEST_DB_USER`
- `LUGEST_DB_PASS`
- `LUGEST_DB_NAME`
- `LUGEST_API_HOST`
- `LUGEST_API_PORT`
- `LUGEST_API_SECRET`
- `LUGEST_API_ALLOWED_ORIGINS`

## Instalar

```powershell
cd impulse_mobile_api
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python verify_installation.py
```

## Arrancar

```powershell
cd impulse_mobile_api
.venv\Scripts\activate
python verify_installation.py
uvicorn app.main:app --host 0.0.0.0 --port 8050
```

## Producao no Windows Server

Para deixar a API a arrancar automaticamente no boot do servidor:

```powershell
cd impulse_mobile_api
powershell -ExecutionPolicy Bypass -File .\install_impulse_api_startup_task.ps1 -OpenFirewall
```

Scripts disponiveis:

- `install_impulse_api_startup_task.ps1`: cria a tarefa agendada em `SYSTEM` e arranca a API.
- `remove_impulse_api_startup_task.ps1`: remove a tarefa; opcionalmente remove a regra de firewall.
- `status_impulse_api_startup_task.ps1`: mostra estado, ultima execucao, porta e logs.
- `run_impulse_api_server.ps1`: runner com logs e reinicio simples em caso de crash.
- `instalar_impulse_api_arranque_automatico_admin.bat`: atalho para instalar o arranque automatico.
- `remover_impulse_api_arranque_automatico_admin.bat`: atalho para remover o arranque automatico.
- `estado_impulse_api_arranque.bat`: atalho para ver estado e logs.

Exemplos:

```powershell
powershell -ExecutionPolicy Bypass -File .\status_impulse_api_startup_task.ps1
powershell -ExecutionPolicy Bypass -File .\remove_impulse_api_startup_task.ps1 -RemoveFirewall
```

## Endpoints

- `POST /api/v1/auth/login`
- `GET /api/v1/health`
- `GET /api/v1/pulse/dashboard`
- `GET /api/v1/pulse/encomendas`
- `GET /api/v1/pulse/encomendas/{numero}`

## Notas

- Esta API depende do mesmo ambiente Python da app principal porque reaproveita `main.py` e `menu_rooting.py`.
- Nao arranca nenhuma janela Tk; apenas importa a logica e a leitura MySQL.
- `LUGEST_API_SECRET` deve ser configurado com uma chave forte; sem isso o login da API fica bloqueado.
- `LUGEST_API_ALLOWED_ORIGINS` so deve ser definido se existir frontend browser. A app mobile nativa nao precisa de `CORS` aberto.
- `python verify_installation.py` confirma `.env`, chave API e ligacao MySQL antes de arrancar.
- Os logs de runtime ficam em `api_stdout.log` e `api_stderr.log`.
