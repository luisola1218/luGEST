# Arranque Qt local

## Setup inicial

Na raiz do projeto:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements-qt.txt
```

## Ativar e arrancar

```powershell
.\.venv\Scripts\Activate.ps1
python main.py
```

Tambem funciona:

```powershell
py main.py
```

Se a `.venv` local da raiz estiver pronta, `py main.py` reencaminha automaticamente para ela, mesmo que outra `.venv` esteja ativa no terminal.
Se a `.venv` local nao existir ou estiver incompleta, use explicitamente:

```powershell
.\.venv\Scripts\python.exe main.py
```

## Batch oficial

O arranque recomendado em Windows continua a ser:

```powershell
.\arrancar_lugest_qt.bat
```

Ordem de prioridade:

1. `.venv\Scripts\python.exe main.py`
2. `dist_qt_stable\lugest_qt\lugest_qt.exe`
3. `dist\lugest_qt\lugest_qt.exe`
4. `dist\main.exe`
5. `main.exe`
6. `py main.py`

Se a `.venv` existir mas estiver incompleta, o batch para e mostra o comando de instalacao em vez de cair num traceback.
