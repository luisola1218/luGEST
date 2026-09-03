# Arranque Qt local

## Setup inicial

Na raiz do projeto:

```powershell
py -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements-qt.txt
```

Em preparação de release ou ao reproduzir um erro de cliente, instalar antes
`requirements-qt.lock.txt` para usar exatamente as versões validadas.

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

## Arranque recomendado no desenvolvimento

Na raiz do projeto:

```powershell
.\.venv\Scripts\python.exe main.py
```

O ficheiro `main.py` valida a configuração MySQL e mantém compatibilidade com as
rotas legacy ainda em migração. A entrada Qt real está em `lugest_qt/app.py`.

Em cliente deve ser usado o atalho criado pelo instalador do pacote comercial;
não se arranca a partir da pasta de desenvolvimento nem se copia apenas o EXE.
