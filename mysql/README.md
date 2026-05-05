# MySQL Runtime

O runtime atual do LuGEST trabalha apenas com MySQL como fonte de verdade.

## Ficheiros principais
- `lugest.sql`: schema atual completo e consolidado para instalacao nova.
- `lugest_instalacao_unica.sql`: SQL unico com schema atual + utilizadores iniciais minimos.
- `IMPORTAR_NO_HEIDI.sql`: copia clara do SQL de instalacao nova para importar diretamente no HeidiSQL.
- `Migracoes\patch_*.sql`: patches incrementais para ambientes mais antigos.
- `install_lugest_mysql.py`: instalador/validador MySQL com suporte a `DryRun`.
- `install_lugest_mysql.ps1`: wrapper PowerShell para Windows Server.
- `apply_lugest_migrations.py`: aplica migracoes incrementais com registo em `schema_migrations`.
- `apply_lugest_migrations.ps1`: wrapper PowerShell para updates seguros em clientes em producao.
- `aplicar_migracoes_lugest_admin.bat`: atalho para aplicar updates incrementais.
- `estado_migracoes_lugest_admin.bat`: atalho para ver o estado do versionamento da base.
- `validate_lugest_mysql.ps1`: validacao rapida do schema final.
- `backup_lugest_mysql.py`: cria backups SQL da base em `.sql` ou `.sql.gz`.
- `restore_lugest_mysql.py`: restaura backups SQL da base.
- `backup_lugest_mysql.ps1` / `restore_lugest_mysql.ps1`: wrappers PowerShell para backup/restauro.
- `instalar_lugest_mysql_admin.bat`: atalho para arrancar a instalacao em Windows.
- `validar_lugest_mysql.bat`: atalho para validar a base em Windows.
- `backup_lugest_mysql_admin.bat`: atalho para backup em Windows.
- `restaurar_lugest_mysql_admin.bat`: atalho para restauro em Windows.

## Instalacao recomendada

Se preferires usar HeidiSQL numa instalacao nova:

1. Criar/abrir a ligacao ao servidor MySQL.
2. Abrir o ficheiro `IMPORTAR_NO_HEIDI.sql`.
3. Executar o script completo.
4. Configurar o `lugest.env` do Desktop App com os mesmos dados da base.

O ficheiro `IMPORTAR_NO_HEIDI.sql` e para instalacao nova. Nao usar por cima de uma base de cliente com dados reais.

Em Windows/PowerShell:

```powershell
cd 'Base de Dados\mysql'
powershell -ExecutionPolicy Bypass -File .\install_lugest_mysql.ps1 `
  -DbHost 127.0.0.1 `
  -Port 3306 `
  -AdminUser root `
  -AdminPassword 'PASSWORD_ADMIN' `
  -Database lugest `
  -AppUser lugest_user `
  -AppPassword 'PASSWORD_APP' `
  -AppHost localhost `
  -ResetDatabase
```

Notas:
- `-ResetDatabase` apaga e recria a base. Usa apenas em instalacao nova ou quando quiseres reinicializar tudo.
- Sem `-ResetDatabase`, o script tenta alinhar a base existente sem a apagar.
- `-DryRun` mostra o plano sem aplicar alteracoes.

## Atualizacao segura em clientes em producao

Regra principal:
- nunca reimportar `lugest.sql` ou `lugest_instalacao_unica.sql` por cima de uma base com dados reais.
- em cliente ja em uso, primeiro faz `backup`, depois aplica apenas migracoes incrementais.

Fluxo recomendado:

```powershell
cd 'Base de Dados\mysql'
powershell -ExecutionPolicy Bypass -File .\backup_lugest_mysql.ps1 `
  -DbHost 127.0.0.1 `
  -Port 3306 `
  -User lugest_user `
  -Password 'PASSWORD_APP' `
  -Database lugest `
  -Label pre_update

powershell -ExecutionPolicy Bypass -File .\apply_lugest_migrations.ps1 `
  -DbHost 127.0.0.1 `
  -Port 3306 `
  -AdminUser root `
  -AdminPassword 'PASSWORD_ADMIN' `
  -Database lugest
```

## Primeira vez a ativar o versionamento numa base ja em uso

Se a base do cliente ja estiver alinhada com a versao atual e so quiseres comecar a registar historico:

```powershell
cd 'Base de Dados\mysql'
powershell -ExecutionPolicy Bypass -File .\apply_lugest_migrations.ps1 `
  -DbHost 127.0.0.1 `
  -Port 3306 `
  -AdminUser root `
  -AdminPassword 'PASSWORD_ADMIN' `
  -Database lugest `
  -BaselineCurrent
```

Isto:
- cria a tabela `schema_migrations`
- regista os `Migracoes\patch_*.sql` atuais como baseline
- nao executa SQL de alteracao

## Base antiga sem historico e com patches por aplicar

Se tens uma base antiga e ainda precisas mesmo de correr todos os `Migracoes\patch_*.sql`, faz backup primeiro e usa:

```powershell
cd 'Base de Dados\mysql'
powershell -ExecutionPolicy Bypass -File .\apply_lugest_migrations.ps1 `
  -DbHost 127.0.0.1 `
  -Port 3306 `
  -AdminUser root `
  -AdminPassword 'PASSWORD_ADMIN' `
  -Database lugest `
  -LegacyApplyAll
```

Este modo:
- executa todos os `Migracoes\patch_*.sql` de forma tolerante
- regista cada patch em `schema_migrations`
- deve ser usado so depois de backup

## Estado do versionamento

```powershell
cd 'Base de Dados\mysql'
powershell -ExecutionPolicy Bypass -File .\apply_lugest_migrations.ps1 `
  -DbHost 127.0.0.1 `
  -Port 3306 `
  -AdminUser root `
  -AdminPassword 'PASSWORD_ADMIN' `
  -Database lugest `
  -Status
```

Ou, em Windows:

```bat
estado_migracoes_lugest_admin.bat
```

## Validacao

```powershell
cd 'Base de Dados\mysql'
powershell -ExecutionPolicy Bypass -File .\validate_lugest_mysql.ps1 `
  -DbHost 127.0.0.1 `
  -Port 3306 `
  -User lugest_user `
  -Password 'PASSWORD_APP' `
  -Database lugest
```

## Resultado esperado

Depois da instalacao, usa estes valores no `lugest.env` do desktop e no `.env` da API:

```text
LUGEST_DB_HOST=127.0.0.1
LUGEST_DB_PORT=3306
LUGEST_DB_USER=lugest_user
LUGEST_DB_PASS=<password_app>
LUGEST_DB_NAME=lugest
```

## Backup

Exemplo em PowerShell:

```powershell
cd 'Base de Dados\mysql'
powershell -ExecutionPolicy Bypass -File .\backup_lugest_mysql.ps1 `
  -DbHost 127.0.0.1 `
  -Port 3306 `
  -User lugest_user `
  -Password 'PASSWORD_APP' `
  -Database lugest `
  -Label pre_update
```

Isto cria por defeito uma pasta em `backups\mysql\...` com:
- dump `.sql.gz`
- `metadata.json`

## Restauro

Exemplo em PowerShell:

```powershell
cd 'Base de Dados\mysql'
powershell -ExecutionPolicy Bypass -File .\restore_lugest_mysql.ps1 `
  -DbHost 127.0.0.1 `
  -Port 3306 `
  -AdminUser root `
  -AdminPassword 'PASSWORD_ADMIN' `
  -Database lugest `
  -Input 'C:\Backups\lugest_20260320_120000_pre_update.sql.gz' `
  -ResetDatabase
```

Notas:
- `-ResetDatabase` apaga e recria a base antes do restauro.
- Sem `-Input`, o script tenta usar o dump mais recente em `backups\mysql`.
- Em `-DryRun`, o script so mostra o plano e a ferramenta necessaria.

## Nota
- O runtime atual e exclusivamente MySQL.
- `lugest_data.json` nao e fonte de dados runtime e nao deve acompanhar novas instalacoes.
- Em instalacoes novas, o ficheiro principal a importar manualmente e `lugest.sql`.
- Se quiseres um arranque rapido com login imediato, importa `lugest_instalacao_unica.sql`.
- O menu `Faturacao` ja segue incluido em `lugest.sql` e `lugest_instalacao_unica.sql`.
- Logins temporarios da instalacao unica: `admin`, `operador`, `orcamentista`, `planeamento`.
- O login `OWNER` do `lugest.env` nao substitui o `admin` local da aplicacao.
- Em instalacoes finais, recomenda-se repor logo o admin com o script `Criar Administrador Inicial.bat` ou `Repor Password Administrador.bat`.
- Os `Migracoes\patch_*.sql` ficam mantidos para upgrades/migracoes de bases antigas.
- Em cliente em producao, o caminho certo e `backup -> apply_lugest_migrations -> validar`, nunca `importar SQL unico por cima`.
