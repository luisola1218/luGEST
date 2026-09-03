# Migrações versionadas

Esta pasta recebe alterações incrementais para bases **já instaladas**. O
ficheiro `mysql/lugest.sql` continua a ser usado apenas em instalações novas.

## Nome

Usar quatro algarismos e uma descrição em minúsculas:

```text
0001_adicionar_indice_orcamentos.sql
0002_normalizar_estado_encomendas.sql
```

Uma migração aplicada é imutável. Se precisar de correção, cria-se outra. O
runner valida o SHA-256 e recusa um ficheiro histórico alterado.

## Planeamento seguro

O comando normal é apenas de leitura:

```powershell
.\.venv\Scripts\python.exe .\mysql\migrate_lugest_mysql.py
```

## Aplicação

Antes de aplicar, criar e validar um backup. O caminho desse backup é obrigatório:

```powershell
.\.venv\Scripts\python.exe .\mysql\migrate_lugest_mysql.py `
  --apply `
  --backup-confirmed "C:\Backups luGEST\lugest_pre_update.sql" `
  --applied-by "suporte@lugest"
```

O runner usa bloqueio exclusivo MySQL e regista chave, checksum, data e operador
em `schema_migrations`. Como DDL MySQL pode fazer commit implícito, cada ficheiro
deve ser pequeno, idempotente e testado primeiro numa cópia da base.
