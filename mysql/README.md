# Base de dados MySQL do LuGEST

## Ficheiro SQL único

`lugest.sql` é a fonte canónica da base de dados. Contém:

- o schema completo e atual;
- todas as tabelas, índices e relações necessárias;
- utilizadores iniciais mínimos para uma instalação nova;
- listas base de operador e orçamentista.

Todas as instalações novas usam `utf8mb4` com `utf8mb4_unicode_ci`. O exportador
normaliza automaticamente tabelas históricas para não voltar a gerar schemas em
`utf8mb3`. Esta normalização do ficheiro de instalação não altera uma base já
existente.

Os antigos schemas parciais e `patch_*.sql` foram consolidados neste ficheiro. Não é
necessário importar vários SQL por ordem.

## Instalação nova

No HeidiSQL:

1. ligar ao servidor MySQL;
2. abrir `lugest.sql`;
3. executar o ficheiro completo;
4. configurar a aplicação com o utilizador dedicado;
5. trocar imediatamente as palavras-passe iniciais.

Também é possível usar o instalador:

```powershell
powershell -ExecutionPolicy Bypass -File .\install_lugest_mysql.ps1 `
  -DbHost 127.0.0.1 `
  -AdminUser root `
  -AdminPassword "PASSWORD_ADMIN" `
  -Database lugest `
  -AppUser lugest_user `
  -AppPassword "PASSWORD_FORTE"
```

Para apagar e recriar uma base de testes, acrescentar `-ResetDatabase`. Nunca usar
essa opção numa base real com dados.

## Validação

```powershell
powershell -ExecutionPolicy Bypass -File .\validate_lugest_mysql.ps1 `
  -DbHost 127.0.0.1 `
  -AdminUser root `
  -AdminPassword "PASSWORD_ADMIN" `
  -Database lugest
```

## Atualizações de bases existentes

As alterações incrementais vivem em `mysql/migrations/` e são planeadas pelo
runner versionado:

```powershell
.\.venv\Scripts\python.exe .\mysql\migrate_lugest_mysql.py
```

Sem `--apply`, o comando é apenas de leitura. A aplicação exige um caminho de
backup existente e regista checksum e responsável em `schema_migrations`. Consulta
`mysql/migrations/README.md` antes de criar ou aplicar uma migração.

## Backup e reposição

Antes de qualquer alteração numa instalação existente:

```powershell
powershell -ExecutionPolicy Bypass -File .\backup_lugest_mysql.ps1
```

Para repor um backup validado:

```powershell
powershell -ExecutionPolicy Bypass -File .\restore_lugest_mysql.ps1
```

O `lugest.sql` destina-se sobretudo a instalações novas. Uma base de cliente com
dados deve ser sempre salvaguardada antes de uma atualização da aplicação.
Conversões de charset em bases existentes exigem backup validado, espaço livre e
janela de manutenção; não devem ser executadas automaticamente pelo desktop.

## Gerar novamente o SQL canónico

Com acesso à base de desenvolvimento atual:

```powershell
python .\export_current_schema_sql.py --with-starter-users --output .\lugest.sql
```

O pacote comercial cria uma cópia com o nome `IMPORTAR_NO_HEIDI.sql`, mantendo
apenas um SQL para o cliente importar.
