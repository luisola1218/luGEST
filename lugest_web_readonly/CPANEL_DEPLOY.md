# Deploy cPanel

Guia rápido para publicar o portal `lugest_web_readonly` em `cPanel`.

## 1. Criar subdomínio

Exemplo:

- `consultas.teudominio.pt`

Idealmente o document root deve apontar para:

- `lugest_web_readonly/public`

## 2. Subir os ficheiros

Faz upload da pasta completa `lugest_web_readonly` para o alojamento.

## 3. Criar a configuração

Na pasta `config`:

1. copia `config.cpanel.example.php`
2. renomeia para `config.php`
3. ajusta:
   - `db.host`
   - `db.name`
   - `db.user`
   - `db.pass`
   - `allowed_roles`
   - branding

## 4. Criar utilizador MySQL read-only

Usa como base o ficheiro:

- `sql/lugest_web_readonly_user.sql`

Recomendação:

- nunca usar o mesmo utilizador MySQL do desktop
- criar um utilizador próprio da web só com `SELECT`

## 5. Opcional: criar views

Se quiseres separar melhor a leitura da web:

- executar `sql/lugest_web_readonly_views.sql`

## 6. Testar

Abrir:

- `https://consultas.teudominio.pt`

Entrar com um utilizador existente na tabela `users` do luGEST, desde que o perfil esteja autorizado.

## 7. Checklist de segurança

- SSL ativo
- utilizador MySQL só com `SELECT`
- `config.php` fora da pasta pública, se o alojamento permitir
- permissões mínimas nos ficheiros
- backups ativos da base de dados
