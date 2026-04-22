# luGEST Web Read-Only

Portal web de consulta para `cPanel`, pensado para ler a base MySQL do luGEST sem escrever dados.

## Objetivo

Esta aplicacao serve para:

- consultar dashboard operacional
- consultar encomendas e detalhe da encomenda
- consultar blocos de planeamento
- consultar clientes
- consultar produtos
- consultar materia-prima
- consultar orcamentos

Nao executa alteracoes na base. O objetivo e publicacao simples em dominio, com login e acesso apenas de leitura.

## Stack

- PHP 8.1+
- MySQL / MariaDB
- HTML + CSS + JS leve
- sessao PHP nativa

## Estrutura

- `config/config.example.php`: exemplo base de configuracao
- `config/config.cpanel.example.php`: exemplo preparado para `cPanel`
- `config/config.php`: configuracao local pronta a ler o `lugest.env`
- `public/`: ficheiros publicos do dominio
- `src/`: autenticacao, acesso MySQL, helpers e queries
- `sql/`: scripts opcionais para utilizador read-only e views
- `start_local_server.ps1`: arranque local em Windows
- `stop_local_server.ps1`: paragem local em Windows

## Teste local no teu PC

1. Executar `arrancar_portal_local.cmd`
2. O portal abre em `http://127.0.0.1:8088/login.php`
3. Entrar com um utilizador real do luGEST

Se quiseres arrancar manualmente:

```powershell
powershell -ExecutionPolicy Bypass -File .\start_local_server.ps1
```

Para parar:

```powershell
powershell -ExecutionPolicy Bypass -File .\stop_local_server.ps1
```

## Arranque rapido para deploy

1. Copiar `config/config.cpanel.example.php` para `config/config.php`
2. Preencher as credenciais MySQL
3. Apontar o dominio/subdominio para a pasta `public`
4. Garantir que o utilizador MySQL usado pela web tem apenas `SELECT`

## Deploy em cPanel

Opcao recomendada:

1. Criar subdominio, por exemplo `consultas.teudominio.pt`
2. Fazer upload da pasta `lugest_web_readonly`
3. Apontar o document root do subdominio para `lugest_web_readonly/public`
4. Criar `config/config.php`
5. Criar utilizador MySQL read-only
6. Testar login

Se o `cPanel` nao permitir apontar diretamente o document root para `public`, ainda podes publicar a pasta inteira e ajustar o acesso, mas o ideal e manter `src/` e `config/` fora da raiz publica.

## Seguranca recomendada

- usar um utilizador MySQL exclusivo da web com permissoes `SELECT`
- limitar por IP se fizer sentido
- ativar SSL no dominio
- nao reutilizar o utilizador MySQL da aplicacao desktop
- manter `config.php` fora do acesso publico sempre que possivel

## Login

Por defeito o portal autentica contra a tabela `users` do luGEST e usa o mesmo formato de password:

- `pbkdf2_sha256$iteracoes$salt$digest`

Tambem e possivel limitar os perfis aceites em `config.php`.

## Scripts SQL

- `sql/lugest_web_readonly_user.sql`: exemplo de utilizador MySQL so com leitura
- `sql/lugest_web_readonly_views.sql`: views opcionais para simplificar consultas do portal

## MVP incluido

- login
- dashboard
- encomendas
- detalhe da encomenda
- planeamento
- clientes
- produtos
- materia-prima
- orcamentos

## Proximos passos recomendados

1. adaptar branding ao cliente
2. afinar permissoes por perfil
3. adicionar paginacao
4. adicionar exportacao CSV/PDF
5. publicar em dominio
