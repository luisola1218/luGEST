# Guia de Instalacao em Outro Computador

Este guia serve para instalar o LuGEST noutra maquina sem copiar segredos nem configuracoes desta instalacao.

## Ordem recomendada
1. Preparar o MySQL.
2. Importar a base de dados.
3. Configurar o desktop.
4. Testar login e operacao.
5. Configurar a API mobile, se for usada.
6. Gerar ou instalar a APK mobile, se necessario.

## 1. Preparar o servidor MySQL

Se for uma instalacao nova em Windows:

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
- Usa um utilizador dedicado, por exemplo `lugest_user`.
- Nao uses `root` como conta da aplicacao.
- Se a base ja existir e nao quiseres apagar tudo, remove `-ResetDatabase`.

## 2. Importar a base de dados

Se vais levar os dados atuais:
- importar `Base de Dados\mysql\lugest.sql` se for restauracao dos dados atuais;
- ou importar `Base de Dados\mysql\lugest_instalacao_unica.sql` se quiseres uma instalacao nova ja pronta de raiz;
- aplicar os `patch_*.sql` da mesma pasta se o ambiente destino for mais antigo;
- se fores restaurar um backup completo, usar `restore_lugest_mysql.ps1`.

Depois valida:

```powershell
cd 'Base de Dados\mysql'
powershell -ExecutionPolicy Bypass -File .\validate_lugest_mysql.ps1 `
  -DbHost 127.0.0.1 `
  -Port 3306 `
  -User lugest_user `
  -Password 'PASSWORD_APP' `
  -Database lugest
```

## 3. Configurar o desktop

Na pasta `Desktop App`:

1. Abrir `lugest.env`.
2. Preencher:

```text
LUGEST_DB_HOST=127.0.0.1
LUGEST_DB_PORT=3306
LUGEST_DB_USER=lugest_user
LUGEST_DB_PASS=PASSWORD_APP
LUGEST_DB_NAME=lugest
LUGEST_SHARED_STORAGE_ROOT=\\SERVIDOR\LuGEST\Ficheiros
LUGEST_OWNER_USERNAME=proprietario
LUGEST_OWNER_PASSWORD=DEFINE_AQUI_UMA_PASSWORD_FORTE
```

Notas:
- `LUGEST_OWNER_USERNAME` e `LUGEST_OWNER_PASSWORD` sao o login do proprietario para reautorizar trial, nao substituem os utilizadores normais da aplicacao.
- Define uma password forte tua no destino.
- Nao copies o `lugest.env` desta maquina para o destino.
- O login do OWNER nao e o login admin normal da aplicacao.
- Em ambiente multi-posto, define `LUGEST_SHARED_STORAGE_ROOT` com um caminho UNC partilhado por todos os postos.

Opcional de contingencia para instalacao limpa sem utilizadores locais:

```text
LUGEST_BOOTSTRAP_ADMIN_USERNAME=admin_inicial
LUGEST_BOOTSTRAP_ADMIN_PASSWORD=PasswordMuitoForte#2026
LUGEST_BOOTSTRAP_ADMIN_ROLE=Admin
```

Mas o recomendado e isto:
- usar `lugest_instalacao_unica.sql` e entrar com o admin temporario;
- ou correr `Criar Administrador Inicial.bat` logo apos configurar o `lugest.env`.

## 4. Arrancar o desktop

Na pasta `Desktop App`:

Se importaste `lugest.sql` e a base ficou sem utilizadores locais:

```text
Criar Administrador Inicial.bat
```

Se importaste `lugest_instalacao_unica.sql`, podes entrar primeiro com:
- `admin` / `Trocar#Admin2026`

Depois troca logo essa password ou usa:

```text
Repor Password Administrador.bat
```

```text
Arrancar LuisGEST Desktop.bat
```

Ou abrir diretamente:

```text
main.exe
```

## 5. Validacao minima apos arranque

Confirmar:
- a app abre sem erro;
- consegues fazer login;
- clientes, encomendas e produtos aparecem;
- o modulo de montagem abre;
- o dashboard carrega;
- consegues abrir `Notas Encomenda`;
- consegues abrir `Transportes`;
- consegues abrir `Faturacao`;
- consegues criar um registo com uma fatura e um pagamento de teste.

## 5A. Validacao minima do menu Transportes

Depois de instalares, testa pelo menos isto:
- criar ou abrir uma `Encomenda` e definir `Transporte a Nosso Cargo`;
- preencher `Local descarga`;
- abrir o menu `Transportes`;
- criar uma `Viagem`;
- afetar essa encomenda a viagem;
- abrir a `Folha de rota PDF`.

Se isto funcionar, a ponte `Encomenda -> Transporte -> PDF` ficou correta.

Se houver erro de ligacao:
- confirmar IP do MySQL;
- confirmar porta `3306`;
- confirmar user e password da base;
- confirmar firewall do Windows ou servidor.

## 6. Configurar a API mobile

Na pasta `Mobile API`:

1. Abrir `.env`.
2. Preencher:

```text
LUGEST_API_HOST=0.0.0.0
LUGEST_API_PORT=8050
LUGEST_API_SECRET=DEFINE_AQUI_UMA_CHAVE_FORTE
LUGEST_DB_HOST=127.0.0.1
LUGEST_DB_PORT=3306
LUGEST_DB_USER=lugest_user
LUGEST_DB_PASS=PASSWORD_APP
LUGEST_DB_NAME=lugest
```

3. Instalar:

```text
instalar_impulse_api.bat
```

4. Testar arranque:

```text
arrancar_impulse_api.bat
```

5. Se for servidor Windows, ativar arranque automatico:

```text
instalar_impulse_api_arranque_automatico_admin.bat
```

## 7. Mobile Android

Se existir APK em `Mobile APK`, instala essa.

Se nao existir:
- abrir `Mobile App Fonte`;
- instalar Flutter SDK;
- correr:

```powershell
flutter pub get
flutter build apk --release
```

Opcional para deixar a APK ja com o servidor API do cliente predefinido:

```powershell
flutter build apk --release --dart-define=LUGEST_DEFAULT_API_HOST=http://IP_DO_SERVIDOR:8050
```

## 8. O que nao deves copiar desta maquina

Nao levar:
- `lugest.env` real desta maquina;
- `impulse_mobile_api\.env` real desta maquina;
- `lugest_trial.json` desta maquina se ja estiver em uso;
- passwords locais antigas;
- ficheiros de backup sensiveis que nao sejam precisos.

## 9. Se o destino for servidor mais postos cliente

Cenario recomendado:
- MySQL e API no servidor;
- pasta partilhada UNC no servidor para `LUGEST_SHARED_STORAGE_ROOT`;
- Desktop LuGEST nos postos;
- cada posto com `lugest.env` a apontar para o IP do servidor;
- mobile a apontar para o IP e porta da API.

Exemplo:

```text
LUGEST_DB_HOST=192.168.1.100
LUGEST_DB_PORT=3306
LUGEST_API_HOST=0.0.0.0
LUGEST_API_PORT=8050
```

## 10. Check final

No destino, no fim:
- validar base;
- abrir desktop;
- testar login;
- testar uma encomenda;
- testar uma nota de encomenda;
- testar o fluxo de montagem;
- testar o fluxo de transportes;
- testar o fluxo de faturacao;
- testar a API mobile se for usada;
- fazer um backup inicial logo apos a instalacao ficar valida.
