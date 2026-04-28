# Guia Muito Simples

Este guia e para instalar o LuGEST noutro computador como se fosses fazer isto pela primeira vez.

## O que tens de saber antes

O programa precisa de 2 coisas para funcionar:

1. Uma base de dados MySQL.
2. O ficheiro `lugest.env` bem preenchido.

Sem isso, o programa nao entra.

## Caso mais normal

Imagina isto:

- computador 1 = servidor
- computador 2 = computador onde vais abrir o LuGEST

Pode ser tudo no mesmo computador tambem. O processo e parecido.

---

## PARTE 1 - Preparar a base de dados

Abre a pasta:

`Base de Dados\mysql`

Depois faz assim:

### Se for uma instalacao nova

1. Abre o ficheiro:
   `instalar_lugest_mysql_admin.bat`
2. Carrega duas vezes.
3. Se o Windows perguntar permissoes, aceita.

Isto serve para criar a base de dados do LuGEST no MySQL.

Se preferires fazer manual:

1. Instala o MySQL Server.
2. Cria a base com nome:
   `lugest`
3. Cria um utilizador proprio para o programa.
4. Importa `lugest.sql` ou `lugest_instalacao_unica.sql`.

Exemplo:

- utilizador: `lugest_user`
- password: uma password tua

### Depois valida

Abre:

`validar_lugest_mysql.bat`

Se estiver tudo bem, a base ficou pronta.

---

## PARTE 2 - Configurar o programa

Abre a pasta:

`Desktop App`

La dentro tens o ficheiro:

`lugest.env`

Abre esse ficheiro com o Bloco de Notas.

Vais ver algo parecido com isto:

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

Agora preenche assim:

### Se o MySQL estiver no mesmo computador

Deixa:

```text
LUGEST_DB_HOST=127.0.0.1
```

### Se o MySQL estiver noutro computador ou servidor

Troca pelo IP do servidor.

Exemplo:

```text
LUGEST_DB_HOST=192.168.1.100
```

### O que tens mesmo de preencher

- `LUGEST_DB_HOST`
- `LUGEST_DB_PORT`
- `LUGEST_DB_USER`
- `LUGEST_DB_PASS`
- `LUGEST_DB_NAME`
- `LUGEST_SHARED_STORAGE_ROOT` se tiveres mais do que um posto
- `LUGEST_OWNER_USERNAME`
- `LUGEST_OWNER_PASSWORD`

### O que e o OWNER

O `OWNER` e o utilizador especial do dono ou administrador principal.

Nao e o mesmo que os utilizadores normais do programa.
Tambem nao e o mesmo que o `admin` local da aplicacao.

Exemplo simples:

```text
LUGEST_OWNER_USERNAME=proprietario
LUGEST_OWNER_PASSWORD=UmaPasswordMuitoForte#2026
```

Grava o ficheiro no fim.

Se tiveres varios postos a trabalhar ao mesmo tempo, mete tambem isto:

```text
LUGEST_SHARED_STORAGE_ROOT=\\SERVIDOR\LuGEST\Ficheiros
```

Assim todos abrem os mesmos desenhos, PDFs e anexos.

---

## PARTE 3 - Abrir o programa

Ainda dentro da pasta `Desktop App`:

1. Se importaste `lugest.sql` e ainda nao tens nenhum utilizador local, corre antes:
   `Criar Administrador Inicial.bat`
2. Se importaste `lugest_instalacao_unica.sql`, podes entrar com:
   `admin` / `Trocar#Admin2026`
3. Depois carrega duas vezes em:
   `Arrancar LuisGEST Desktop.bat`

Se tudo estiver bem:

- o programa abre;
- aparece o login;
- ja consegues entrar.

Se precisares de trocar a password do `admin`, usa:

`Repor Password Administrador.bat`

---

## PARTE 4 - Se nao abrir

Se der erro, normalmente e uma destas 4 coisas:

1. O MySQL nao esta ligado.
2. O IP no `lugest.env` esta errado.
3. O utilizador ou password do MySQL estao errados.
4. A firewall esta a bloquear.

### Coisa mais importante para testar

Verifica:

- se o MySQL esta a correr
- se a porta e `3306`
- se o nome da base e `lugest`
- se o utilizador e a password do MySQL estao certos

---

## PARTE 5 - Se fores usar telemovel ou app mobile

Abre a pasta:

`Mobile API`

La dentro:

1. abre o ficheiro `.env`
2. mete os mesmos dados MySQL
3. define uma chave forte para a API

Depois:

1. corre `instalar_impulse_api.bat`
2. corre `arrancar_impulse_api.bat`

Se for para servidor:

usa tambem:

`instalar_impulse_api_arranque_automatico_admin.bat`

Se precisares de gerar APK nova para o cliente:

```powershell
flutter pub get
flutter build apk --release --dart-define=LUGEST_DEFAULT_API_HOST=http://IP_DO_SERVIDOR:8050
```

---

## PARTE 6 - Resumindo muito

Se quiseres a versao mais curta de todas, e esta:

1. Instalar ou preparar MySQL.
2. Criar base `lugest`.
3. Importar `lugest.sql` ou `lugest_instalacao_unica.sql`.
4. Criar utilizador MySQL do programa.
5. Abrir `Desktop App\lugest.env`.
6. Preencher os dados MySQL.
7. Definir `LUGEST_OWNER_USERNAME` e `LUGEST_OWNER_PASSWORD`.
8. Criar ou repor o admin com `Criar Administrador Inicial.bat`.
9. Gravar.
10. Abrir `Arrancar LuisGEST Desktop.bat`.

---

## Exemplo real simples

Se tudo estiver no mesmo computador:

```text
LUGEST_DB_HOST=127.0.0.1
LUGEST_DB_PORT=3306
LUGEST_DB_USER=lugest_user
LUGEST_DB_PASS=MinhaPasswordSegura123!
LUGEST_DB_NAME=lugest
LUGEST_SHARED_STORAGE_ROOT=\\SERVIDOR\LuGEST\Ficheiros
LUGEST_OWNER_USERNAME=proprietario
LUGEST_OWNER_PASSWORD=OutraPasswordForte#2026
```

Depois:

1. guardar
2. abrir `Arrancar LuisGEST Desktop.bat`

E pronto.

---

## Muito importante

Nao copies para a nova maquina:

- passwords antigas desta maquina
- `.env` reais desta maquina
- ficheiros com segredos antigos

Na maquina nova, mete sempre passwords novas.
