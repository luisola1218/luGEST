# Guia de Instalacao Total do LuGEST

Este guia foi escrito para uma instalacao completa desde o zero, de forma simples, para quem nunca fez este processo.

Objetivo final:
- ter o servidor pronto;
- ter a base de dados pronta;
- ter o LuGEST a abrir sem erros;
- ter todos os postos a trabalhar na mesma base de dados;
- ter anexos, PDFs e ficheiros partilhados entre todos os postos;
- ter a API mobile pronta, se o cliente a usar.

## 1. O cenario recomendado

Para um cliente normal, faz assim:
- `1 servidor principal`
- `1 ou mais postos`
- `1 base de dados MySQL`
- `1 pasta partilhada na rede`

No `servidor` ficam:
- MySQL
- LuGEST Desktop principal
- API mobile, se for usada
- pasta partilhada para PDFs, comprovativos, anexos e ficheiros

Nos `postos` ficam:
- LuGEST Desktop
- sem MySQL local
- sempre a apontar para o servidor

## 2. O que tens de ter antes de comecar

Antes de instalar, prepara isto:
- o computador que vai ser o `servidor`
- o nome ou IP do servidor
- acesso de administrador ao Windows
- instalador do MySQL Server, se ainda nao estiver instalado
- a pasta `App LuisGEST - Cliente`
- a password do MySQL que queres usar para a aplicacao
- as credenciais do `OWNER`, entregues separadamente pelo responsavel do projeto

Se houver varios postos, o ideal e o servidor ficar com:
- IP fixo
- nome de rede fixo
- firewall configurada

## 3. Ordem certa da instalacao

Segue sempre esta ordem:
1. Escolher o servidor.
2. Preparar a pasta partilhada.
3. Instalar ou validar o MySQL no servidor.
4. Importar a base de dados.
5. Configurar o LuGEST no servidor.
6. Abrir o LuGEST no servidor e validar login.
7. Configurar trial/licenca no servidor, se necessario.
8. Instalar e configurar cada posto.
9. Testar o trabalho em rede.
10. Configurar a API mobile, se o cliente a usar.

## 4. Preparar a pasta partilhada do servidor

No servidor, cria uma pasta para ficheiros partilhados. Exemplo:

```text
C:\LuGEST\Ficheiros
```

Depois partilha essa pasta na rede.

Exemplo de caminho final:

```text
\\SERVIDOR\LuGEST\Ficheiros
```

Essa pasta vai servir para:
- PDFs de faturacao
- comprovativos
- anexos
- outros ficheiros gerados pela app

Muito importante:
- todos os postos devem conseguir abrir esta pasta;
- o caminho tem de ser o mesmo para todos;
- nao uses uma pasta local diferente em cada posto.

## 5. Instalar ou validar o MySQL no servidor

Abre a pasta:

```text
Base de Dados\mysql
```

Tens duas formas de preparar a base.

### Opcao A - mais simples

Se vais fazer uma instalacao nova e tens acesso de administrador:

```text
instalar_lugest_mysql_admin.bat
```

Carrega duas vezes e segue o processo.

### Opcao B - manual

Se preferires fazer manualmente:
1. Instala o MySQL Server no servidor.
2. Cria a base de dados `lugest`.
3. Cria um utilizador dedicado para a aplicacao.
4. Nao uses `root` como utilizador normal da app.

Exemplo recomendado:

```text
Base de dados: lugest
Utilizador da app: lugest_user
Password da app: uma password forte tua
```

## 6. Importar a base de dados certa

Ainda na pasta:

```text
Base de Dados\mysql
```

Escolhe uma destas opcoes:

### Se e uma instalacao nova

Importa:

```text
lugest_instalacao_unica.sql
```

Usa esta opcao quando queres arrancar de raiz, com uma instalacao limpa e pronta.

### Se queres levar os dados reais do cliente

Importa:

```text
lugest.sql
```

Usa esta opcao quando queres restaurar os dados atuais da empresa.

### Depois de importar

Valida a ligacao com:

```text
validar_lugest_mysql.bat
```

Se der erro, normalmente e um destes pontos:
- MySQL desligado
- password errada
- utilizador errado
- firewall a bloquear
- base de dados nao foi importada

## 7. Configurar o LuGEST no servidor

Abre a pasta:

```text
Desktop App
```

No servidor, usa este ficheiro como base:

```text
lugest.env.servidor.example
```

Faz uma copia e grava com o nome:

```text
lugest.env
```

Depois abre `lugest.env` no Bloco de Notas e confirma estes campos:

```text
LUGEST_DB_HOST=127.0.0.1
LUGEST_DB_PORT=3306
LUGEST_DB_USER=lugest_user
LUGEST_DB_PASS=A_TUA_PASSWORD_MYSQL
LUGEST_DB_NAME=lugest
LUGEST_SHARED_STORAGE_ROOT=\\SERVIDOR\LuGEST\Ficheiros
```

Regras simples:
- no servidor, `LUGEST_DB_HOST` costuma ser `127.0.0.1`;
- `LUGEST_DB_PASS` tem de ser a password real do utilizador MySQL da app;
- `LUGEST_SHARED_STORAGE_ROOT` tem de apontar para a pasta partilhada do servidor.

## 8. Sobre o OWNER

No `lugest.env` tambem existe isto:

```text
LUGEST_OWNER_USERNAME=...
LUGEST_OWNER_PASSWORD=...
```

Isto serve para:
- trial
- licenciamento
- reautorizacao quando a maquina muda

Nao serve para:
- login normal dos utilizadores
- login do administrador normal da app

Importante:
- neste pacote, a password do `OWNER` vai em hash;
- se alguem abrir o `lugest.env`, nao consegue ler a password real do `OWNER`;
- a password real do `OWNER` deve ser guardada fora deste pacote e entregue separadamente.

## 9. Abrir o LuGEST no servidor pela primeira vez

Ainda em `Desktop App`:

### Se importaste `lugest_instalacao_unica.sql`

Abre:

```text
Arrancar LuisGEST Desktop.bat
```

Depois entra com:

```text
admin
Trocar#Admin2026
```

Assim que entrares:
- troca logo a password do `admin`;
- confirma que o dashboard abre;
- confirma que nao aparece erro MySQL.

### Se importaste `lugest.sql` e nao ha utilizadores locais

Corre primeiro:

```text
Criar Administrador Inicial.bat
```

Se precisares de redefinir esse admin:

```text
Repor Password Administrador.bat
```

Depois abre:

```text
Arrancar LuisGEST Desktop.bat
```

## 10. Trial ou licenca no servidor

Se aparecer aviso de trial bloqueado ou licenca presa a outro equipamento:
1. entra com o `OWNER`;
2. faz a reautorizacao nessa maquina;
3. confirma que o servidor ficou funcional.

Regra muito importante:
- nao copies um `lugest_trial.json` ja usado noutra maquina;
- o trial fica ligado ao equipamento final;
- faz sempre a ativacao no computador real do cliente.

## 11. Teste minimo no servidor

Antes de instalar os postos, confirma no servidor:
- a app abre;
- o login funciona;
- consegues abrir clientes;
- consegues abrir produtos;
- consegues abrir encomendas;
- consegues abrir transportes;
- consegues abrir faturacao.

Se possivel, faz tambem um teste simples:
- criar um registo de teste;
- criar uma fatura de teste;
- criar um pagamento de teste.

## 12. Instalar o primeiro posto

No posto cliente, copia a pasta:

```text
Desktop App
```

No posto, usa este ficheiro como base:

```text
lugest.env.posto.example
```

Faz uma copia com o nome:

```text
lugest.env
```

Depois preenche assim:

```text
LUGEST_DB_HOST=IP_DO_SERVIDOR
LUGEST_DB_PORT=3306
LUGEST_DB_USER=lugest_user
LUGEST_DB_PASS=A_TUA_PASSWORD_MYSQL
LUGEST_DB_NAME=lugest
LUGEST_SHARED_STORAGE_ROOT=\\SERVIDOR\LuGEST\Ficheiros
```

Exemplo:

```text
LUGEST_DB_HOST=192.168.1.100
```

Regras do posto:
- o `LUGEST_DB_HOST` tem de ser o IP ou nome do servidor;
- o `LUGEST_SHARED_STORAGE_ROOT` tem de ser igual ao do servidor;
- o `OWNER` deve ser o mesmo do servidor;
- nao instalar MySQL local em cada posto.

## 13. Abrir o primeiro posto

No posto, abre:

```text
Arrancar LuisGEST Desktop.bat
```

Se tudo estiver correto:
- aparece o login;
- consegues entrar com o utilizador normal;
- vês os mesmos dados do servidor.

## 14. Repetir nos restantes postos

Para os outros postos, faz exatamente o mesmo processo:
1. copiar `Desktop App`;
2. criar `lugest.env` a partir de `lugest.env.posto.example`;
3. meter o IP do servidor;
4. meter a password correta do MySQL;
5. confirmar a pasta partilhada;
6. abrir a app e testar login.

Quando tudo esta bem feito:
- todos os postos veem as mesmas encomendas;
- todos os postos veem os mesmos clientes;
- todos os postos acedem aos mesmos PDFs e anexos.

## 15. Teste de rede obrigatorio

Em cada posto, confirma isto:
- faz `ping` ao servidor;
- a pasta `\\SERVIDOR\LuGEST\Ficheiros` abre;
- a app entra sem erro;
- consegues abrir uma encomenda;
- consegues abrir uma fatura;
- consegues abrir transportes;
- consegues abrir documentos partilhados.

Se nao funcionar, verifica:
- IP errado no `lugest.env`
- password MySQL errada
- firewall do servidor
- partilha de rede sem permissoes
- MySQL parado

## 16. API mobile, se o cliente usar

No servidor, abre a pasta:

```text
Mobile API
```

Usa `.env.example` como base e grava como:

```text
.env
```

Preenche pelo menos:

```text
LUGEST_API_HOST=0.0.0.0
LUGEST_API_PORT=8050
LUGEST_API_SECRET=UMA_CHAVE_FORTE
LUGEST_DB_HOST=127.0.0.1
LUGEST_DB_PORT=3306
LUGEST_DB_USER=lugest_user
LUGEST_DB_PASS=A_TUA_PASSWORD_MYSQL
LUGEST_DB_NAME=lugest
```

Depois corre:

```text
instalar_impulse_api.bat
```

E depois:

```text
arrancar_impulse_api.bat
```

Se for para deixar no servidor em arranque automatico:

```text
instalar_impulse_api_arranque_automatico_admin.bat
```

## 17. Mobile Android, se o cliente usar

Se a pasta `Mobile APK` trouxer uma APK pronta:
- instala essa APK no telemovel;
- no primeiro arranque indica o endereco da API.

Exemplo:

```text
http://192.168.1.100:8050
```

Se nao existir APK pronta:
- usar o codigo da pasta `Mobile App Fonte`;
- gerar a APK mais tarde.

## 18. O que nunca deves fazer

Nao faças isto:
- instalar um MySQL diferente em cada posto;
- usar caminhos locais diferentes para PDFs e anexos;
- usar `root` como conta normal da aplicacao;
- copiar o `lugest_trial.json` ja ativado desta maquina para outra;
- deixar o `lugest.env` com um IP errado;
- esquecer a partilha UNC em ambiente multi-posto.

## 19. Checklist final de fecho

No fim da instalacao, confirma:
1. o servidor abre o LuGEST sem erros;
2. o primeiro posto entra sem erros;
3. os restantes postos entram sem erros;
4. todos veem a mesma base de dados;
5. todos conseguem abrir a mesma pasta partilhada;
6. faturacao abre;
7. transportes abre;
8. documentos e PDFs sao gerados sem erro;
9. a API mobile responde, se existir;
10. fazes um backup inicial logo apos a instalacao ficar valida.

## 20. Se quiseres a versao curta de tudo

Resumo:
1. Instalar MySQL no servidor.
2. Importar `lugest_instalacao_unica.sql` ou `lugest.sql`.
3. Criar a pasta partilhada `\\SERVIDOR\LuGEST\Ficheiros`.
4. Configurar `Desktop App\lugest.env` no servidor.
5. Abrir e testar o LuGEST no servidor.
6. Reautorizar trial no servidor, se preciso.
7. Configurar cada posto com `lugest.env.posto.example`.
8. Testar todos os postos.
9. Configurar API mobile, se o cliente a usar.
10. Fazer backup inicial.
