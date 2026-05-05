# Guia simples: email e atualizacoes

Este guia serve para instalar o LuisGEST num cliente sem depender de conhecimentos tecnicos avancados.

## Email

O LuisGEST prepara emails atraves do Microsoft Outlook instalado no computador.

Para funcionar bem:

1. Instalar o Microsoft Outlook desktop no posto que vai enviar emails.
2. Abrir o Outlook uma vez e configurar a conta da empresa.
3. Confirmar que o Outlook consegue enviar um email normal.
4. No LuisGEST, preencher os emails dos clientes e fornecedores nas respetivas fichas.
5. Ao enviar orcamentos, prazos ou notas de encomenda, o LuisGEST abre/cria o email no Outlook com o PDF anexado.

Se o Outlook nao estiver disponivel, o sistema tenta abrir o programa de email predefinido do Windows, mas nesse modo os anexos podem nao entrar automaticamente. Para uso profissional, usar Outlook desktop.

## Atualizacoes

Para clientes reais, a forma mais segura e controlada e trabalhar com pacotes de versao.
O LuisGEST ja segue com um atualizador preparado.

Fluxo recomendado:

1. O programador cria uma nova versao do LuisGEST.
2. A versao e testada numa base de dados de teste.
3. E gerada uma pasta de entrega nova no Ambiente de Trabalho.
4. A pasta `Atualizacoes` dessa entrega contem:
   - `latest.json`
   - `LuisGEST-Desktop-*.zip`
5. Esses dois ficheiros sao colocados numa pasta partilhada, GitHub Release privado ou servidor proprio.
6. No cliente, o menu `Extras > Atualizacoes` aponta para o `latest.json`.
7. O operador/admin carrega em `Verificar` e depois `Atualizar agora`.
8. O atualizador cria backup da pasta do programa e tenta criar backup MySQL com `mysqldump`.
9. O LuisGEST e fechado, os ficheiros sao substituidos e a nova versao arranca.

## GitHub ou servidor proprio?

GitHub e bom para guardar codigo e criar releases privadas, mas nao deve ser usado como atualizacao automatica sem controlo.

Para vender a clientes, a solucao recomendada e:

- GitHub privado para desenvolvimento e historico.
- Releases fechadas, assinadas e com numero de versao.
- Um manifest de atualizacao controlado, por exemplo `latest.json`, com versao, ficheiro, checksum e notas.
- O botao `Extras > Atualizacoes` consulta esse manifest, verifica checksum SHA256 e arranca o instalador externo.

O ideal em producao e o cliente nao instalar nada sem copia de seguranca e sem validacao da base de dados.

## Configuracao minima no cliente

Cada cliente deve ter:

- MySQL instalado no servidor.
- Base de dados LuisGEST importada.
- `lugest.env` configurado nos postos.
- Pasta partilhada para ficheiros, desenhos e PDFs, se houver varios utilizadores.
- Outlook configurado nos postos que enviam emails.
- Utilizadores criados no LuisGEST.
- Procedimento de backup diario da base de dados.

## Antes de atualizar

Confirmar sempre:

1. Backup da base de dados.
2. Backup da pasta atual do programa.
3. Nenhum utilizador esta a trabalhar no sistema.
4. Nova versao foi testada antes.
5. Existe forma de voltar atras caso algo falhe.

## Como configurar no cliente

1. Abrir LuisGEST como admin.
2. Ir a `Extras`.
3. Abrir `Atualizacoes`.
4. Em `Manifest`, indicar o caminho/URL do `latest.json`.

Exemplos:

```text
\\SERVIDOR\LuGEST\Atualizacoes\latest.json
https://exemplo.pt/lugest/latest.json
https://github.com/EMPRESA/REPO/releases/download/v2026.05.05/latest.json
```

Se for GitHub privado, preencher o token no campo `Token GitHub`.
O token deve ter apenas permissao de leitura da release.

Passo a passo simples com GitHub:

1. Criar um repositório privado para releases do LuisGEST.
2. Criar uma Release, por exemplo `v2026.05.05.1`.
3. Enviar como ficheiros da Release:
   - `latest.json`
   - `LuisGEST-Desktop-2026-05-05-1.zip`
4. No campo `Manifest`, usar o URL do `latest.json` dessa Release.
5. No campo `Token GitHub`, colar o token.
6. Carregar em `Guardar`.
7. Carregar em `Verificar`.
8. Se aparecer nova versão, carregar em `Atualizar agora`.

Exemplo:

```text
https://github.com/EMPRESA/lugest-releases/releases/download/v2026.05.05.1/latest.json
```

## Como publicar uma nova versao

1. Gerar a release com `scripts\prepare_final_release.ps1`.
2. Copiar os ficheiros da pasta `Atualizacoes` para o local usado pelo cliente.
3. Garantir que o `latest.json` e o `.zip` ficam na mesma pasta, ou ajustar `package_url`.
4. No cliente, carregar em `Verificar`.
