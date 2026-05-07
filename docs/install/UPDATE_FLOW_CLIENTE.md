# Fluxo Correto De Update Do Cliente

Este ficheiro existe para evitar regressões no mecanismo de atualização do desktop.

## Regra principal

O cliente **nao** deve depender diretamente do updater antigo que ja esta instalado.

O fluxo que ficou validado em cliente foi este:

1. A app le o `latest.json` para saber se existe versao nova.
2. A app descarrega o asset remoto `bootstrap_url`.
3. Esse asset deve ser o reparador:
   `Reparar_Atualizador_Instalado.ps1`
4. A app grava esse reparador novo dentro da pasta instalada do cliente.
5. A app executa o reparador **local** atualizado.
6. O reparador local trata de renovar os restantes scripts de update e arrancar a atualizacao real.

## Porque este fluxo e o certo

Foi o unico fluxo validado de forma consistente em cliente real.

Quando a atualizacao falhava, o procedimento manual que funcionava era:

1. copiar o `Reparar Atualizador Instalado.ps1` novo para a pasta instalada
2. substituir o antigo
3. executar esse reparador
4. deixar o reparador atualizar a aplicacao

Por isso o software passou a automatizar exatamente essa logica.

## O que nao devemos voltar a fazer

Nao voltar a estas abordagens sem revalidar em cliente:

- correr apenas `Atualizar LuisGEST.ps1` diretamente a partir da app
- depender apenas do `update_config.json` antigo da instalacao
- executar um bootstrap remoto isolado sem primeiro atualizar o reparador local
- usar nomes de assets com espacos quando isso puder afetar URLs
- copiar scripts por cima deles proprios sem verificar se origem e destino sao o mesmo ficheiro

## Assets esperados na release

Cada release desktop deve publicar:

- `latest.json`
- `LuisGEST-Desktop-<versao>.zip`
- `Reparar_Atualizador_Instalado.ps1`

## Manifest recomendado no cliente

Manter o cliente apontado para:

`https://github.com/luisola1218/luGEST/releases/latest/download/latest.json`

Assim nao e preciso mudar manualmente o URL a cada release.

## Checklist rapida antes de publicar

1. subir `VERSION`
2. rebuild dos binarios
3. correr `scripts\\prepare_final_release.ps1`
4. confirmar em `Atualizacoes\\latest.json`:
   - `version`
   - `package_url`
   - `bootstrap_url`
5. confirmar que o asset do reparador nao tem espacos no nome
6. publicar os 3 assets da release

## Checklist rapida de teste em cliente

1. `Verificar` encontra a nova versao
2. `Atualizar agora` arranca o reparador
3. o reparador nao falha por self-overwrite
4. a app reinicia com a alteracao visual esperada

## Ficheiros criticos

- `lugest_qt/services/main_bridge.py`
- `scripts/repair_installed_updater.ps1`
- `scripts/prepare_final_release.ps1`
- `scripts/lugest_update.ps1`

Se alguem alterar estes ficheiros, deve reler este documento antes.
