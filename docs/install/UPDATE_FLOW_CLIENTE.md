# Fluxo Correto De Update Do Cliente

Este ficheiro existe para evitar regressões no mecanismo de atualização do desktop.

## Regra principal

O cliente **nao** deve depender diretamente do updater antigo que ja esta instalado.

O fluxo que ficou validado em cliente foi este:

1. A app le o `latest.json` para saber se existe versao nova.
2. A app valida o formato, os URLs HTTPS e os dois hashes SHA-256 obrigatorios.
3. A app descarrega o asset remoto `bootstrap_url`.
4. Esse asset deve ser o reparador:
   `Reparar_Atualizador_Instalado.ps1`
5. O hash real do reparador tem de coincidir com `bootstrap_sha256` antes de substituir o ficheiro instalado.
6. A app grava esse reparador novo dentro da pasta instalada do cliente.
7. A app executa o reparador **local** atualizado.
8. O reparador valida o SHA-256 do ZIP, renova os restantes scripts e arranca a atualizacao real.

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

O ZIP comercial inclui ainda `Atualizar LuisGEST.ps1/.bat` e
`Reparar Atualizador Instalado.ps1/.bat`; o empacotador valida estas inclusoes.

## Campos obrigatorios do manifesto

- `schema_version`: atualmente `1`
- `version`
- `channel`: `stable`, `beta` ou `pilot`
- `package_url`: HTTPS ou caminho local controlado
- `sha256`: hash completo do ZIP
- `bootstrap_url`: HTTPS ou caminho local controlado
- `bootstrap_sha256`: hash completo do reparador
- `notes`: opcional

HTTP simples e hashes vazios ou incompletos sao recusados. SHA-256 deteta
alteracao/corrupcao, mas nao substitui a assinatura digital do executavel e do
instalador, que continua obrigatoria antes da venda geral.

O manifesto deve ser gerado pelo script, evitando copiar hashes a mao:

```powershell
.\.venv\Scripts\python.exe scripts\create_update_manifest.py `
  --version 2026.09.04.1 `
  --package-file .\LuisGEST-Desktop-2026.09.04.1.zip `
  --package-url https://updates.exemplo.pt/LuisGEST-Desktop-2026.09.04.1.zip `
  --bootstrap-file .\scripts\repair_installed_updater.ps1 `
  --bootstrap-url https://updates.exemplo.pt/Reparar_Atualizador_Instalado.ps1 `
  --output .\latest.json
```

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
   - `sha256`
   - `bootstrap_url`
   - `bootstrap_sha256`
5. confirmar que o asset do reparador nao tem espacos no nome
6. publicar os 3 assets da release
7. adulterar uma copia do ZIP e confirmar que e recusada no ensaio

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
