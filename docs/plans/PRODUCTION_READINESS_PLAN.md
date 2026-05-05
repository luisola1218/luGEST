# Plano de Prontidao para Cliente

Este documento define o caminho seguro para levar o LuGEST para clientes reais,
com multiutilizador, base de dados central, instalacao controlada e atualizacoes
sem risco para os dados.

## Principio

Nao reestruturar tudo de uma vez.

O sistema ja tem muitos fluxos operacionais em funcionamento. A prioridade agora
e estabilizar a base tecnica para producao:

- MySQL como fonte de verdade.
- Configuracao local separada dos dados da empresa.
- Migracoes versionadas.
- Backups antes de qualquer update.
- Instalador/release sem segredos da maquina de desenvolvimento.
- Testes ponta a ponta antes de entregar a cada cliente.

## Estado Atual

O projeto ja tem uma boa base para instalacao profissional:

- Desktop Qt em `lugest_qt_app.py`.
- Build PyInstaller em `lugest_qt.spec`.
- Configuracao por `lugest.env`.
- MySQL ativo no runtime.
- Scripts de instalacao, backup, restore e migracoes em `mysql/`.
- Guias de instalacao em `docs/install/`.
- Script de release em `scripts/prepare_final_release.ps1`.
- Auditoria de seguranca em `scripts/security_audit.py`.
- Testes funcionais em `scripts/verify_*.py`.

Ainda ha zonas que devem ser tratadas antes de vender:

- `main.py` e `runtime_pages.py` continuam grandes e misturam muita logica.
- Alguns dados/configuracoes ainda passam por JSON local ou `app_config`.
- O schema MySQL existe, mas a disciplina de migracoes tem de ser obrigatoria.
- O processo de update ainda deve ser fechado num fluxo simples para cliente.
- O licenciamento atual e trial/local; precisa de regras comerciais claras.

## Decisao de Arquitetura

### Servidor do cliente

Instalar:

- MySQL/MariaDB.
- Base `lugest`.
- Utilizador MySQL dedicado, nunca `root`.
- Pasta partilhada para ficheiros, se houver varios postos.
- Backups automaticos da base.

### Postos de trabalho

Cada posto deve ter:

- Executavel LuGEST.
- `lugest.env` local apontado para o servidor.
- Branding local.
- Sem base local com dados reais.
- Sem passwords embutidas no codigo.

### Base de dados

Para cliente real:

- MySQL e a fonte principal.
- JSON local serve apenas para configuracao/fallback controlado.
- Dados produtivos nao devem depender de ficheiros locais.

## Modelo de Dados Recomendado

Manter o modelo atual, mas consolidar por fases.

### Fase 1: estabilizar sem partir

Obrigatorio antes da primeira venda:

- Confirmar que `encomendas`, `orcamentos`, `clientes`, `materiais`,
  `produtos`, `plano`, `op_eventos`, `faturacao`, `expedicoes`,
  `notas_encomenda`, `operations_catalog` e `workcenter_catalog` persistem
  corretamente em MySQL.
- Garantir `schema_migrations` em todas as bases de cliente.
- Nunca aplicar `lugest.sql` por cima de base em producao.
- Usar sempre:
  `backup -> apply_lugest_migrations -> validate -> testes funcionais`.

### Fase 2: normalizacao incremental

Criar migracoes para tabelas mais relacionais quando o fluxo ja estiver
estavel:

- `operations`
  - `id`
  - `name`
  - `active`
  - `planeavel`

- `workcenters`
  - `id`
  - `operation_id`
  - `name`
  - `active`

- `fabrication_orders`
  - `id`
  - `order_id`
  - `code`
  - `status`
  - `created_at`

- `piece_production_refs`
  - `id`
  - `fabrication_order_id`
  - `piece_id`
  - `opp_code`

Esta fase deve ser feita por migracoes e adaptadores, nao por substituicao
brusca do backend.

### Fase 3: regras de integridade

Depois da normalizacao:

- `UNIQUE` em nomes de operacoes e postos.
- `FOREIGN KEY` posto -> operacao.
- Bloqueio de delete quando ha uso produtivo.
- Indices para encomenda, cliente, OF, OPP, planeamento e datas.

## Protecao Comercial

Antes de vender:

- `lugest.env` nunca entra no Git nem no pacote final com dados reais.
- Cada cliente recebe credenciais DB proprias.
- OWNER serve apenas para licenca/trial, nao para administracao normal.
- Administrador normal e criado no cliente com `setup_lugest_admin.ps1`.
- Trial/licenca deve ficar associado ao cliente/equipamento.
- Backups e restore devem ser testados no cliente antes de arrancar.

Recomendado para a fase comercial:

- Chave de licenca por cliente.
- Numero maximo de postos/utilizadores.
- Validade ou contrato ativo.
- Assinatura local da licenca para evitar edicao manual.
- Processo de reativacao controlado pelo OWNER.

## Atualizacoes em Cliente

Nao atualizar cliente diretamente a partir da pasta de desenvolvimento.

Fluxo recomendado:

1. Criar branch/release no Git.
2. Correr testes locais.
3. Gerar executavel.
4. Preparar pacote com `scripts/prepare_final_release.ps1`.
5. Fazer backup no cliente.
6. Aplicar migracoes MySQL.
7. Substituir executavel/pasta app.
8. Validar arranque e fluxos principais.

Em cliente:

```powershell
powershell -ExecutionPolicy Bypass -File .\backup_lugest_mysql.ps1 -Label pre_update
powershell -ExecutionPolicy Bypass -File .\apply_lugest_migrations.ps1
powershell -ExecutionPolicy Bypass -File .\validate_lugest_mysql.ps1
```

Nunca:

- Fazer `git pull` diretamente numa instalacao de cliente final.
- Copiar `lugest.env` da maquina de desenvolvimento.
- Copiar `lugest_trial.json` de outro cliente.
- Importar SQL completo por cima de dados reais.

## Checklist Antes da Primeira Venda

### Funcional

- Criar cliente.
- Criar orcamento.
- Criar conjunto/modelo.
- Converter ou carregar modelo para encomenda.
- Criar encomenda interna.
- Criar pecas.
- Definir tempos por operacao/posto.
- Gerar OF.
- Gerar PDF de OF.
- Ler codigo OF no operador.
- Ler codigo de espessura/grupo.
- Produzir/concluir operacao.
- Gerar planeamento.
- Gerar PDF de planeamento.
- Emitir documentos de expedicao/faturacao quando aplicavel.

### Tecnico

- `python scripts/verify_core_flows.py`
- `python scripts/security_audit.py`
- `python scripts/verify_mysql_schema.py`
- `python scripts/verify_mysql_persistence_smoke.py`
- Backup e restore testados.
- Build PyInstaller testado numa maquina limpa.
- Instalacao servidor/posto testada.

### Seguranca

- `lugest.env` sem segredos de desenvolvimento.
- MySQL sem utilizador `root` na app.
- Passwords fortes.
- Admin temporario trocado.
- Pasta partilhada com permissoes corretas.
- Backups protegidos.

## Plano de Execucao

### Etapa 1: fechar produto base

- Limpar menus onde campos globais confundem a producao.
- Consolidar Extras: operacoes e postos.
- Validar PDFs.
- Melhorar operador com leitura OF/grupo.
- Validar planeamento.

### Etapa 2: fechar instalacao

- Rever `prepare_final_release.ps1`.
- Garantir que pacote final nao leva segredos reais.
- Criar checklist de instalacao por cliente.
- Testar numa VM limpa.

### Etapa 3: fechar base de dados

- Gerar schema consolidado.
- Aplicar baseline `schema_migrations`.
- Criar migracoes pequenas para tabelas industriais novas.
- Testar backup/update/restore.

### Etapa 4: fechar atualizacoes

- Versionar releases.
- Criar notas de versao.
- Atualizar cliente sempre com backup previo.
- Guardar historico de schema aplicado por cliente.

## Regra de Ouro

Qualquer mudanca em cliente real tem de cumprir:

`backup -> migracao -> validacao -> teste funcional -> entrega`

Se um passo falhar, para-se e restaura-se. O objetivo e vender um sistema que
parece simples ao utilizador, mas que por baixo e previsivel, recuperavel e
facil de manter.
