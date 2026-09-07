# Auditoria de arquitetura e prontidao comercial — 2026-09-04

Versao: `2026.09.04.1`

## Resultado desta passagem

- A porta segura compilou 214 ficheiros Python e passou 13 verificacoes
  automaticas, sem aceder em escrita a base MySQL remota configurada.
- `pip check` nao encontrou dependencias quebradas.
- A auditoria de seguranca terminou com 0 ocorrencias altas, 0 medias e 1 baixa:
  o ficheiro local `lugest.env`, necessario ao posto, tem de continuar fora de
  pacotes comerciais e protegido nos backups.
- As fronteiras de dependencia passaram: 20 ficheiros/8054 linhas em
  `lugest_core` e 21 ficheiros/3293 linhas em `lugest_infra`.
- Calculo e nesting laser, schema, migracoes, licenciamento, diagnostico,
  controlos de orcamentos e layout de parceiros passaram os testes seguros.

## Alteracoes estruturais

1. `AppPaths` centraliza dados mutaveis:
   configuracao/logs/cache em `%LOCALAPPDATA%\luGEST-data` e licenca em
   `%PROGRAMDATA%\luGEST\licensing`.
2. Configuracao antiga e copiada para o novo local no primeiro arranque, sem
   apagar a origem. JSON local passa a ter escrita atomica, limite de tamanho e
   copia de recuperacao.
3. Ler configuracao deixou de tentar criar tabelas MySQL. Guardar deixa de
   ocultar uma falha total dos destinos.
4. Instalacoes novas por utilizador usam `%LOCALAPPDATA%\Programs\luGEST`,
   separando binarios de dados.
5. O formato de manifesto de atualizacao passou para `lugest_core/updates`.
   HTTPS, hash do ZIP e hash do reparador sao obrigatorios.
6. O bootstrap e descarregado para staging, validado e substituido de forma
   atomica. Ficheiros temporarios sao removidos.
7. O gerador `scripts/create_update_manifest.py` calcula os hashes; o pacote
   comercial inclui obrigatoriamente os quatro scripts `.ps1/.bat` do updater.
8. Tokens GitHub deixam de ser escritos no novo `update_config.json`.

## Estado comercial honesto

O codigo esta preparado para demonstracao e piloto acompanhado. “Pronto para
comercializar” em venda geral exige ainda evidencias que nao podem ser criadas
apenas dentro deste repositorio:

- certificado de code signing e assinatura dos binarios/instalador;
- endpoint HTTPS de releases e, idealmente, assinatura criptografica do
  manifesto para garantir a origem, nao apenas a integridade;
- teste de instalacao, atualizacao, rollback e concorrencia num Windows limpo e
  numa base MySQL de staging representativa;
- decisao das edicoes, modulos, validade, tolerancia offline e processo de
  emissao/revogacao das licencas antes de ligar o bloqueio ao arranque;
- contrato, suporte, RGPD, backups e recuperacao de desastre;
- certificacao AT antes de vender o modulo de faturacao como certificado.

## Divida tecnica que permanece

`lugest_qt/services/main_bridge.py` e `lugest_qt/ui/pages/runtime_pages.py`
continuam demasiado grandes. Uma separacao integral num unico passo teria risco
desnecessario sobre fluxos produtivos. A proxima extracao deve ser incremental:

1. mover o adaptador de atualizacoes para um mixin dedicado;
2. separar as paginas de orcamentos e encomendas de `runtime_pages.py`;
3. substituir `module_context.py` por dependencias explicitas por modulo;
4. cobrir cada extracao com teste de contrato antes de remover o shim legacy.

## Comando de regressao seguro

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\verify_project.ps1 -SafeOnly
```

Os testes funcionais que escrevem dados so devem ser executados depois de
configurar uma base de staging; nunca contra a base remota do cliente.
