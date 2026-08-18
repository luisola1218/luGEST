# Auditoria técnica do sistema — 2026-08-18

## Verificações executadas

- compilação de 185 ficheiros Python: sem erros;
- contrato visual da Faturação: 4 separadores, 10 colunas e próxima ação;
- seletor de data clicável: popup nativo abriu corretamente pelo corpo do campo;
- schema MySQL mínimo: válido;
- auditoria de integridade MySQL em modo só de leitura: 13 verificações, zero
  problemas encontrados;
- desempenho Qt: carga inicial `0,194 s`, página fria máxima `0,431 s` e
  navegação quente máxima `0,054 s`;
- auditoria de segurança: zero ocorrências altas ou médias e duas baixas.

## Estado geral

O desktop está funcional, com desempenho saudável e sem evidência de corrupção
relacional nos fluxos auditados. A separação parcial entre UI, bridge, domínio e
infraestrutura é adequada para evolução incremental.

## Dívida técnica relevante

1. `main.py`, `runtime_pages.py` e partes da bridge continuam demasiado grandes.
2. O runtime ainda faz evolução automática do schema; deve migrar para versões
   explícitas e reversíveis antes de expor uma API pública.
3. O modelo de serviços usa `linhas_json`; deve ser normalizado para mobile e
   reporting avançado.
4. Existem 52 tabelas fora de `utf8mb4`.
5. O ficheiro local de ambiente contém segredos e deve permanecer fora de Git,
   com permissões e backups protegidos.
6. Existem artefactos regeneráveis (`build` e cache Python) no posto; devem ser
   limpos apenas na preparação da entrega, não durante o desenvolvimento.

## Limites desta execução

Os fluxos funcionais que criam e removem dados não foram lançados contra a base
remota de produção. Foram executados apenas testes de compilação, UI, desempenho,
segurança e consultas de integridade sem alterações aos dados.
