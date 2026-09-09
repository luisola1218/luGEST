# Documentos de qualidade

CRUD e consultas passaram para `quality/application/documents.py`, com
repositorio de catalogo e auditoria partilhado apenas dentro do modulo.
O adaptador historico de NC reutiliza esse repositorio, preservando a sua API.
As operacoes deixam de publicar alteracoes durante validacao e passam a
aguardar a gravacao; erros imediatos recuperam registos e auditoria.

Validacao: 43 verificacoes locais aprovadas; 395 ficheiros Python compilados;
487 contratos do backend preservados. O fluxo SQL de documentos incluiu
ficheiro, criacao, edicao, remocao e releitura (7.922 s). Apos rollback, o
conteudo das 55 tabelas ficou igual. Os ficheiros desse teste ficaram
isolados no armazenamento temporario do executor.

A fronteira ainda historica e os limites da integracao com ficheiros estao
registados em `lugest_modules/quality/README.md`. Rececao e libertacao de stock
nao foram migradas nesta alteracao.
