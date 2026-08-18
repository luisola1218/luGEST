# Arquitetura da aplicação móvel de serviços

## Decisão

A aplicação chama-se **LuGEST Field** e vive em `impulse_mobile_app/`. É uma app
Flutter separada do desktop Qt, otimizada para uso com uma mão, em deslocação e
com rede intermitente.

O produto é transversal a trabalhadores independentes: eletricistas,
canalizadores, manutenção, montagem, assistência, limpeza e outras atividades.
O desktop continua responsável por configuração, faturação fiscal, SAF-T(PT),
gestão de catálogos e controlo global. A app fica responsável pela execução no
terreno: agenda, chegada, diagnóstico, tempo, materiais, fotografias, assinatura,
conta apresentada ao cliente e conclusão.

## Fluxo funcional

1. O serviço é criado no telemóvel ou atribuído no desktop.
2. O técnico inicia a deslocação e confirma a chegada.
3. Regista diagnóstico, trabalho, materiais, fotografias e assinatura.
4. Apresenta ao cliente um resumo não fiscal com base, IVA e total.
5. O cliente valida e assina; o serviço passa a `Por faturar`.
6. A API cria ou atualiza o registo em `servicos_diretos` como rascunho; a baixa
   auditada do stock só ocorre quando o utilizador confirma no desktop.
7. O desktop emite a fatura no menu Faturação e devolve o estado à app.

## Estado implementado em 0.3.0

- câmara e galeria funcionais, com recuperação após interrupção do Android;
- assinatura manuscrita guardada em PNG dentro da área privada da app;
- linhas manuais de trabalho, horas e material;
- conta final com taxas de IVA 0%, 6%, 13% e 23%;
- gateway local autenticado em `mobile_gateway/`;
- consulta real, pesquisa e seleção de produtos e matérias-primas do MySQL;
- sincronização dos serviços para o desktop como rascunhos protegidos;
- envio idempotente das fotografias e assinaturas para arquivo isolado;
- fila offline persistente: um trabalho só fica sincronizado após todos os anexos;
- chave da API protegida pelo Android KeyStore na app;
- HTTPS obrigatório no Android e Cloudflare Tunnel ligado apenas a `127.0.0.1`;
- chaves individuais revogáveis, armazenadas apenas por hash no computador;
- limite de pedidos e proteção contra alteração de documentos fechados.

O gateway 0.3.0 é deliberadamente só de leitura para o stock. A baixa automática
só deve ser ativada juntamente com autenticação por utilizador, idempotência,
transação e confirmação final do serviço; isto evita movimentos duplicados se a
rede cair durante a sincronização.

## Segurança e sincronização

- a porta do gateway não é exposta ao router nem à rede local no modo remoto;
- HTTPS obrigatório e uma chave revogável por dispositivo;
- documentos `Confirmado`, `Faturado` e `Anulado` nunca são sobrescritos;
- anexos são identificados pelo SHA-256, validados e limitados a 10 MB;
- a fila local é reprocessada e mantém o erro visível ao técnico.

## API mínima

```text
POST   /v1/auth/session
GET    /v1/field/dashboard
GET    /v1/service-jobs?from=&to=&status=&cursor=
POST   /v1/service-jobs
GET    /v1/service-jobs/{id}
PATCH  /v1/service-jobs/{id}
POST   /v1/service-jobs/{id}/events
POST   /v1/service-jobs/attachments
POST   /v1/service-jobs/{id}/complete
POST   /v1/service-jobs/{id}/send-to-billing
GET    /v1/clients?query=
GET    /v1/catalog?query=&kind=
GET    /v1/stock/summary
GET    /v1/stock/products?query=
GET    /v1/stock/materials?query=
POST   /v1/service-jobs/{id}/consume-stock
POST   /v1/sync/batch
```

## Evolução da base de dados

O modelo atual guarda linhas do serviço em `linhas_json`. É aceitável para o
protótipo e mantém compatibilidade com o desktop, mas a API de produção deve
normalizar os dados em tabelas de linhas, eventos e anexos. Isso permite pesquisa,
auditoria, sincronização incremental e relatórios sem analisar JSON em cada
consulta.

Campos recomendados para a próxima migração:

- `servicos_diretos`: `company_id`, `version`, `assigned_user_id`,
  `sync_updated_at`, `deleted_at`;
- `servicos_diretos_linhas`: quantidade, unidade, preço, IVA, produto/material;
- `servicos_diretos_eventos`: tipo, instante, utilizador, localização e payload;
- `servicos_diretos_anexos`: tipo, caminho, hash, tamanho e autor;
- `mobile_operations`: `operation_id`, dispositivo, estado e resultado.

## Compatibilidade com o desktop

O executável e o processo do LuGEST desktop não são usados pelo túnel. O gateway
é um processo separado, o túnel só encaminha para a interface local e os anexos
ficam em `mobile_gateway_data/attachments`. Apenas rascunhos móveis podem ser
atualizados; confirmação, consumo de stock e faturação continuam sob controlo do
desktop.

## Fases

1. Protótipo Flutter local, anexos, assinatura, conta e leitura real de stock.
2. API autenticada por utilizador e leitura real de agenda/clientes/catálogo.
3. Escrita offline, upload de anexos e sincronização idempotente. **Concluído.**
4. Piloto controlado, backups e endurecimento de permissões MySQL.
5. Publicação Android/iOS e operação assistida.
