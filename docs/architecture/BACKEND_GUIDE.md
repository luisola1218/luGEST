# Guia de manutencao do backend

Este guia e o ponto de entrada para alterar o Lugest sem depender de IA ou de
conhecimento previo do antigo `main_bridge.py`.

Os componentes migrados por negocio vivem agora em `lugest_modules/`.
Consultar primeiro o [guia do monolito modular](MODULAR_MONOLITH.md), que identifica
as implementacoes de clientes, editores, regras, nesting e PDFs de orcamentos,
os contratos e as partes que ainda dependem do runtime historico.

## Comecar uma alteracao

1. Encontrar a pagina em `lugest_qt/ui/pages/` e a chamada `self.backend.metodo`.
2. Localizar a implementacao e as dependencias com:

   ```powershell
   .venv\Scripts\python.exe scripts\backend_map.py order_create_or_update
   .venv\Scripts\python.exe scripts\backend_map.py --area materials
   ```

3. Consultar o [indice completo de metodos](BACKEND_METHOD_INDEX.md), que mostra
   a implementacao efetiva, incluindo os casos de sobreposicao por heranca.
4. Alterar a regra no dominio, o acesso a dados no repositorio ou a coordenacao
   no adaptador da respetiva area. Evitar colocar uma regra nova no widget.
5. Executar os testes da responsabilidade alterada e a porta `-SafeOnly`.
   Fluxos com escrita requerem staging ou o executor transacional de rollback,
   quando houver autorizacao para testar na base atual.

## Camadas e composicao

```text
Pagina Qt
  -> LegacyBackend (API de compatibilidade)
      -> adaptador da area em services/bridge_mixins/
          -> lugest_core: regras independentes de UI e armazenamento
          -> lugest_infra: acesso a armazenamento e servicos externos
          -> LegacyRuntime: integracao historica ainda em migracao
```

`services/legacy_backend.py` e a entrada publica. `services/main_bridge.py`
contem a composicao das areas e a inicializacao do estado; nao recebe regras
de negocio novas. As 32 areas extraidas mantem a precedencia que os metodos
tinham antes sobre os nove adaptadores ja existentes.

Os mixins sao adaptadores de compatibilidade, nao servicos autonomos. Partilham
o estado do backend e alguns chamam outras areas. A separacao permite localizar
o codigo, mas nao elimina por si so esse acoplamento. `backend_map.py metodo`
mostra os metodos chamados e os atributos do estado utilizados.

Para codigo novo, preferir um servico com argumentos explicitos e testes
isolados, chamado por um metodo curto no adaptador. Exemplos implementados:

- Clientes: `lugest_core/clients.py` contem casos de uso e o contrato
  `ClientRepository`. `lugest_infra/legacy/clients_repository.py` adapta a
  persistencia antiga atraves de callbacks explicitos; `bridge_mixins/clients.py`
  apenas compoe estas dependencias. O teste `verify_clients_integration.py`
  percorre a pagina Qt, o backend, o servico e o repositorio sem MySQL.
- `lugest_core/operation_costing.py`: estimativa de custos com configuracao e
  funcoes de normalizacao explicitas; nao le MySQL nem configuracao local.
- `lugest_core/snapshots.py`: comparacao e combinacao de snapshots; recebe os
  dados local, original e remoto e devolve o resultado sem alterar as entradas.
- `lugest_infra/config/repository.py`: persistencia com um `ConfigurationStore`
  e uma fabrica de ligacoes SQL injetados. O adaptador Qt gere apenas cache e
  invalidacao de configuracao.

## Onde alterar cada area

Os caminhos desta tabela sao relativos a `lugest_qt/services/bridge_mixins/`.

| Funcionalidade | Modulos principais |
| --- | --- |
| Clientes | `clients.py` |
| Materiais, geometrias, consumos e precos | `materials.py` |
| Produtos e movimentos | `products.py` |
| Taxonomia, sugestoes e aprendizagem de produtos | `product_catalog.py` |
| PDFs e etiquetas de materiais/produtos | `material_reports.py`, `product_reports.py` |
| Codigos de leitura de stock | `inventory_scanning.py` |
| Encomendas, pecas, reservas e modelos | `orders.py` |
| PDF de ordem de fabrico | `order_reports.py` |
| Consulta de OPP e carteira de producao | `production_orders.py` |
| Operador: inicio, pausa, conclusao e avarias | `operator.py` |
| Etiquetas de operador e paletes | `operator_labels.py` |
| Montagem: necessidades, compras e consumo | `assembly_stock.py` |
| Orcamentos e conjuntos | `quotes.py` |
| Referencias de cliente e sincronizacao orcamento/peca | `quote_references.py` |
| Adaptacao dos calculos laser e custos de operacoes | `operation_costing.py` |
| Catalogo de operacoes | `operation_catalog.py` |
| Planeamento e sequenciamento | `planning.py`, `planning_operations.py` |
| Justificacao de atrasos | `planning_delays.py` |
| Centros de trabalho, maquinas e postos | `workcenters.py` |
| Assistente de materiais e listas de separacao | `material_assistant.py`, `material_assistant_reports.py` |
| Compras e documentos de rececao | `purchasing.py`, `purchasing_documents.py` |
| Expedicao, transportes e faturacao | `shipping.py`, `transport.py`, `billing.py` |
| Servicos diretos e dashboards | `direct_services.py`, `dashboards.py` |
| Qualidade: rececoes, NC, quarentena e documentos | `quality.py` |
| PDFs de qualidade | `quality_reports.py` |
| Utilizadores, permissoes e trial historico | `users.py` |
| Configuracao de interface | `configuration.py` |
| Marca, logotipo e tema de PDF | `branding.py` |
| Atualizacoes de software | `updates.py` |
| Copiloto e comandos de materiais | `intelligence.py` |
| Carregamento, snapshots, gravacao e auditoria | `data_runtime.py` |
| Documentos partilhados e caminhos de ficheiros | `documents.py` |
| Diagnosticos e consulta do historico de auditoria | `diagnostics.py` |
| Conversoes e utilitarios de compatibilidade | `common.py`, `../bridge_helpers.py` |

## Exemplos de alteracoes

**Mudar o custo de uma operacao:** alterar `OperationCostingEngine.estimate`
em `lugest_core/operation_costing.py`. As tarifas predefinidas vivem em
`default_settings`; as escolhas guardadas continuam a chegar pelo adaptador.
Validar com `scripts/verify_operation_costing.py`.

**Mudar onde se guarda configuracao:** alterar o repositorio em
`lugest_infra/config/repository.py` ou a resolucao de caminhos em
`lugest_infra/app_paths.py`. O repositorio prefere MySQL na leitura, usa o
ficheiro como fallback e considera uma escrita bem sucedida quando pelo menos
um destino foi guardado. Falha parcial fica em diagnostico; falha total levanta
erro. Validar com `verify_configuration_repository.py` e `verify_app_storage.py`.

**Mudar uma reserva de material numa encomenda:** procurar `order_reserve_stock`
com `backend_map.py`. A coordenacao esta em `orders.py`; seguir as dependencias
listadas antes de alterar a gravacao ou consumo partilhados.

**Mudar validacao ou pesquisa de clientes:** alterar `ClientService` em
`lugest_core/clients.py`. Para alterar a escrita, seguir o contrato do
repositorio. O adaptador fornece a normalizacao de pesquisa historica, mantendo
a compatibilidade com a interface. O caminho assincrono confirma aceitacao na
fila, nao durabilidade; erros posteriores continuam a pertencer ao worker.

**Mudar a combinacao de alteracoes entre postos:** usar `SnapshotMergePolicy`.
A politica atual preserva linhas remotas que o posto nao alterou; conflitos
na mesma linha identificada continuam a favorecer a linha local inteira.
Nao existe detecao de conflito por campo nem nova garantia transacional.

## Dependencias e estado historicos

`LegacyRuntime`, em `services/legacy_runtime.py`, declara os onze componentes
historicos recebidos pelo backend. O arranque normal chama `load_legacy_runtime`;
testes podem passar `LegacyBackend(runtime=...)`, sem importar `main.py`.

| Estado | Responsabilidade |
| --- | --- |
| `data`, `_base_data_snapshot` | dados em edicao e referencia da ultima leitura/gravacao |
| `_data_loaded_at`, `_data_cache_generation` | idade e geracao do snapshot |
| `user` | sessao autenticada |
| `_qt_config_cache`, `_qt_config_last_error` | cache independente e falhas de configuracao |
| `_operation_catalog_cache`, `_product_taxonomy_nodes_cache` | caches derivados, sujeitos a invalidacao |
| `desktop_main`, `*_actions` | dependencias historicas fornecidas no construtor |

Nao criar instancias isoladas dos mixins nem atualizar os caches diretamente
na UI. Usar os metodos publicos e o fluxo de invalidacao existente.

## Contratos e verificacao

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File scripts\verify_project.ps1 -SafeOnly
.venv\Scripts\python.exe scripts\backend_map.py --write-index
```

`scripts/contracts/backend_api.json` conserva as assinaturas dos 487 metodos
do backend original. `verify_backend_modules.py` valida a API, decoradores,
referencias globais dos callbacks e construcao com dependencias simuladas.
Impede tambem que regras voltem ao ficheiro de composicao. Uma alteracao
intencional de API requer rever os consumidores e o contrato explicitamente.

As verificacoes isoladas nao substituem testes integrados de stock, faturacao
e producao numa base de staging nem a validacao do executavel empacotado.

Quando houver autorizacao explicita para usar a base atual, o runner
`verify_database_rollback.py` permite testar um fluxo com rollback e comparacao
do conteudo das tabelas. Nao pertence a `-SafeOnly`: exige InnoDB, serializa
ligacoes, bloqueia DDL e nao valida durabilidade assincrona. Ver o
[relatorio desta fase](../plans/CLIENTS_PHASE_2026-09-07.md).

Metodos que chamam outros fluxos de gravacao devem voltar a obter os dados com
`ensure_data`/`get_encomenda_by_numero` antes de continuar a altera-los: a
gravacao pode substituir o snapshot. A regressao e coberta por
`verify_quote_snapshot_save.py`.

## Trabalho que continua a ser legacy

- Muitos adaptadores ainda combinam regras, coordenacao e chamadas ao runtime
  antigo. Extrair essas regras quando a area for alterada, com testes de dominio.
- A gravacao de datasets ainda passa pelo runtime antigo; nao foi redesenhado o
  modelo MySQL nem o mecanismo de gravacao assincrona.
- `module_context` continua a configurar os modulos antigos. A injecao de
  `LegacyRuntime` limita o ponto de entrada, mas nao remove os globais internos.
- A pagina de orcamentos continua grande; este trabalho incidiu no backend.
- O sistema comercial de licencas continua separado e sem bloqueio ativo.
