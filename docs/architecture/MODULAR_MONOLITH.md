# Monolito modular: estado e regras de manutencao

O Lugest continua a ser uma aplicacao e uma instalacao. A unidade de organizacao
passa a ser o modulo de negocio em `lugest_modules/`, com as suas regras,
casos de uso, adaptadores de dados e interface proximos uns dos outros.

**A migracao ainda nao esta concluida.** Estes limites ja se aplicam aos novos
componentes de clientes, orcamentos, inventario, transportes, faturacao e qualidade.
O backend historico ainda contem estado
partilhado e outras areas continuam nos adaptadores antigos. Nao confundir a
existencia desta estrutura com a migracao de todo o ERP.

## Onde alterar uma funcionalidade

| Funcionalidade | Implementacao |
| --- | --- |
| Validar, pesquisar, guardar e remover clientes | `lugest_modules/clients/application/service.py` |
| Persistir clientes no runtime existente | `lugest_modules/clients/infrastructure/legacy_repository.py` |
| Formulario de clientes | `lugest_modules/clients/presentation/page.py` |
| Consultas de produtos, movimentos e valores de stock | `lugest_modules/inventory/application/product_queries.py` |
| Validacao de baixas e entregas a operadores | `lugest_modules/inventory/application/consumption.py` |
| Criacao, edicao e remocao de produtos | `lugest_modules/inventory/application/product_commands.py` |
| Normalizacao e previsao de preco de produto | `lugest_modules/inventory/application/product_definition.py` |
| Leitura de produtos e gravacao das baixas no runtime | `lugest_modules/inventory/infrastructure/` |
| Criar linhas comerciais de produto e servico | `lugest_modules/quotes/domain/lines.py` |
| Normalizar uma linha completa de orcamento | `lugest_modules/quotes/application/line_normalization.py` |
| Guardar e combinar estudos de nesting | `lugest_modules/quotes/application/nesting_studies.py` |
| Guardar/remover orcamento e alterar estado | `lugest_modules/quotes/application/commands.py` |
| Consultar lista, anos e detalhe de orcamentos | `lugest_modules/quotes/application/queries.py` |
| Guardar/remover conjuntos e modelos | `lugest_modules/quotes/application/assembly_catalog.py` |
| Listas e detalhes dos catalogos de conjuntos | `lugest_modules/quotes/application/assembly_queries.py` |
| Normalizacao, precos e expansao dos componentes | `lugest_modules/quotes/application/assemblies.py` |
| Atualizacao preparada de precos e codigos de parametro | `lugest_modules/quotes/application/assembly_refresh.py` |
| Persistencia dos catalogos e recuperacao perante erro | `lugest_modules/quotes/infrastructure/legacy_assembly_repository.py` |
| Preparar pecas, montagem e tempos para encomenda | `lugest_modules/quotes/application/order_lines.py` |
| Coordenar conversao de orcamento aprovado | `lugest_modules/quotes/application/conversion.py` |
| Preparacao isolada e gravacao da conversao | `lugest_modules/quotes/infrastructure/legacy_conversion_repository.py` |
| Necessidades de compra com stock partilhado entre linhas | `lugest_modules/quotes/application/purchase_needs.py` |
| Atualizar o snapshot de um estudo | `lugest_modules/quotes/infrastructure/legacy_nesting_repository.py` |
| SQL dos estudos de nesting | `lugest_modules/quotes/infrastructure/mysql_nesting_store.py` |
| PDF de conjunto e de nesting | `lugest_modules/quotes/infrastructure/assembly_report.py`, `nesting_report.py` |
| Editores de MP, produto, mao de obra, consumiveis, estrutura, linha e STEP/IGS | `lugest_modules/quotes/presentation/*_editor.py` |
| Consulta e atualizacao de precos nos editores | `lugest_modules/quotes/presentation/material_prices.py` |
| Ligar estes componentes ao runtime existente | `lugest_qt/services/quote_*_composition.py` |
| Lista, selecao, estado e coordenacao da pagina de orcamentos | `lugest_modules/quotes/presentation/page.py` |
| Construcao de paineis e encaminhamento de eventos | `lugest_modules/quotes/presentation/workspace.py` |
| Aparencia dos paineis | `lugest_modules/quotes/presentation/workspace_styles.py` |
| Dependencias do controlador de orcamentos | `lugest_modules/quotes/presentation/page_services.py` |
| CRUD, selecao e calculo de tarifarios | `lugest_modules/transport/application/tariffs.py` |
| Paragens, pedido externo e progressao da viagem | `lugest_modules/transport/application/stops.py` |
| Ligacao das encomendas as viagens ativas | `lugest_modules/transport/application/order_links.py` |
| Pagamentos de faturacao | `lugest_modules/billing/application/payments.py` |
| Criar, fechar, consultar e tratar duplicados de NC | `lugest_modules/quality/application/nonconformities.py` |
| Construtor visual de conjuntos calculados | `lugest_modules/quotes/presentation/calculated_assembly_editor.py` |

Cada modulo expoe `api.py` para outros modulos de negocio. O ponto de composicao
da aplicacao pode importar adaptadores concretos para construir os servicos.
Os caminhos antigos de clientes sao imports de compatibilidade; novas regras
nao devem voltar a ser implementadas nesses ficheiros.

`scripts/backend_map.py nome_do_metodo` mostra o adaptador e segue as funcoes
e fabricas tipadas ate ao caso de uso no modulo. O indice gerado contem a mesma
ligacao na coluna "Modulo de negocio".

## Dependencias permitidas

- `domain`: valores e regras de negocio; sem Qt, MySQL, ficheiros ou runtime.
- `application`: casos de uso, contratos de repositorio e capacidades recebidas
  no construtor; sem importar infraestrutura ou apresentacao.
- `infrastructure`: implementacoes dos contratos e integracoes externas.
- `presentation`: widgets e eventos; recebe operacoes explicitas e devolve
  resultados. Nao recebe o snapshot global nem importa `main`.
- A composicao liga os adaptadores antigos aos contratos. Os ports congelados
  enumeram dependencias, mas os callbacks ainda podem executar codigo antigo:
  retirar o import direto nao significa que essa implementacao ja foi reescrita.

`verify_business_modules.py` verifica imports, acesso ao runtime e referencias
globais dentro dos callbacks. Imports entre modulos de negocio devem passar
pelo `api.py` do destino. A infraestrutura partilhada existente permanece em
`lugest_infra`; os componentes visuais partilhados permanecem em `lugest_qt/ui`.

## Contratos de interface

`ClientPage` recebe quatro acoes, sem conhecer `LegacyBackend`. Os editores de
orcamento recebem dados iniciais e capacidades declaradas em `QuoteEditorPorts`
ou `ProfileEditorPorts`. O `owner` so serve para parentesco de widgets Qt.
Nao e usado para obter campos ou metodos da pagina.

O editor STEP/IGS devolve linhas. A pagina decide quando as acrescentar e quando
atualizar a tabela. Cancelar o editor nao altera as linhas da pagina.

Os contratos ainda usam alguns nomes historicos e callbacks genericos. Tornar
essas interfaces mais pequenas e tipadas faz parte da migracao restante.

O controlador `QuotePage` recebe `QuotePageServices`, sem `LegacyBackend`,
`desktop_main` ou acesso ao snapshot. A composicao fornece tambem as fabricas
dos dialogos laser historicos. Os campos visuais estao em `page.view`; o estado
de edicao continua no controlador. O ponto antigo em `lugest_qt/ui/pages` e
apenas o construtor de compatibilidade. Um teste constroi e utiliza o controlador
sem importar `main.py` nem criar o backend. A pesquisa tem um temporizador de
180 ms explicitamente criado e testado; antes o callback referia um temporizador
inexistente.

As consultas da carteira pedem resumos com os campos visiveis e a contagem das
linhas. Nao copiam desenhos, estudos de nesting ou todas as linhas de cada
orcamento apenas para apresentar a lista. O detalhe continua a devolver uma
copia do agregado.

## Persistencia dos estudos

O servico valida o orcamento e o grupo, preserva a data de criacao e recebe o
relogio explicitamente. O repositorio devolve copias e restaura os campos do
cache quando a gravacao principal falha imediatamente. Isto corrige o caso em
que o estudo parecia guardado em memoria apesar de a gravacao ter falhado.

Mantem-se o comportamento historico de um snapshot principal e um espelho SQL
opcional. A falha do espelho nao desfaz uma gravacao principal ja aceite. Nao e
uma transacao atomica entre ambos. A aceitaçao por um worker assincrono tambem
nao garante durabilidade; estes aspetos continuam a exigir evolucao propria.

## Inventario

As consultas de produtos recebem um repositorio de leitura que devolve copias,
regras de valorizacao/catalogo/qualidade e um relogio para o ano de recurso.
Nao recebem `LegacyBackend` nem o snapshot completo. O detalhe devolvido nao
permite alterar metadados aninhados no cache por acidente.

O servico de baixas valida quantidade, existencia, disponibilidade numerica e
operador antes de pedir qualquer escrita. Corrige a baixa em memoria que ocorria
antes do erro de operador em falta. Quantidades NaN e infinitas sao rejeitadas.
O repositorio verifica a quantidade esperada e restaura os campos e movimentos
se a gravacao falhar imediatamente. Esta verificacao protege a janela local
entre leitura e escrita; nao substitui controlo de concorrencia transacional
entre instalacoes. Criacao, edicao e remocao tambem passam agora por casos de uso
e repositorios do modulo. A reposicao perante erro inclui produtos, movimentos
e sequencia. A normalizacao acontece antes de obter o snapshot de escrita,
evitando guardar num snapshot que uma dependencia ja substituiu. A classificacao
de catalogo ainda usa callbacks antigos. Os precos dos conjuntos sao calculados
em `quotes/application/assemblies.py`, com capacidades explicitas para consultar
produtos, materiais e linhas de orcamentos. `AssemblyRefresh` prepara o catalogo
completo antes de escrever; um item invalido nao altera os modelos anteriores.
O adaptador restaura o catalogo perante falha imediata de gravacao e rejeita
substituicoes de um snapshot local entretanto alterado. A expansao de modelos
e conjuntos usa a mesma funcao e devolve linhas independentes do original.

## Verificar uma alteracao

```powershell
.venv\Scripts\python.exe scripts\verify_business_modules.py
.venv\Scripts\python.exe scripts\verify_clients_integration.py
.venv\Scripts\python.exe scripts\verify_quote_editors.py
.venv\Scripts\python.exe scripts\verify_nesting_study_service.py
.venv\Scripts\python.exe scripts\verify_inventory_services.py
.venv\Scripts\python.exe scripts\verify_product_commands.py
.venv\Scripts\python.exe scripts\verify_quote_commands.py
.venv\Scripts\python.exe scripts\verify_quote_workspace.py
.venv\Scripts\python.exe scripts\verify_assembly_rules.py
.venv\Scripts\python.exe scripts\verify_quote_order_lines.py
.venv\Scripts\python.exe scripts\verify_quote_conversion.py
.venv\Scripts\python.exe scripts\verify_calculated_assembly_editor.py
.venv\Scripts\python.exe scripts\verify_transport_tariffs.py
.venv\Scripts\python.exe scripts\verify_transport_stops.py
.venv\Scripts\python.exe scripts\verify_transport_trips.py
.venv\Scripts\python.exe scripts\verify_transport_assignments.py
.venv\Scripts\python.exe scripts\verify_transport_queries.py
.venv\Scripts\python.exe scripts\verify_billing_payments.py
.venv\Scripts\python.exe scripts\verify_quality_nonconformities.py
.venv\Scripts\python.exe scripts\verify_purchase_needs.py
powershell -File scripts\verify_project.ps1 -SafeOnly
```

Os testes de editores constroem oito dialogos sem a pagina nem o runtime,
verificam cancelamento e cinco percursos de confirmacao. O teste de nesting
cobre validacao, substituicao do snapshot, falha de gravacao, copias,
precedencia por data e parametros SQL.

Para testes autorizados na base atual, usar o executor transacional, nunca um
script funcional com escrita direta por engano:

```powershell
.venv\Scripts\python.exe scripts\verify_database_rollback.py verify_quote_nesting_flow
.venv\Scripts\python.exe scripts\verify_database_rollback.py verify_conjuntos_montagem_flow
.venv\Scripts\python.exe scripts\verify_database_rollback.py verify_inventory_flow
```

O executor confirma InnoDB, bloqueia commits reais/DDL e compara o conteudo das
55 tabelas depois do rollback. Pode haver intervalos nas sequencias de
AUTO_INCREMENT. Nao testa concorrencia real nem persistencia assincrona.

## Trabalho ainda necessario para concluir a migracao

1. Continuar a dividir o controlador de orcamentos: os paineis ja sao uma vista
   independente, assim como o construtor de conjuntos calculados. O controlador
   ainda concentra os gestores de conjuntos, agrupamento de linhas e outras
   interacoes extensas. Guardar modelo e conjunto no construtor ainda sao duas
   operacoes; falta definir a transacao conjunta deste percurso.
2. Migrar a criacao do pedido de compra e substituir os adaptadores historicos
   da conversao. QuoteConversion ja coordena cliente, encomenda e orcamento com
   repositorio explicito: prepara em copia, verifica alteracoes locais e pede
   gravacao bloqueante. A transacao SQL entre modulos continua por resolver;
   contadores podem ser reservados independentemente. Preparacao de linhas e
   calculo de faltas ja sao casos de uso separados. CRUD,
   consultas, precos, normalizacao, expansao e atualizacao do catalogo de
   conjuntos ja usam casos de uso; a alocacao de codigos e procura da linha de
   origem ainda ligam ao runtime antigo pela composicao. Os comandos de
   gravacao, remocao, estado e consulta de orcamentos ja migraram.
3. Migrar encomendas, restantes comandos de inventario, compras, producao,
   restantes operacoes de faturacao e as restantes areas. Pagamentos ja usam
   agregados isolados; a emissao de faturas ainda pertence ao adaptador antigo.
   Transportes ja tem servicos para tarifarios, viagens, atribuicoes, paragens
   e consultas, com renderizacao PDF separada. Os adaptadores de persistencia,
   numeracao e calculos de expedicao ainda dependem de capacidades historicas.
   NC ja usam repositorio com auditoria, mas rececao e libertacao de stock de
   qualidade ainda partilham estado atraves dos mixins de `LegacyBackend`.
4. Substituir os adaptadores do snapshot global por repositorios por modulo
   e definir as transacoes que envolvem mais de um modulo.
5. Retirar a inicializacao global historica de `main.py` e validar concorrencia,
   gravacoes assincronas, empacotamento e instalacao num Windows limpo.

Os testes aprovados demonstram os percursos verificados. Nao demonstram que
esta lista esta concluida nem que o ERP esteja pronto para comercializacao.
