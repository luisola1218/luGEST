# Monolito modular: estado e regras de manutencao

O Lugest continua a ser uma aplicacao e uma instalacao. A unidade de organizacao
passa a ser o modulo de negocio em `lugest_modules/`, com as suas regras,
casos de uso, adaptadores de dados e interface proximos uns dos outros.

**A migracao ainda nao esta concluida.** Estes limites ja se aplicam aos novos
componentes de clientes e orcamentos. O backend historico ainda contem estado
partilhado e outras areas continuam nos adaptadores antigos. Nao confundir a
existencia desta estrutura com a migracao de todo o ERP.

## Onde alterar uma funcionalidade

| Funcionalidade | Implementacao |
| --- | --- |
| Validar, pesquisar, guardar e remover clientes | `lugest_modules/clients/application/service.py` |
| Persistir clientes no runtime existente | `lugest_modules/clients/infrastructure/legacy_repository.py` |
| Formulario de clientes | `lugest_modules/clients/presentation/page.py` |
| Criar linhas comerciais de produto e servico | `lugest_modules/quotes/domain/lines.py` |
| Normalizar uma linha completa de orcamento | `lugest_modules/quotes/application/line_normalization.py` |
| Guardar e combinar estudos de nesting | `lugest_modules/quotes/application/nesting_studies.py` |
| Atualizar o snapshot de um estudo | `lugest_modules/quotes/infrastructure/legacy_nesting_repository.py` |
| SQL dos estudos de nesting | `lugest_modules/quotes/infrastructure/mysql_nesting_store.py` |
| PDF de conjunto e de nesting | `lugest_modules/quotes/infrastructure/assembly_report.py`, `nesting_report.py` |
| Editores de MP, produto, mao de obra, consumiveis, estrutura, linha e STEP/IGS | `lugest_modules/quotes/presentation/*_editor.py` |
| Consulta e atualizacao de precos nos editores | `lugest_modules/quotes/presentation/material_prices.py` |
| Ligar estes componentes ao runtime existente | `lugest_qt/services/quote_*_composition.py` |
| Lista, selecao, estado e coordenacao da pagina de orcamentos | `lugest_qt/ui/pages/quotes_page.py` |

Cada modulo expoe `api.py` para outros modulos de negocio. O ponto de composicao
da aplicacao pode importar adaptadores concretos para construir os servicos.
Os caminhos antigos de clientes sao imports de compatibilidade; novas regras
nao devem voltar a ser implementadas nesses ficheiros.

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

## Persistencia dos estudos

O servico valida o orcamento e o grupo, preserva a data de criacao e recebe o
relogio explicitamente. O repositorio devolve copias e restaura os campos do
cache quando a gravacao principal falha imediatamente. Isto corrige o caso em
que o estudo parecia guardado em memoria apesar de a gravacao ter falhado.

Mantem-se o comportamento historico de um snapshot principal e um espelho SQL
opcional. A falha do espelho nao desfaz uma gravacao principal ja aceite. Nao e
uma transacao atomica entre ambos. A aceitaçao por um worker assincrono tambem
nao garante durabilidade; estes aspetos continuam a exigir evolucao propria.

## Verificar uma alteracao

```powershell
.venv\Scripts\python.exe scripts\verify_business_modules.py
.venv\Scripts\python.exe scripts\verify_clients_integration.py
.venv\Scripts\python.exe scripts\verify_quote_editors.py
.venv\Scripts\python.exe scripts\verify_nesting_study_service.py
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
```

O executor confirma InnoDB, bloqueia commits reais/DDL e compara o conteudo das
55 tabelas depois do rollback. Pode haver intervalos nas sequencias de
AUTO_INCREMENT. Nao testa concorrencia real nem persistencia assincrona.

## Trabalho ainda necessario para concluir a migracao

1. Retirar o estado de edicao e a construcao dos paineis do controlador de
   orcamentos. Ainda tem 5350 linhas, incluindo um construtor extenso.
2. Migrar os comandos de orcamento, conversao para encomenda e conjuntos para
   servicos com repositorios. O adaptador de orcamentos ainda tem 1898 linhas.
3. Migrar encomendas, inventario, compras, producao, faturacao e as restantes
   areas. Ainda partilham estado atraves dos mixins de `LegacyBackend`.
4. Substituir os adaptadores do snapshot global por repositorios por modulo
   e definir as transacoes que envolvem mais de um modulo.
5. Retirar a inicializacao global historica de `main.py` e validar concorrencia,
   gravacoes assincronas, empacotamento e instalacao num Windows limpo.

Os testes aprovados demonstram os percursos verificados. Nao demonstram que
esta lista esta concluida nem que o ERP esteja pronto para comercializacao.
