# Transportes

## Onde alterar

| Comportamento | Implementacao |
| --- | --- |
| Criar, editar ou remover viagem | `application/trips.py` |
| Atribuir encomendas e aplicar custos | `application/assignments.py` |
| Editar e ordenar paragens, estados, POD | `application/stops.py` |
| Tarifarios, prioridade e sugestao de custo | `application/tariffs.py` |
| Listas, detalhe, filtros e elegibilidade | `application/queries.py` |
| Ligacoes de viagens a encomendas | `application/order_links.py` |
| Documento PDF de viagem | `infrastructure/route_report.py` |
| Leitura do armazenamento atual | `infrastructure/legacy_transport_read_repository.py` |
| Gravacao de viagens e ligacoes | `infrastructure/legacy_trip_repository.py` |
| Gravacao de tarifarios | `infrastructure/legacy_tariff_repository.py` |

`api.py` exporta os casos de uso. A composicao com o desktop fica em
`lugest_qt/services/transport_composition.py`. O mixin de transportes preserva
os nomes publicos usados pela interface. Para seguir uma chamada, executar
`python scripts/backend_map.py transport_assign_orders` na raiz do projeto.

## Regras para manutencao

- Casos de uso recebem repositorios e capacidades explicitas, nunca o backend
  ou o dicionario global da aplicacao.
- Leituras devolvem copias. Escritas preparam o resultado, verificam o estado
  esperado e aguardam a gravacao; nao publicar alteracoes durante validacao.
- A atualizacao dos vinculos de encomendas passa por `synchronize_orders`.
- O pedido de transporte regista o seu estado; nao envia mensagens externas.
- A recuperacao local perante erro nao substitui uma transacao SQL distribuida
  pelos restantes adaptadores. Numeracao e concorrencia entre instalacoes
  continuam dependentes da infraestrutura historica.

## Verificacao

Executar `powershell -File scripts/verify_project.ps1 -SafeOnly` na raiz.
Os cinco testes `verify_transport_*.py` de tarifarios, paragens, viagens,
atribuicoes e consultas nao precisam de acesso a MySQL.

Para integracao autorizada na base configurada, executar
`python scripts/verify_database_rollback.py verify_transport_stop_flow`.
Este percurso grava, relê e reverte os dados. O teste historico
`verify_transportes_module` tambem gera o PDF, mas substitui `_save`;
nao serve para comprovar durabilidade.
