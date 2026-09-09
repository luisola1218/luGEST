# Transportes, pagamentos e qualidade

## Implementacao

- `transport/application/tariffs.py`: CRUD, selecao por transportadora/zona e
  calculo de custos. Corrigida a remocao que chamava `int(float, base)` e a
  comparacao de dicionarios em tarifarios com identificadores empatados.
- `transport/application/stops.py`: paragens, progressao da viagem, pedido
  externo e checklist/POD. A troca de posicao agora renumera as paragens antes
  de serem novamente ordenadas. Uma guia desconhecida e rejeitada mesmo quando
  a encomenda nao tem guias validas.
- `transport/application/order_links.py`: projecao unica das ligacoes entre
  viagens ativas e encomendas, usada tambem pela compatibilidade historica.
- `billing/application/payments.py`: criacao, edicao e remocao de pagamentos.
  Valores nao finitos sao rejeitados. A fatura herdada do pagamento existente
  tambem e validada, impedindo contornar a verificacao de anulacao ao omitir
  `fatura_id` na edicao.
- `quality/application/nonconformities.py`: ciclo de NC, consultas e tratamento
  de duplicados. O repositorio preserva a auditoria e recupera os campos locais
  perante falha imediata de gravacao ou de criacao do evento.

Os repositorios recebem capacidades explicitas. Os servicos nao recebem o
backend nem o snapshot global. As escritas preparadas verificam o estado local
esperado e aguardam a gravacao. Isto nao substitui controlo transacional entre
instalacoes nem elimina os restantes adaptadores historicos.

## Validacao

- 37 verificacoes locais aprovadas; 375 ficheiros Python compilados.
- Os 487 contratos do backend foram preservados.
- Na base atual, com rollback e conteudo igual nas 55 tabelas: tarifarios
  (8.481 s), paragens com gravacao/releitura (15.710 s), NC (8.020 s), compras
  (18.959 s) e faturacao (19.659 s).
- O cenario mais amplo de transportes/PDF passou (1.217 s). Esse script usa um
  substituto de `_save`; a durabilidade dos tarifarios e paragens foi verificada
  separadamente nos dois fluxos anteriores com gravacao real protegida.

Relatorios locais: `reports/integration/`. O rollback pode deixar intervalos nos
contadores AUTO_INCREMENT. Rececao de qualidade,
emissao de faturas e outras areas continuam em migracao; consultar a lista no
guia do monolito modular.


## Continuacao: viagens, atribuicoes e consultas

- `transport/application/trips.py`: criacao, edicao e remocao, com validacao
  anterior a publicacao no snapshot e a reserva do identificador.
- `transport/application/assignments.py`: validacao de todas as encomendas
  antes da atribuicao, detecao de outra viagem e aplicacao do custo sugerido.
- `transport/application/queries.py`: listas, detalhe, elegibilidade, metricas
  e opcoes sobre copias dos registos. Recalcular o estado de expedicao para
  apresentar uma encomenda nao altera a encomenda partilhada.
- `transport/infrastructure/route_report.py`: os dois formatos do documento de
  viagem recebem detalhe e capacidades explicitas; nao recebem o backend.
- A remocao de NC passou a usar o mesmo repositorio com auditoria e recuperacao
  das outras operacoes do ciclo de vida.

Validacao local: 40 verificacoes aprovadas, 383 ficheiros compilados e 487
contratos preservados. Novos testes cobrem dados invalidos sem reserva de
numero, atribuicoes parcialmente invalidas sem efeitos, conflitos locais,
falhas de escrita, consultas sem alteracao do snapshot e remocao de NC.
O fluxo SQL de viagens/paragens passou novamente (15.682 s), com releitura;
o fluxo de transportes/PDF passou (1.272 s, com substituto de `_save`).
Os dois verificaram igualdade do conteudo das 55 tabelas apos rollback.

As verificacoes de estado esperado protegem alteracoes locais. Nao sao um
mecanismo de concorrencia entre instalacoes. O adaptador de numeracao ainda
pode reservar contadores/configuracao independentemente da gravacao da viagem.
A migracao global continua com os pontos registados em MODULAR_MONOLITH.md.
