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
contadores AUTO_INCREMENT. Criacao/atribuicao de viagens, rececao de qualidade,
emissao de faturas e outras areas continuam em migracao; consultar a lista no
guia do monolito modular.
