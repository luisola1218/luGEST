# Plano Faturacao LuGEST

## Objetivo

Criar um menu `Faturacao` profissional no LuGEST que centralize:

- orcamentos vendidos
- encomendas associadas e estado operacional
- faturas emitidas
- pagamentos recebidos
- comprovativos anexados
- dashboard comercial e de cobranca

O foco e dar seguimento ao ciclo comercial completo sem complicar o registo da encomenda nem misturar faturacao com expedicao ou producao.

## Principios funcionais

- A venda nasce no orcamento aprovado ou na encomenda direta.
- A producao e a expedicao continuam a ter estado proprio.
- A faturacao fica separada da expedicao.
- O pagamento fica separado da faturacao.
- Um registo comercial pode ter varias faturas.
- Uma fatura pode ter um ou varios pagamentos.
- O estado de cobranca deve refletir saldo real, atrasos e pendencias.

## Estrutura funcional

### Lista principal

Cada linha do menu `Faturacao` mostra:

- numero do registo de faturacao
- numero do orcamento
- numero da encomenda
- cliente
- estado da producao
- estado da expedicao
- estado da faturacao
- estado do pagamento
- valor vendido
- saldo em aberto

### Detalhe do registo

No detalhe ficam agrupados:

- origem comercial
- cliente
- data da venda
- data de vencimento
- observacoes
- resumo financeiro
- historico de faturas
- historico de pagamentos

### Dashboard

O dashboard do menu deve mostrar pelo menos:

- total vendido
- total faturado
- total recebido
- saldo em aberto
- numero de registos em atraso
- numero de registos por faturar

## Modelo de dados

### `faturacao_registos`

Cabecalho comercial do registo.

Campos principais:

- `numero`
- `ano`
- `origem`
- `orcamento_numero`
- `encomenda_numero`
- `cliente_codigo`
- `cliente_nome`
- `data_venda`
- `data_vencimento`
- `valor_venda_manual`
- `estado_pagamento_manual`
- `obs`
- `created_at`
- `updated_at`

### `faturacao_faturas`

Documentos faturados ligados ao registo.

Campos principais:

- `registo_numero`
- `documento_id`
- `numero_fatura`
- `serie`
- `guia_numero`
- `data_emissao`
- `data_vencimento`
- `valor_total`
- `caminho`
- `obs`

### `faturacao_pagamentos`

Recebimentos ligados ao registo e, quando aplicavel, a uma fatura.

Campos principais:

- `registo_numero`
- `pagamento_id`
- `fatura_documento_id`
- `data_pagamento`
- `valor`
- `metodo`
- `referencia`
- `titulo_comprovativo`
- `caminho_comprovativo`
- `obs`

## Regras de negocio

### Origem do registo

- Se um orcamento estiver aprovado/vendido, pode gerar registo de faturacao.
- Se existir encomenda direta sem orcamento, tambem pode gerar registo.
- O registo de faturacao e unico por origem comercial.

### Estado de faturacao

- `Por faturar`: sem faturas
- `Parcialmente faturada`: valor faturado abaixo do valor vendido
- `Faturada`: valor faturado igual ou superior ao valor vendido

### Estado de pagamento

- `Por receber`: sem recebimentos
- `Parcialmente paga`: recebido abaixo do faturado
- `Paga`: recebido igual ou superior ao faturado
- `Em atraso`: existe vencimento ultrapassado com saldo aberto

Pode existir override manual do estado de pagamento quando a empresa precisar de uma leitura comercial diferente.

## Fluxo recomendado

1. O comercial aprova/vende o orcamento.
2. O sistema converte em encomenda ou liga a encomenda ja existente.
3. O menu `Faturacao` mostra o registo comercial.
4. A equipa associa a fatura emitida.
5. A equipa financeira associa pagamentos e comprovativos.
6. O dashboard passa a refletir faturado, recebido e saldo.

## Boas praticas adotadas

- separacao entre estados operacionais e financeiros
- suporte a faturacao parcial
- suporte a recebimentos parciais
- suporte a anexos/documentos
- visibilidade clara de atraso e saldo
- ligacao com orcamento, encomenda e guias

## Validacao minima

Em instalacao nova ou release nova validar sempre:

1. abrir o menu `Faturacao`
2. confirmar que aparecem vendas elegiveis
3. abrir um registo e associar uma fatura
4. confirmar o documento na base
5. associar um pagamento
6. confirmar o comprovativo e o saldo
7. confirmar que o dashboard atualiza

## Resultado esperado

Com esta estrutura, o LuGEST passa a ter um fluxo comercial mais profissional:

- vende
- produz
- expede
- fatura
- recebe
- acompanha o saldo e os atrasos

Tudo sem perder ligacao ao resto do ERP e sem obrigar a trabalho duplicado.
