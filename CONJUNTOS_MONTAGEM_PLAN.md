# Melhoria Proposta: Conjuntos, Componentes de Stock e Montagem

## Objetivo

Permitir orcamentar e repetir "conjuntos" ou "projetos" completos, onde coexistem:

- pecas fabricadas internamente
- componentes de stock (`produtos`) como parafusos, porcas, rolamentos, motores, etc.
- uma fase nova chamada `Montagem`

O objetivo e conseguir:

1. orcamentar um conjunto completo
2. guardar esse conjunto como modelo reutilizavel
3. converter o orcamento em encomenda mantendo a estrutura
4. dar baixa dos componentes de stock apenas no momento de `Montagem`

## O que o software atual ja faz

- O orcamento trabalha por linhas tecnicas (`orcamentos` + `orcamento_linhas`).
- A conversao de orcamento para encomenda transforma cada linha em peca fabricada.
- As operacoes de fabrico sao geradas em fluxo por peca.
- O stock de `produtos` ja existe e suporta movimentos (`produtos_mov`).
- Existe uma operacao obrigatoria atual: `Embalamento`.

## Problema do modelo atual

Hoje o sistema esta otimizado para fabricar pecas, nao para gerir um conjunto com:

- subconjuntos
- componentes comprados/stock
- consumo final em montagem
- reutilizacao como "receita" ou "modelo"

Se tentarmos meter parafusos e componentes de stock como se fossem pecas normais, vamos misturar conceitos errados:

- uma peca exige `material`, `espessura`, `ref_interna`, `fluxo de operacoes`
- um parafuso de stock nao precisa disso tudo
- a baixa de um componente de stock deve acontecer no consumo de montagem, nao como se fosse producao

## Recomendacao principal

### 1. Criar um novo conceito: Modelo de Conjunto

Este deve ser o centro da melhoria.

Sugestao de entidade:

- `conjuntos_modelo`
  - `codigo`
  - `nome`
  - `descricao`
  - `cliente_tipo` ou `familia`
  - `ativo`
  - `versao`
  - `notas`
  - `montagem_tempo_min`
  - `preco_montagem_hora`
  - `margem_default`
  - `updated_at`

- `conjuntos_modelo_itens`
  - `conjunto_codigo`
  - `ordem`
  - `tipo_item`
  - `codigo_item`
  - `descricao`
  - `quantidade`
  - `unid`
  - `origem`
  - `operacoes`
  - `custo_unit`
  - `observacoes`

### 2. Tipos de item no conjunto

Cada linha do conjunto deve poder ser de um destes tipos:

- `peca_fabricada`
  - gera linha tecnica / peca / OPP
- `produto_stock`
  - referencia um `produto` do stock
  - nao gera peca de fabrico
  - so gera consumo na `Montagem`
- `servico`
  - opcional, para transporte, montagem externa, pintura subcontratada, etc.

## Como eu aconselho a funcionar no dia a dia

### A. Criar o conjunto modelo

Exemplo: `CJ-0001 | Maquina Tostas Mistas`

Dentro do conjunto guardas:

- pecas fabricadas
  - estrutura
  - tampas
  - suportes
- produtos de stock
  - parafuso M6
  - porca M6
  - anilha
  - resistencia
  - termostato
- parametros de montagem
  - tempo de montagem por conjunto
  - custo/hora
  - opcionalmente quantidade de operadores

### B. Orçamentar a partir do conjunto

No orcamento, em vez de construirmos tudo de raiz, o utilizador escolhe:

- cliente
- modelo de conjunto
- quantidade de conjuntos

O sistema expande o conjunto em:

- linhas fabricadas
- linhas de componentes de stock
- linha de montagem

No PDF do orcamento podes escolher mostrar:

- vista resumida por conjunto
- vista detalhada por componentes

### C. Guardar para repetir

Se o cliente pedir outra vez a mesma maquina:

- abrir modelo existente
- duplicar ou aplicar diretamente
- alterar so quantidades, acabamentos ou componentes diferentes

Isto evita voltar a construir o mesmo orcamento do zero.

## Recomendacao importante sobre a Montagem

### Nao aconselho meter `Montagem` em todas as pecas

Isso tornava o fluxo pesado e confuso.

O melhor e:

- as `pecas_fabricadas` mantem o fluxo normal delas
  - corte
  - quinagem
  - soldadura
  - pintura
  - embalamento, se aplicavel
- o `conjunto` ou `subconjunto` tem uma ordem propria de `Montagem`

### Melhor abordagem

Criar uma entidade especifica para a execucao do conjunto:

- `conjunto_instancias` ou `encomenda_conjuntos`
  - `id`
  - `numero_encomenda`
  - `conjunto_codigo`
  - `descricao`
  - `qtd_conjuntos`
  - `estado`
  - `montagem_estado`
  - `montagem_inicio`
  - `montagem_fim`
  - `montagem_operador`

- `conjunto_consumos`
  - `conjunto_instancia_id`
  - `produto_codigo`
  - `qtd_planeada`
  - `qtd_consumida`
  - `origem`

## Regra funcional recomendada

### Quando e que o stock baixa?

Os produtos de stock devem:

- aparecer como necessidade do conjunto logo na encomenda
- poder ser reservados cedo
- mas a baixa real deve acontecer ao concluir `Montagem`

Isto bate certo com o que queres.

### Evento de baixa

Quando o operador fecha `Montagem`, o sistema:

1. valida que ha stock suficiente dos componentes
2. faz `produtos_mov` para cada produto associado
3. grava `origem = CONJUNTO_MONTAGEM`
4. grava `ref_doc = id do conjunto / encomenda`
5. fecha a montagem

## Como encaixa no software atual

### O que pode ficar igual

- `produtos`
- `produtos_mov`
- fluxo de pecas fabricadas
- OPP / OF das pecas
- notas de encomenda a fornecedores

### O que deve ser acrescentado

- modelos de conjunto
- itens do conjunto
- instancia do conjunto dentro da encomenda
- operacao/posto de `Montagem`
- consumo de stock de produtos por conjunto

## MVP recomendado

### Fase 1

Entregar primeiro isto:

- criar e guardar modelos de conjunto
- permitir inserir no orcamento um conjunto parametrizado
- ter dois tipos de linha:
  - fabricada
  - produto de stock
- somar automaticamente custo dos componentes de stock
- acrescentar linha de custo de montagem

### Fase 2

- converter orcamento em encomenda preservando o conjunto
- criar registo de `Montagem`
- permitir baixa de stock no fecho da montagem

### Fase 3

- reservas de stock por conjunto
- subconjuntos
- versoes do conjunto
- clonagem de modelos
- historico de alteracoes por cliente

## Sugestao de UX

### Novo menu

- `Conjuntos / Projetos`

### Dentro do menu

- `Modelos`
- `Componentes`
- `Custos e Montagem`
- `Duplicar modelo`
- `Gerar orçamento`

### No orçamento

Botao novo:

- `Adicionar conjunto`

Quando clicas:

1. escolhes o modelo
2. defines quantidade
3. decides se queres expandir em detalhe ou resumido
4. o sistema gera as linhas

## O que eu aconselho tecnicamente

### Nao recomendo

- usar produtos de stock como se fossem pecas normais
- obrigar cada componente de stock a ter `ref_interna`
- misturar a `Montagem` com o fluxo de todas as pecas

### Recomendo

- manter duas naturezas distintas:
  - fabricacao
  - consumo de stock
- criar uma camada acima chamada `conjunto`
- usar a `Montagem` como o momento de fecho/consumo

## Ordem ideal de implementacao

1. Normalizar `Montagem` como operacao reconhecida pelo sistema.
2. Criar tabelas/modelos de conjunto.
3. Permitir adicionar conjunto a um orcamento.
4. Levar o conjunto para a encomenda.
5. Criar baixa automatica dos `produtos` no fecho da montagem.

## Conclusao

Sim, a ideia faz muito sentido e eu aconselho fortemente este caminho.

Em resumo:

- o que vendes passa a poder ser um `conjunto`
- esse conjunto pode misturar `pecas fabricadas + componentes de stock`
- a baixa dos componentes acontece em `Montagem`
- o conjunto fica guardado e reutilizavel no futuro

Este e o caminho mais limpo para nao estragar o que ja funciona no fluxo atual.
