# Manual de Operacao Profissional do luGEST

Versao: 2026-04-07  
Aplicacao: `luGEST Qt`  
Objetivo: manual de utilizacao com capturas reais e exemplo pratico de fluxo operacional.

## 1. Objetivo do manual

Este manual foi preparado para utilizadores finais do `luGEST` e foca os menus mais importantes para o funcionamento diario da empresa.

O documento foi organizado para que qualquer utilizador consiga:

- perceber a logica global do software
- criar e gerir um orcamento
- converter um orcamento aprovado em encomenda
- acompanhar a encomenda no planeamento
- operar a encomenda no posto `Operador`
- consultar os menus de apoio mais usados

## 2. Fluxo principal do software

O fluxo principal recomendado no `luGEST` e este:

1. Criar o orcamento.
2. Adicionar linhas, pecas, DXF/DWG e operacoes.
3. Rever valores, nesting e desconto.
4. Aprovar o orcamento.
5. Converter em encomenda.
6. Planear a producao.
7. Executar no menu `Operador`.

Neste manual foi usado um exemplo real de demonstracao:

- Orcamento: `ORC-2026-0001`
- Encomenda gerada: `BARCELBAL0001`
- Cliente: `CL0010 - Atlas Planeamento`

## 3. Orcamentos

### 3.1 Objetivo do menu

O menu `Orcamentos` e o centro da fase comercial e tecnica.  
Aqui e onde se constroem propostas, se definem linhas, se analisam pecas e se prepara a conversao para producao.

### 3.2 Lista de orcamentos

![Lista de orcamentos](capturas_reais/manual_orcamentos_lista.png)

Nesta vista o utilizador consegue:

- pesquisar por numero, cliente ou encomenda
- filtrar por estado
- filtrar por ano
- abrir um orcamento existente
- criar um novo orcamento
- remover um orcamento quando aplicavel

### 3.3 Detalhe do orcamento

![Detalhe do orcamento](capturas_reais/manual_orcamento_detalhe.png)

No detalhe do orcamento, as zonas principais sao:

- `Dados do Cliente`: identifica quem recebe a proposta
- `Dados do Orcamentista`: posto de trabalho, responsavel e contexto comercial
- `Resumo financeiro`: linhas, transporte, desconto, subtotal, IVA e total
- `Notas do orcamento`: transporte, operacoes e texto para PDF
- `Referencias do orcamento`: linhas, DXF/DWG, nesting e configuracoes

### 3.4 O que se faz dentro do orcamento

Dentro do detalhe do orcamento, o utilizador pode:

- adicionar linhas manuais
- carregar pecas por `Peca Unit. DXF/DWG`
- carregar varios ficheiros por `Lote DXF/DWG`
- configurar perfis de laser
- detalhar operacoes adicionais
- abrir o `Plano de chapa`
- aplicar desconto global
- gerar PDF
- aprovar ou rejeitar
- converter em encomenda

### 3.5 Leitura da tabela de linhas

Na tabela de referencias do orcamento, cada linha representa uma unidade comercial ou tecnica.

As colunas mais importantes sao:

- `Tipo`: peca fabricada, produto de stock ou servico
- `Ref./Cod.`: referencia interna
- `Ref. Ext.`: referencia externa do cliente ou desenho
- `Descricao`: designacao comercial/tecnica
- `Material/Produto`
- `Esp./Unid`
- `Operacao`
- `Tempo`
- `Qtd`
- `Preco`
- `Total`

### 3.6 Exemplo pratico no menu Orcamentos

No exemplo deste manual foi criado um orcamento com:

- 2 pecas fabricadas em `Aco inox 4 mm`
- 1 artigo de stock
- transporte definido
- desconto global aplicado

Depois:

1. o orcamento foi guardado
2. foi marcado como `Aprovado`
3. foi convertido em encomenda

## 4. Encomendas

### 4.1 Objetivo do menu

O menu `Encomendas` recebe o trabalho ja aprovado e transforma-o num registo operacional de producao.

Aqui a logica ja nao e comercial.  
O foco passa a ser:

- materiais
- espessuras
- pecas
- montagem
- reservas
- transporte

### 4.2 Lista de encomendas

![Lista de encomendas](capturas_reais/manual_encomendas_lista.png)

Na lista de encomendas o utilizador consegue:

- pesquisar por numero
- filtrar por cliente
- filtrar por estado e ano
- abrir a encomenda
- editar cabecalho
- remover encomenda, quando permitido

### 4.3 Detalhe da encomenda

![Detalhe da encomenda](capturas_reais/manual_encomenda_detalhe.png)

No detalhe da encomenda, as zonas mais importantes sao:

- `Cabecalho`: numero, cliente, nota cliente, transporte, estado e custos
- `Materiais`: agrupamento por material
- `Espessuras`: detalhe por espessura e tempos
- `Pecas`: pecas reais da encomenda, com operacoes e estado
- `Montagem / Componentes de stock`: consumiveis, produtos e faltas

### 4.4 O que significa cada nivel

O detalhe da encomenda esta organizado por niveis:

1. `Material`
2. `Espessura`
3. `Peca`

Isto permite perceber:

- quanto tempo existe por espessura
- que operacoes estao associadas
- que pecas estao em preparacao, em producao ou concluidas

### 4.5 Exemplo pratico do fluxo

No exemplo deste manual:

- o orcamento `ORC-2026-0001` foi convertido em `BARCELBAL0001`
- a encomenda ficou com `Aco inox | 4 mm`
- foram geradas 2 pecas fabricadas
- foi mantido 1 item de stock para montagem

Este menu e o ponto de passagem entre o trabalho comercial e o trabalho operacional.

## 5. Planeamento

### 5.1 Objetivo do menu

O menu `Planeamento` organiza a carga da producao por semana e por operacao.

Serve para:

- encaixar encomendas na semana
- distribuir blocos de producao
- ver backlog
- controlar ocupacao

### 5.2 Vista de planeamento

![Planeamento](capturas_reais/manual_planeamento.png)

Na pagina de planeamento tens:

- `Pendentes`: backlog por planear
- `Quadro semanal`: grelha horaria por dias
- `Blocos ativos`: tempo atualmente planeado
- `Carga semanal`: leitura do total encaixado

### 5.3 Como usar o planeamento

Fluxo recomendado:

1. escolher a operacao no topo, por exemplo `Corte Laser`
2. analisar as encomendas pendentes
3. colocar um bloco na semana ou usar auto-planeamento
4. verificar conflitos ou bloqueios
5. rever a carga semanal

### 5.4 Exemplo pratico

No exemplo deste manual:

- a encomenda `BARCELBAL0001` entrou no backlog
- foi colocado um bloco de `Corte Laser`
- o quadro semanal passou a mostrar esse encaixe numa data futura

Isto mostra ao utilizador como o trabalho sai da encomenda e entra no calendario operacional.

## 6. Operador

### 6.1 Objetivo do menu

O menu `Operador` e o posto de execucao da producao.

E aqui que o operador:

- entra numa encomenda
- escolhe uma peca
- inicia a operacao
- pausa, termina ou regista avaria
- consulta desenho e etiquetas

### 6.2 Vista do operador

![Operador](capturas_reais/manual_operador_detalhe.png)

O ecrã do operador esta organizado por:

- cabecalho da encomenda
- filtros de estado
- grupos por material/espessura
- tabela de pecas
- botoes de operacao

### 6.3 Leitura da tabela de pecas

Na tabela de pecas, o operador acompanha:

- referencia interna e externa
- estado atual
- operacao atual
- operador em curso
- produzido
- tempo
- plano
- operacoes pendentes

### 6.4 Botoes principais do operador

Os botoes principais deste menu sao:

- `Iniciar`
- `Finalizar`
- `Retomar`
- `Interromper`
- `Dar Baixa`
- `Ver desenho`
- `Etiquetas`
- `Atualizar`
- `Registar Avaria`
- `Alertar Chefia`

### 6.5 Exemplo pratico

No exemplo deste manual:

- a encomenda `BARCELBAL0001` foi aberta no `Operador`
- uma das pecas ficou `Em producao`
- a segunda permaneceu em `Preparacao`

Isto permite demonstrar claramente a diferenca entre:

- trabalho iniciado
- trabalho ainda pendente
- operacoes futuras por executar

## 7. Menus de apoio

Depois do fluxo principal, estes menus dao suporte ao funcionamento diario.

### 7.1 Materia-Prima

![Materia-Prima](capturas_reais/manual_materia_prima.png)

Objetivo:

- gerir lotes
- consultar stock de chapa, perfil, tubo e retalho
- reservar material
- analisar disponibilidade

Quando usar:

- antes do nesting
- para entradas e saidas de material
- para controlar formatos e retalhos

### 7.2 Produtos

![Produtos](capturas_reais/manual_produtos.png)

Objetivo:

- gerir artigos de stock interno
- controlar consumiveis e componentes
- dar suporte a montagem, expedicao e outras areas

Quando usar:

- para ajustar quantidades
- para criar produtos internos
- para consultar preco ou unidade

### 7.3 Clientes

![Clientes](capturas_reais/manual_clientes.png)

Objetivo:

- manter a ficha comercial dos clientes
- garantir que orcamentos e encomendas usam dados corretos

Quando usar:

- antes de criar um novo orcamento
- para atualizar morada, NIF, contacto e email

## 8. Fluxo recomendado para utilizadores

### 8.1 Fluxo comercial e tecnico

1. Abrir `Orcamentos`.
2. Criar o registo.
3. Inserir linhas, pecas e operacoes.
4. Rever valores e PDF.
5. Aprovar.
6. Converter em encomenda.

### 8.2 Fluxo operacional

1. Abrir `Encomendas`.
2. Confirmar materiais, espessuras e pecas.
3. Abrir `Planeamento`.
4. Planear a operacao.
5. Abrir `Operador`.
6. Iniciar a producao.

## 9. Boas praticas

- Guardar sempre depois de alteracoes importantes.
- Confirmar cliente e nota cliente antes de aprovar.
- Rever materiais e espessuras depois da conversao em encomenda.
- No planeamento, verificar sempre a semana e os bloqueios antes de encaixar.
- No operador, confirmar a peca selecionada antes de iniciar.

## 10. Erros comuns a evitar

- aprovar orcamento sem rever linhas e desconto
- converter para encomenda sem confirmar o cliente
- planear na operacao errada
- iniciar a peca errada no operador
- assumir que o tempo comercial e igual ao tempo operacional sem confirmar no planeamento

## 11. Conclusao

O `luGEST` foi desenhado para ligar a fase comercial a producao real.

O fluxo mais importante do software e:

`Orcamentos -> Encomendas -> Planeamento -> Operador`

Se o utilizador dominar bem estes quatro menus, consegue trabalhar a maior parte do processo diario com seguranca e consistencia.
