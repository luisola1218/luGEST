# Auditoria de fluidez e robustez - 23/07/2026

Versao final desta revisao: `2026.07.23.13`

## Ambito

Medicao local dos 20 menus permitidos ao utilizador administrativo, com a base de
dados atual, incluindo construcao da pagina e primeiro refresh. A auditoria foi
executada no runtime Qt real e complementada com pesquisa de chamadas bloqueantes,
temporizadores, threads, reloads e caches.

## Resultados

| Menu | Tempo total |
| --- | ---: |
| Dashboard | 361 ms |
| Pulse | 265 ms |
| Operador | 240 ms |
| Avarias | 220 ms |
| Fornecedores | 64 ms |
| Orcamentos | 64 ms |
| Restantes 14 menus | 9-35 ms |

Conclusao: nao existe atualmente um gargalo geral de navegacao. Todos os menus
medidos ficaram abaixo de 400 ms e o backend ja aplica cache de dados por 30
segundos e caches operacionais de 3 a 5 segundos.

## Melhorias aplicadas

- Corrigido o ciclo de vida da thread de atualizacao automatica. O worker deixa de
  ser libertado durante a emissao do sinal Qt.
- O smoke test passou a testar cinco ciclos completos da thread de atualizacao em
  cada compilacao.
- Falhas ao construir ou atualizar um menu ficam contidas nesse menu e deixam de
  contaminar a aplicacao inteira.
- Erros de refresh automatico deixam um estado visivel sem abrir caixas repetidas.
- Orcamentos passou para um workbench com referencias em primeiro plano e inspetor
  lateral recolhivel para Cliente, Condicoes e Financeiro.
- Os seletores do inspetor e das Condicoes distribuem a largura disponivel sem
  setas, cortes ou scroll horizontal a partir de 1180 px.
- Os comandos de transporte adaptam-se a duas linhas e o Financeiro separa total,
  base tributavel, IVA, composicao e ajustes comerciais.
- A barra principal de Orcamentos passou para uma unica linha compacta, sem perder
  comandos e sem ultrapassar a pagina a partir de 1180 px.
- A lista de orcamentos apresenta valor visivel e contagem por estado.
- Notas de Encomenda passou a apresentar carteira, indicadores, documento,
  fornecedor, resumo e linhas num workbench responsivo.
- A carteira de aprovisionamento ganhou metricas visuais, numero de documento
  estavel e grelha operacional adaptativa: abaixo de 1450 px preserva codigo,
  material, descricao, fornecedor, quantidade, preco, IVA, total e entrega sem
  scroll horizontal; em ecras largos apresenta as 18 colunas.
- Clientes e Fornecedores usam uma lista principal e um inspetor por separadores
  para Identificacao, Contacto e contexto Comercial/Compras.
- Produtos passou a integrar um catalogo de conjuntos parametrizados com codigo de
  parametrizacao, BOM, ficha tecnica, margem, custo, preco e origem viva dos precos.
- Produtos passou para um workbench com indicadores de carteira, filtros
  identificados, barra de comandos unica, catalogo de leitura rapida e inspetor
  lateral por Identificacao, Stock e precos e Movimentos.
- As colunas secundarias deixaram de competir com codigo, descricao, quantidades e
  preco na grelha principal; continuam disponiveis no inspetor e na grelha completa.
- Os testes de operador e expedicao drenam as gravacoes assincronas antes da
  limpeza e deixam de repor documentos VERIFY depois de terminarem.
- Existe uma limpeza dedicada, com preview e backup, para artefactos tecnicos de
  verificacao.

Na verificacao final desta iteracao, a navegacao quente apresentou um maximo de
24 ms e a navegacao fria um maximo de 290 ms. A leitura inicial da base demorou
233 ms.

## Revisao detalhada 2026.07.23.11

### Materia-Prima

- O menu passou de formulario extenso para um workbench de stock.
- O topo apresenta registos, disponibilidade total e valor em stock.
- Pesquisa, formato, material, espessura, local, estado e stock disponivel ficam
  acessiveis numa unica faixa responsiva.
- A grelha principal preserva apenas ID, material, formato, lote, espessura,
  quantidade, disponibilidade e estado; a grelha completa continua acessivel.
- O inspetor lateral separa Identificacao, Geometria e Stock e valor.
- Preco unitario e valor do stock no inspetor acompanham o calculo do material
  selecionado em tempo real.
- Etiqueta agrupa pre-visualizacao, impressao e gravacao num menu unico.
- Preview PDF continua a permitir escolher os formatos a incluir.
- A pagina mede 1131 px de largura minima e nao cria scroll horizontal a 1180 px.

### Notas de Encomenda

- A carteira apresenta documentos, ativos e valor total antes do detalhe.
- A lista ganhou um inspetor lateral com numero, estado, fornecedor, entrega,
  valor, linhas e contexto da fase de aprovisionamento.
- Duplo clique ou Abrir nota entra no documento completo; Voltar regressa a lista.
- Criacao, aprovacao, cotacao, geracao de NEs, rececao, anexos, fornecedores,
  envio e PDF mantiveram o mesmo contrato funcional.
- A tabela principal e o detalhe adaptativo nao criam scroll horizontal a 1180 px.

### Correcao e limpeza

- Repostos os nomes funcionais `Preview PDF` em Materia-Prima e `Preview stock`
  em Produtos, mantendo os testes e a linguagem do utilizador.
- O ensaio de ciclo interno deixou de exigir cinco clientes e passou a trabalhar
  corretamente com a base comercial atual de dois clientes.
- O mesmo ensaio seleciona uma chapa real, configura Corte Laser e Embalamento e
  limpa sempre orcamento, OF, planeamento, guia, faturacao, pagamento e ficheiros.
- A limpeza tecnica passou a cobrir `plano`, `faturacao` e
  `faturacao_registos`, além das restantes entidades `VERIFY`.
- Foram removidos, com backup, uma guia de teste e dois blocos de planeamento
  orfaos encontrados durante a propria auditoria.
- Dois retalhos anteriores ao registo de proveniencia receberam a marca explicita
  `LEGADO-SEM-ORIGEM`; nao foi inventado um lote de fornecedor.
- A verificacao final nao encontrou stock invalido, referencias duplicadas, OPP
  duplicadas, OF duplicadas, planeamento orfao nem expedicao orfa.

### Matriz executada

- `verify_core_flows.py`: aprovado, incluindo compras, conjuntos, OF,
  planeamento, operador, expedicao, relatórios, OPP, nesting e baixa parcial.
- 17 verificacoes complementares: aprovadas depois da correcao do ensaio interno.
- MySQL: esquema, persistencia e metadados sem distincao indevida de maiusculas
  aprovados.
- PDFs: dossier tecnico, orcamento vertical, faturacao, relatorios, etiquetas e
  stocks filtrados aprovados.
- Fiscalidade interna: guardas, SAF-T, sequencias, hash, ATCUD e preparacao de
  comunicacao aprovados, sem substituir certificacao formal.
- Dependencias: `pip check` sem pacotes quebrados.
- Compilacao Python: todos os modulos de aplicacao e scripts compilam.
- Desempenho Qt: carga inicial 0,233 s, pagina fria maxima 0,290 s e navegacao
  quente maxima 0,024 s.

## Revisao detalhada 2026.07.23.12

### Materia-Prima

- A grelha principal passou a apresentar `Dimensoes (mm)` imediatamente depois
  de Material, no formato `comprimento x largura`.
- A coluna composta preserva o comportamento responsivo e nao introduz scroll
  horizontal a 1180 px.

### Notas de Encomenda

- O detalhe ganhou um inspetor da linha selecionada com codigo, estado, descricao,
  material, espessura, origem, quantidade, preco unitario, total e fornecedor.
- O apoio a compra apresenta fornecedor habitual, ultima compra, preco medio,
  prazo e alternativas disponiveis em stock.
- A visibilidade das colunas usa tres modos de ecra; a vista corrente preserva
  as colunas operacionais e transfere o contexto secundario para o inspetor.
- O aconselhamento de compra e calculado uma vez por linha e reutilizado na
  grelha, tooltips e inspetor, eliminando calculos repetidos no mesmo refresh.
- O ensaio com 60 linhas demorou 16 ms e nao criou scroll horizontal a 1180 nem
  a 1900 px.

### Fluidez

- A preparacao de emails Outlook em Compras, Orcamentos e Planeamento deixou de
  bloquear a thread visual ate 30 segundos.
- Os processos externos usam agora `QProcess`, com timeout, limpeza automatica e
  fallback para o cliente de email predefinido.
- A medicao final registou carga inicial de 0,185 s, pagina fria maxima de
  0,258 s e navegacao quente maxima de 0,032 s.
- Os 165 ficheiros Python compilaram e a matriz completa de fluxos, integridade,
  nesting, PDFs filtrados e seguranca terminou sem falhas altas ou medias.

## Revisao detalhada 2026.07.23.13

### Encomendas e ordens de fabrico

- A lista passou a funcionar como carteira de ordens, com total, ordens ativas e
  progresso medio antes da abertura do detalhe.
- Pesquisa, estado, ano e cliente ficaram numa unica faixa sem sobreposicao de
  campos; a pesquisa usa debounce de 180 ms.
- Os comandos foram renomeados por intencao: Nova ordem de fabrico, Abrir ordem,
  Editar ordem, Eliminar ordem, Adicionar conjunto, Pre-visualizar OF e Abrir
  desenho.
- O detalhe apresenta identidade da OF, cliente, referencia, entrega, numero de
  pecas, materiais, progresso, logistica e material cativado numa leitura unica.
- Materiais e espessuras passaram para uma hierarquia vertical, libertando largura
  para a grelha principal de pecas.
- A grelha de pecas distingue quantidade planeada e produzida; a carteira apresenta
  progresso visual por ordem.
- Montagem e componentes ganhou contexto de stock e comandos diretos para Operador
  e Nota de Compra.
- O detalhe mede no maximo 996 x 708 px, nao cria scroll horizontal a 1180 px e
  o menu abre quente em aproximadamente 6 ms.

## Analise geral por prioridade

### P1 - piloto comercial

1. Executar um ensaio continuo de oito horas com dois postos e o servidor MySQL
   real, incluindo impressoras, leitores de codigo e encerramento de turno.
2. Ensaiar restauracao integral de backup numa maquina limpa e registar o tempo
   de recuperacao.
3. Fechar a matriz de permissoes por funcao e testar cada perfil sem privilegios
   administrativos.
4. Validar instalacao, ativacao, atualizacao e desinstalacao num Windows sem
   ferramentas de desenvolvimento.

### P2 - experiencia consistente

1. Aplicar o padrao de workbench usado em Orcamentos aos menus operacionais que
   ainda combinam demasiados campos e acoes na mesma faixa.
2. Uniformizar estados vazios, carregamento, erro recuperavel e confirmacoes
   destrutivas em todos os menus.
3. Acrescentar atalhos apenas aos fluxos repetitivos medidos no piloto e manter
   todos os comandos acessiveis por rato e toque.
4. Rever contraste, foco de teclado e escala do Windows a 125 e 150 por cento.

### P3 - escala e observabilidade

1. Registar, de forma anonima e local, tempos acima de 500 ms por operacao para
   identificar gargalos com dados reais.
2. Migrar apenas grelhas acima de 1.000 linhas para modelos virtuais e paginados.
3. Mover calculos pesados para workers quando o percentil 95 ultrapassar 700 ms.
4. Criar um painel tecnico de saude com base de dados, fila de gravacoes, backups,
   versao e ultima sincronizacao.

## Proximas melhorias condicionais

Estas alteracoes so devem avancar quando os dados reais justificarem o custo:

1. Executar Dashboard, Pulse, Operador e Avarias em workers apenas se excederem
   consistentemente 700 ms em bases de dados maiores.
2. Introduzir paginacao ou modelos virtuais nas grelhas quando ultrapassarem 1.000
   linhas visiveis.
3. Registar consultas MySQL acima de 500 ms e criar indices a partir de evidencia,
   nao por antecipacao.
4. Medir tempos num posto cliente ligado ao servidor real, incluindo latencia de
   rede e impressoras.

## Criterios de aceitacao

- Arranque e login sem encerramento inesperado.
- Mudanca de menu abaixo de 700 ms no percentil 95.
- Nenhuma operacao de rede executada na thread visual.
- Erros de um menu nao encerram a aplicacao.
- Grelhas e dialogos sem cortes a partir de 1180 x 760.

## Revisao 2026.07.28 - mapas, transportes e nova medicao

- A matriz Qt voltou a medir todos os menus principais com a base atual.
- Carga inicial de dados: 0,238 s.
- Pagina fria mais lenta: 0,275 s.
- Navegacao quente mais lenta: 0,032 s.
- Nenhum menu ultrapassou os limites de 0,700 s definidos para o piloto.
- A pesquisa de Transportes passou a usar debounce de 180 ms. A tabela deixa de
  percorrer encomendas, viagens, tarifas e expedicoes a cada tecla premida.
- A atualizacao automatica global continua limitada a paginas explicitamente
  autorizadas e respeita edicao ativa, cache de backend e gravacoes pendentes.
- A integracao Google Maps e executada pelo navegador do sistema e nao bloqueia a
  interface Qt. Funciona com coordenadas e usa a morada como alternativa.
- A viagem abre o itinerario completo pela ordem das paragens; cada destino pode
  tambem ser aberto isoladamente no mapa.

Conclusao desta revisao: nao foi encontrado um gargalo geral que justificasse
threads adicionais, paginacao antecipada ou reescrita das grelhas atuais. Essas
alteracoes aumentariam complexidade sem ganho mensuravel na base comercial atual.
