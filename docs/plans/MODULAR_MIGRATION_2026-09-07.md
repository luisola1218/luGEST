# Migracao modular de 2026-09-07

Estado: migracao parcial, com componentes funcionais de clientes, orcamentos
e inventario. O [guia de arquitetura](../architecture/MODULAR_MONOLITH.md)
regista as responsabilidades, as dependencias e o trabalho restante.

## Alteracoes verificadas

- Clientes: casos de uso, repositorio e pagina agrupados no modulo de negocio.
  A pagina depende de quatro acoes; o construtor antigo continua compativel.
- Orcamentos: oito editores independentes da pagina, criacao e normalizacao
  de linhas, casos de uso de nesting e dois geradores de PDF.
- Persistencia de nesting separada da validacao; cache restaurado perante
  falha imediata da gravacao principal. O espelho SQL mantem a politica antiga.
- Inventario: consultas e detalhes isolados por repositorio; consumo de stock
  separado da interface e validado antes da escrita. Operador em falta deixou
  de baixar o stock em memoria. Falhas imediatas restauram quantidade e movimento.
- O mapa do backend segue funcoes e fabricas tipadas ate ao modulo de negocio,
  sem executar o runtime, facilitando encontrar a implementacao alem da fachada.

A pagina de orcamentos passou de 10890 para 5350 linhas. O adaptador de
orcamentos passou de 3497 para 1898 linhas. Estes numeros medem a reducao dos
ficheiros centrais; nao sao uma percentagem de conclusao da arquitetura.

## Validacao

- `verify_project.ps1 -SafeOnly`: 25 verificacoes aprovadas; 320 ficheiros Python
  compilados; dependencias instaladas consistentes.
- Contratos dos 487 metodos anteriores preservados.
- Testes independentes: oito cancelamentos de editores e cinco confirmacoes,
  CRUD de clientes, regras de nesting, isolamento de consultas e baixas de stock.
- Verificacao automatica das dependencias dos modulos e das referencias globais
  nos callbacks, incluindo o mapa de implementacoes.

| Teste na base atual | Resultado | Rollback e conteudo das 55 tabelas |
| --- | --- | --- |
| Criar/atualizar/remover estudo de nesting | Aprovado, 5.359 s | Iguais ao inicio |
| Conjuntos, montagem e conversao | Aprovado, 38.068 s | Iguais ao inicio |
| Criar produto, rejeitar entrega sem operador, consumir e reler | Aprovado, 7.342 s | Iguais ao inicio |

O teste de nesting inicialmente dependia de um orcamento existente e nao
encontrou nenhum no snapshot carregado. Foi corrigido para criar o seu proprio
cliente e orcamento temporarios dentro da transacao protegida. A tentativa e
a repeticao terminaram com rollback e conteudos identicos.

Os relatorios JSON locais ficam em `reports/integration/`, ignorados pelo Git.
O rollback nao reverte necessariamente os contadores AUTO_INCREMENT.

## Limites

Os callbacks de composicao ainda ligam parte das regras a implementacoes
historicas. Continuam por migrar a coordenacao da pagina de orcamentos, os
comandos restantes, as outras areas do ERP, o estado global e a inicializacao.
Nao foram validadas concorrencia entre instalacoes, durabilidade assincrona
nem instalacao limpa do executavel. Este resultado nao e uma declaracao de
projeto finalizado ou de prontidao comercial.

## Ciclo de 8 de setembro: comandos, consultas e interface

Produtos passam a ter comandos de criacao, alteracao e remocao com repositorio
explicito e recuperacao do snapshot em caso de erro imediato. Orcamentos passam
a ter comandos e consultas independentes do bridge. A listagem copia apenas
os campos de resumo, evitando copiar estudos pesados de nesting.

A pagina canonica reside no modulo de orcamentos e recebe QuotePageServices.
O workspace possui os widgets, estilos e sinais; o controlador usa `self.view`
explicitamente. Foi corrigido o temporizador de pesquisa em falta. O ficheiro
antigo e apenas um construtor de compatibilidade; o controlador canonico ainda
precisa de ser dividido por casos de utilizacao.

Validacao: 28 verificacoes locais aprovadas, 337 ficheiros compilados e os 487
contratos preservados. Na base atual passaram inventario (15.372 s), conjuntos
e montagem (40.362 s), nesting (5.518 s), fabrico (19.567 s), compras (19.694 s),
faturacao (19.697 s) e planeamento (74.461 s). Todos terminaram com rollback e
conteudo identico nas 55 tabelas verificadas. Os limites sobre contadores,
concorrencia e durabilidade referidos acima continuam a aplicar-se.

### Catalogos de conjuntos

CRUD e consultas dos dois catalogos passam por AssemblyCatalog/AssemblyQueries.
Normalizacao e expansao partilham regras independentes; as duas implementacoes
identicas de expansao foram consolidadas. AssemblyRefresh prepara o catalogo
antes de gravar, rejeita alteracoes locais concorrentes e recupera o snapshot
perante erro imediato. Os testes cobrem produtos, MP, precos laser, valores
manuais, copias aninhadas, falhas de validacao e escrita, preservacao de campos
historicos e separacao entre modelos e conjuntos.

Validacao final deste passo: 29 verificacoes locais, 344 ficheiros compilados,
487 contratos mantidos. Conjuntos/montagem/conversao passaram novamente na base
atual (39.670 s), com rollback e conteudo identico nas 55 tabelas.

### Preparacao de encomenda e faltas de compra

A preparacao de pecas, montagem e tempos saiu do bridge para `order_lines.py`.
Foi corrigida a substituicao acidental do total acumulado pelo tempo de uma
operacao: o teste com duas pecas e montagem confirma 20 minutos, antes 14.
Os alocadores de identificadores continuam ligados ao runtime, e a conversao
completa ainda precisa de uma fronteira transacional entre modulos.

O calculo de necessidades de compra passa por repositorio de leitura explicito.
O saldo de stock e usado uma vez por produto/material ao percorrer as linhas.
Duas linhas de 6 unidades com stock 10 originam uma falta de 2, em vez de zero.
O caso tambem foi validado atraves do backend na base atual, sem consumir stock.

Validacao: 31 verificacoes locais, 351 ficheiros compilados. Na base atual:
inventario 15.753 s, compras 18.863 s, conjuntos/montagem 38.856 s, fabrico
19.285 s e planeamento 71.912 s. Todos com rollback e conteudo das 55 tabelas
igual ao inicio. Guia de manutencao e indice atualizados.
