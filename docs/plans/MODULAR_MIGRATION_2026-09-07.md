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
