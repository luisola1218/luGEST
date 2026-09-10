# Ciclo das notas de compra

Rascunhos, aprovacao, marcacao de envio, remocao e conversao por fornecedor
passaram para `purchasing/application/note_lifecycle.py`. O repositorio publica
as notas preparadas e a sequencia local com uma gravacao bloqueante, verifica
alteracoes locais e recupera ambos perante erro imediato. A conversao rejeita
fornecedores sem ficha e cotacoes ja convertidas antes de reservar numeros.

O teste SQL com releitura encontrou duas falhas anteriores:

- Datas de aprovacao e envio nao eram persistidas na tabela historica.
  Agora ficam no payload runtime existente, na mesma transacao SQL, juntamente
  com data de criacao e referencias de orcamentos. Nao houve alteracao de schema.
- A normalizacao substituia estados Enviada e Cotacao aprovada por Aprovada.
  A regra foi isolada em `note_status.py` e preserva estes estados quando ainda
  nao existem entregas, mantendo a progressao parcial/concluida.

Validacao: 44 verificacoes locais aprovadas, 405 ficheiros compilados e os
487 contratos do backend preservados. O ciclo com releitura passou em
12.714 s; o fluxo de compras, conversao e rececao passou em 20.053 s.
Ambos reverteram os dados e confirmaram igualdade do conteudo das 55 tabelas.
Contadores SQL podem ficar com intervalos apos rollback.

Continuam historicos `ne_save`, precos e sincronizacao dos catalogos, rececao
e documentos. As regras de concorrencia entre instalacoes e a eliminacao do
runtime global continuam pendentes; este registo nao declara a migracao total.
