# Gravacao de compras e catalogos de precos

`NoteCommands` prepara a nota, materiais, produtos e conjuntos em copias.
Validacao de linhas e fornecedores e calculo de precos antecedem a reserva
do numero e a publicacao. O repositorio verifica os quatro catalogos esperados
e pede uma gravacao bloqueante unica. Perante falha imediata, recupera os
catalogos e a sequencia local. A reserva SQL pode continuar a deixar intervalos.

`PurchasePricing` concentra a conversao de precos por unidade, peso e metros,
assim como a sincronizacao de produtos. `material_lines.py` concentra a
sincronizacao de materia-prima; os adaptadores Qt e desktop usam a mesma regra.
Os conjuntos sao atualizados uma vez, depois das alteracoes de preco.

`PurchaseLines` valida e normaliza linhas usando leituras de catalogo isoladas,
incluindo inferencia de perfis, tubos, barras e cantoneiras. Quantidades e
precos nao finitos sao rejeitados. `PurchaseQuote` cria o pedido de cotacao a
partir das faltas calculadas, usando a nova gravacao de compras.

## Validacao

- 46 verificacoes locais aprovadas; 414 ficheiros Python compilados.
- 487 contratos do backend preservados.
- Testes locais de validacao antes da reserva, falhas de calculo e gravacao,
  recuperacao dos quatro catalogos, conflitos locais e retorno sem referencias
  mutaveis ao estado partilhado.
- Fluxo completo de compras e rececao: 19.633 s.
- Criacao/edicao com releitura de precos, falha de gravacao provocada e pedido
  por falta de stock: 12.266 s.
- Ambos os fluxos SQL reverteram e confirmaram igualdade do conteudo das
  55 tabelas. Nao testam concorrencia entre instalacoes ou falha de energia.

Sugestoes de compra, geometria, rececao, documentos e outras areas ainda usam
capacidades/adaptadores historicos. Esta alteracao conclui o percurso descrito,
nao a migracao integral do ERP.
