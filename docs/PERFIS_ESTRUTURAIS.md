# Perfis estruturais no luGEST

O registo de matéria-prima distingue quatro conceitos:

- **Material**: classe do aço, por exemplo `S275JR`.
- **Série**: geometria normalizada, por exemplo `IPE`, `IPN`, `UPN`, `HEA` ou `HEB`.
- **Tamanho nominal**: designação comercial da série, por exemplo `IPE 220`.
- **Comprimento da barra**: comprimento físico em metros, por exemplo `6 m`.

Assim, `IPE 220` não representa uma espessura de 220 mm. É a designação
comercial que seleciona a ficha técnica: altura, largura de abas, espessuras da
alma e das abas e massa linear.

## Funcionamento

Ao selecionar uma série normalizada, o campo **Tamanho nominal** apresenta
apenas os tamanhos existentes nessa série. A ficha técnica é mostrada no
formulário e o sistema calcula automaticamente:

`peso por barra = kg/m × comprimento em metros`

Exemplo: `IPE 220 × 6 m = 26,2 kg/m × 6 = 157,2 kg`.

O Copiloto de stock reconhece comandos como:

`cria stock de 10 vigas Perfil IPE 220 S275JR com 6 metros de comprimento`

e prepara os campos para confirmação sem gravar automaticamente.

## Séries incluídas

- IPE 80–600
- IPN 80–600
- UPN 80–400
- HEA 100–600
- HEB 100–600
- HEM e UPE (tabela de massa para compatibilidade)

As dimensões são expressas em milímetros e a massa em kg/m. Antes de uma
utilização contratual, deve confirmar-se a edição aplicável da norma e a ficha
do fornecedor.
