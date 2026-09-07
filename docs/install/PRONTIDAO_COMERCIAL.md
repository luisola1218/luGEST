# Prontidao Comercial LuisGEST

Versao auditada: 2026.09.04.1

## Estado tecnico

- Aplicacao desktop, base MySQL, PDFs, stock, conjuntos, planeamento, OPP, expedicao, transportes e relatorios validados.
- Instalacao e atualizacao preservam configuracao, licenca, branding, PDFs gerados e backups locais.
- Atualizacoes exigem HTTPS e manifesto validado com hashes SHA-256 obrigatorios
  para o ZIP e para o reparador; tokens privados nao sao persistidos pelo fluxo
  novo.
- Configuracao, logs e estado do atualizador ficam fora da pasta da aplicacao,
  com escrita JSON atomica e recuperacao por backup.
- O pacote comercial inclui e valida os quatro launchers de atualizacao, evitando
  releases incompletas.
- Auditoria de seguranca sem ocorrencias altas ou medias.
- Matéria-Prima, Produtos e Notas de Encomenda usam o mesmo padrão de carteira,
  catálogo e inspetor, validado a partir de 1180 x 760 sem scroll horizontal.
- Matriz automatizada segura de compilação, dependências, arquitetura,
  segurança, licenciamento, schema, cálculo laser, nesting e controlos de
  orçamentos concluída sem falhas na versão indicada.
- Dependências Python bloqueadas por versões exatas e verificadas antes de cada
  build comercial.
- Diagnóstico de exceções e eventos de execução isolado da interface e com
  rotação de ficheiros.
- Fundação de licenciamento por assinatura Ed25519 criada, ainda sem bloquear a
  aplicação enquanto não forem aprovadas as regras comerciais.
- Build Windows otimizado validado com 1229 ficheiros e cerca de 232,2 MB; mapas
  abrem no navegador externo para evitar distribuir o motor WebEngine completo.

## Utilizacao recomendada nesta fase

O pacote esta apto para demonstracao e piloto controlado em cliente, com backup
e acompanhamento tecnico. Ainda nao deve ser apresentado como produto pronto
para venda geral: a ativacao de licencas, assinatura digital, validacao numa base
de staging e os fechos legais/fiscais abaixo continuam pendentes.

## Fechos externos antes de venda geral

1. Revogar qualquer token GitHub que tenha sido anteriormente guardado em configuracoes locais.
2. Assinar digitalmente o executavel, o instalador e os pacotes de atualizacao com certificado de code signing.
3. Publicar o manifesto e os assets num endpoint HTTPS controlado; para maior
   garantia de origem, acrescentar assinatura do manifesto alem dos hashes.
4. Concluir testes de instalacao, atualizacao e rollback num computador Windows
   limpo e numa rede multiutilizador real, usando uma base de staging.
5. Formalizar licenca, contrato de suporte, politica de privacidade/RGPD e procedimento de recuperacao de desastre.
6. Submeter e obter certificacao da Autoridade Tributaria antes de comercializar o modulo como software de faturacao certificado.

## Regra fiscal

Os testes internos de SAF-T, sequencialidade, anulacao, hash, ATCUD e preparacao de comunicacao nao substituem a certificacao formal da Autoridade Tributaria. Ate existir numero de certificado valido, os PDFs de faturacao de demonstracao devem permanecer identificados como exemplos sem validade fiscal.
