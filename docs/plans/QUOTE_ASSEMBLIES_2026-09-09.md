# Conjuntos e encaminhamento de orcamentos

## Alteracoes

- `quotes/domain/routing.py` concentra classificacao de materia-prima,
  operacoes, prontidao e encaminhamento para laser, serralharia ou montagem.
  Recebe normalizadores explicitos; nao carrega o runtime nem o estado global.
- `presentation/assembly_models.py` contem o editor e o gestor de modelos.
  `saved_assemblies.py` contem o catalogo de conjuntos guardados. Os testes
  abrem estes componentes com um `QWidget` simples, sem a pagina de orcamentos.
- `presentation/group_editor.py` recebe linhas e selecao por valor. Devolve
  o resultado depois da gravacao; cancelamento e erro nao alteram as linhas
  recebidas. A pagina mantem a responsabilidade de atualizar a sua vista.
- `AssemblyCatalog.prepare` separa preparacao de persistencia.
  `AssemblyPair` prepara o modelo e o conjunto, alinha o codigo de parametro
  e pede uma unica gravacao bloqueante. O repositorio verifica os dois estados
  esperados e recupera ambos perante falha imediata.
- O construtor calculado e a opcao "Conjunto + Modelo" usam essa gravacao
  conjunta. Deixaram de executar duas gravacoes independentes.

O controlador de orcamentos passou de 3199 para 2462 linhas neste percurso.
A reducao resulta de componentes com entradas e capacidades explicitas,
testados sem depender de atributos da pagina principal.

## Evidencia

- 42 verificacoes locais aprovadas, 392 ficheiros compilados e 487 contratos
  do backend preservados.

- Testes locais de encaminhamento, CRUD dos gestores, cancelamento, retorno
  de linhas, falha de gravacao e ausencia de alteracoes nos dados de entrada.
- Testes do par: uma chamada de gravacao, dados invalidos no segundo catalogo,
  falha de persistencia e rejeicao de alteracao local concorrente.
- Conjuntos/montagem na base atual: aprovado em 40.754 s.
- Gravacao conjunta dos dois catalogos e releitura: aprovada em 5.917 s,
  com os codigos de parametro iguais e os totais esperados.
- Ambos os percursos SQL foram revertidos, confirmando igualdade do conteudo
  das 55 tabelas. Contadores AUTO_INCREMENT podem ficar com intervalos.

## Limites

A gravacao conjunta usa o mecanismo de persistencia atual; nao demonstra
controlo de concorrencia entre instalacoes ou recuperacao de falha de energia.
O controlador ainda contem renderizacao de linhas, email, integracoes laser
e outras interacoes. O runtime global e varias areas do ERP continuam em
migracao, conforme `docs/architecture/MODULAR_MONOLITH.md`.
