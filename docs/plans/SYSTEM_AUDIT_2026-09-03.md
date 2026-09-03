# Auditoria geral do sistema — 2026-09-03

Versão: `2026.09.03.1`

## Resultado

A base técnica foi reorganizada de forma incremental, sem substituir os fluxos
operacionais existentes nem executar alterações na base de dados remota. O
projeto tem agora verificações repetíveis para arquitetura, segurança,
dependências, diagnóstico, licenciamento, schema MySQL, migrações, cálculo laser,
nesting e controlos da interface de orçamentos.

## Trabalho concluído

- Diagnóstico de execução extraído para `lugest_infra/diagnostics`, com escrita
  segura entre threads, rotação de logs e restauro dos hooks globais.
- Fundação de licenciamento separada entre domínio (`lugest_core/licensing`) e
  persistência (`lugest_infra/licensing`). O cliente valida assinaturas com chave
  pública; nenhuma chave privada é distribuída.
- Fingerprint do posto centralizado, preservando compatibilidade com o trial já
  existente.
- Dependências de produção fixadas em `requirements-qt.lock.txt` e verificadas
  antes do build.
- Schema de novas instalações normalizado para `utf8mb4_unicode_ci`.
- Framework de migrações MySQL criado com plano em modo de leitura por defeito,
  checksum imutável, lock e exigência de confirmação explícita de backup.
- Build Windows reduzido de cerca de 671,4 MB/4009 ficheiros para cerca de
  232,2 MB/1229 ficheiros, sem perder os módulos usados. O mapa integrado passa
  a usar o navegador externo quando o WebEngine não está distribuído.
- Recursos e verificações dos controlos numéricos e de operações dos orçamentos
  incluídos no pacote.

## Evidência de validação

- Compilação integral do código Python: aprovada.
- `pip check`: nenhuma dependência quebrada.
- Auditoria de segurança: zero ocorrências altas e médias.
- Fronteiras de arquitetura: `lugest_core` independente de Qt/desktop/infra e
  `lugest_infra` independente de Qt/desktop.
- Assinatura, adulteração, validade, equipamento, módulos, postos e gravação
  atómica de licenças: aprovados.
- Definição canónica do schema: 55 tabelas em `utf8mb4`.
- Auditoria remota apenas de leitura: 13 verificações de integridade sem
  ocorrências e nenhuma tabela fora de InnoDB.
- Build PyInstaller: executável, Qt, criptografia, Shapely e recursos SVG
  confirmados.

## Limites e dívida técnica conhecida

- A base remota ainda tem 52 tabelas com charset legado. A conversão deve ser
  feita apenas após backup testado, janela de manutenção e ensaio numa cópia.
- `lugest_qt/ui/pages/runtime_pages.py` e
  `lugest_qt/services/main_bridge.py` continuam demasiado grandes. A divisão
  deverá ser gradual, por funcionalidade, acompanhada por testes de regressão;
  uma reescrita total agora aumentaria o risco operacional.
- Os testes que gravam em MySQL não foram executados contra o servidor remoto.
  Devem correr numa base descartável antes de cada release.
- O licenciamento comercial ainda não está ligado ao arranque. Faltam decisões
  de edições, módulos, postos/utilizadores, tolerância offline, transferência,
  revogação e período de graça.
- Falta validar o instalador numa máquina Windows limpa, numa rede real de vários
  utilizadores, e assinar digitalmente executável, instalador e atualizações.
- O módulo de faturação não deve ser comercializado como certificado antes da
  certificação formal aplicável da Autoridade Tributária.

## Próxima etapa recomendada

Criar uma base MySQL de homologação isolada, restaurar nela uma cópia anonimizada,
executar a matriz completa de testes mutáveis e ensaiar a conversão de charset.
Em paralelo, fechar a grelha comercial de edições e módulos para integrar o
`LicenseService` sem bloquear backups, exportações ou acesso ao suporte.
