# Rececao e conciliacao de qualidade

Casos de uso migrados: avaliacao, conciliacao, projecao de movimentos,
resumo, integridade e opcoes de ligacao. Rececao prepara produtos, materiais,
notas, NC, devolucoes e eventos antes de uma gravacao bloqueante.

Correcoes demonstradas: rejeicao de excesso/NaN/infinito sem alteracoes;
decisao limitada a linha do movimento; quantidade de devolucao igual a rejeitada;
metadados quantitativos e identificadores preservados apos releitura SQL;
libertacao resolve pendentes nos movimentos e nao duplica stock ao repetir.

Validacao: suite SafeOnly aprovada (425 ficheiros compilados, 487 contratos),
verify_quality_reception_flow aprovado em 27.255 s com rollback e conteudo
inalterado nas 55 tabelas. Teste inclui decisoes parciais, rejeicao, devolucao,
libertacao de material ligada a rececao, conciliacao e consultas apos reload.
A primeira execucao identificou perda de campos no mapeamento SQL; foi corrigida
com metadados do modulo no runtime_state existente, sem DDL.

Pendente global: restantes modulos e adaptadores de persistencia, concorrencia,
bootstrap e instalacao limpa. Relatorios de qualidade continuam historicos.
