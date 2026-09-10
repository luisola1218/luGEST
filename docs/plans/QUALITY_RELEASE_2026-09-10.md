# Libertacao de material pela qualidade

Migrada a operacao quality_nc_release_material para MaterialRelease, com
catalogos destacados, sincronizacao de compras sobre copias, auditoria e
movimento de stock preparados antes da publicacao, uma gravacao bloqueante
e recuperacao dos catalogos em erro imediato. Politica de quarentena extraida
para stock_policy; contratos publicos preservados.

Validacao: verify_project.ps1 -SafeOnly passou, 418 ficheiros compilados,
487 contratos; verify_quality_nc_flow passou em 17.764 segundos, rollback
confirmado e conteudo das 55 tabelas inalterado. Testes unitarios cobrem
falha de gravacao, NC inexistente, repeticao sem duplicar stock e snapshot
local desatualizado. O teste SQL inclui material sem rececoes associadas.

Pendente: migrar rececao e conciliacao de movimentos de qualidade; testar
libertacao com rececoes associadas. Persistencia ainda usa o runtime legado;
esta alteracao nao conclui a migracao integral do ERP nem garante atomicidade
entre processos concorrentes ou falhas posteriores a um commit SQL.
