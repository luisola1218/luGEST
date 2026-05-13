# Plano de Testes de Protecao

## Antes de testar

1. Correr a auditoria:

```powershell
py scripts\security_audit.py
```

2. Simular a limpeza do workspace:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\cleanup_workspace.ps1 -DryRun
```

3. Executar a limpeza real do lixo regeneravel:

```powershell
powershell -ExecutionPolicy Bypass -File scripts\cleanup_workspace.ps1
```

## Fechos minimos antes de entregar

- Alterar todos os utilizadores seed com passwords previsiveis.
  Podes preparar essa rotacao com:

```powershell
py scripts\harden_local_users.py
```

  E aplicar quando decidires:

```powershell
py scripts\harden_local_users.py --write
```

- Definir uma password forte de supervisor no `Extras`.
- Confirmar `LUGEST_OWNER_USERNAME` e `LUGEST_OWNER_PASSWORD` diferentes das contas normais.
- Guardar `lugest.env` e backups fora da pasta partilhada com utilizadores.

## Testes funcionais de protecao

### Login e perfis

- Tentar entrar com password errada no desktop.
- Confirmar que um utilizador sem permissao nao abre menus bloqueados.
- Confirmar que um utilizador nao-admin nao entra em configuracoes criticas.

### Trial e licenciamento

- Ativar um trial curto numa base de testes.
- Simular expirar o trial e verificar bloqueio no arranque.
- Validar que so a conta proprietaria volta a desbloquear.

### Operador e supervisor

- Tentar `Dar Baixa` sem password de supervisor.
- Tentar com password errada.
- Tentar com password certa e confirmar que o registo fica auditavel.

### Base de dados e integridade

- Correr:

```powershell
py scripts\verify_mysql_schema.py
py scripts\verify_data_integrity.py
```

- Validar backup e restauro numa base de testes, nunca em producao.

## Testes de entrega

- Instalar a app noutro PC.
- Confirmar que `py main.py` e `main.exe` arrancam corretamente.
- Confirmar paths de documentos, PDF e trial nesse ambiente.
