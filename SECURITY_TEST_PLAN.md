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
- Definir `LUGEST_API_SECRET` com 16+ caracteres fortes.
- Definir `LUGEST_API_ALLOWED_ORIGINS` apenas se houver frontend browser. Para app mobile nativa pode ficar vazio.
- Guardar `lugest.env`, `impulse_mobile_api\.env` e backups fora da pasta partilhada com utilizadores.

## Testes funcionais de protecao

### Login e perfis

- Tentar entrar com password errada em `desktop`, `mobile` e `API`.
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

### API mobile

- Arrancar a API sem `LUGEST_API_SECRET` e confirmar que o login falha com erro de configuracao.
- Confirmar que tokens expiram e que um token alterado e rejeitado.
- Confirmar que a API nao depende de `CORS` aberto para a app mobile.

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
- Confirmar paths de documentos, PDF, trial e mobile API nesse ambiente.
