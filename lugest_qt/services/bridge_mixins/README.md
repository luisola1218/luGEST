# Adaptadores do backend

Consultar o [guia de manutencao](../../../docs/architecture/BACKEND_GUIDE.md)
para localizar funcionalidades, dependencias e testes.

Estes componentes sao compostos por `LegacyBackend` e partilham o seu estado.
Nao sao servicos autonomos. Regras novas devem ser extraidas para componentes
com dependencias explicitas; o adaptador conserva a API usada pelas paginas.

Para localizar um metodo sem arrancar a aplicacao:

```powershell
.venv\Scripts\python.exe scripts\backend_map.py nome_do_metodo
```
