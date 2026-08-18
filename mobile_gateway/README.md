# LuGEST Mobile Gateway

Ponte autenticada entre a base MySQL do LuGEST e a aplicação LuGEST Field. O
APK nunca recebe credenciais MySQL e o processo fica acessível apenas em
`127.0.0.1`; o acesso exterior é feito por HTTPS através de Cloudflare Tunnel.

## Ligação remota permanente

O `cloudflared` já está instalado neste computador. Depois de o domínio estar na
conta Cloudflare, executar uma única vez:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\setup_cloudflare_tunnel.ps1 -Hostname field.empresa.pt
powershell -ExecutionPolicy Bypass -File .\scripts\install_mobile_remote_autostart.ps1
powershell -ExecutionPolicy Bypass -File .\scripts\verify_mobile_remote.ps1
```

Não é necessário abrir portas no router. O computador tem de estar ligado e com
Internet; o telemóvel pode usar dados móveis ou qualquer Wi-Fi. Para remover o
arranque automático sem tocar no LuGEST desktop:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\remove_mobile_remote_autostart.ps1
```

O modo LAN antigo continua disponível apenas para diagnóstico com
`run_mobile_gateway.ps1 -Mode Lan`. A versão 0.3.0 da app exige HTTPS.

## Endpoints

- `GET /health`
- `GET /v1/stock/summary`
- `GET /v1/stock/products?q=...&in_stock=1`
- `GET /v1/stock/materials?q=...&in_stock=1`
- `POST /v1/service-jobs/sync`
- `POST /v1/service-jobs/attachments`

Todos exigem `Authorization: Bearer <chave>`. As respostas não expõem as
credenciais da base de dados e estão marcadas com `Cache-Control: no-store`.

Cada telemóvel tem uma chave revogável. O registo local guarda apenas o SHA-256
da chave. Fotografias e assinaturas são validadas, limitadas a 10 MB e arquivadas
em `mobile_gateway_data/attachments`, fora da base e dos ficheiros do desktop.

```powershell
.\.venv\Scripts\python.exe -m mobile_gateway.devices list
.\.venv\Scripts\python.exe -m mobile_gateway.devices create --name "Telemóvel João"
.\.venv\Scripts\python.exe -m mobile_gateway.devices revoke ID_DO_DISPOSITIVO
```
