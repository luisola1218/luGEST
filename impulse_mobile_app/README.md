# LuGEST Field

Aplicação Flutter para trabalhadores independentes e equipas de serviços no
terreno: manutenção, reparação, montagem, assistência, limpeza e outras áreas.

## O que já está implementado

- painel diário com agenda, serviços ativos e valor pronto a faturar;
- fluxo `Agendado -> A caminho -> Em execução -> Concluído -> Por faturar`;
- criação rápida de serviço com o campo de data e hora clicável em toda a área;
- ficha do serviço com cliente, contacto, rota, checklist, materiais e totais;
- fotografias reais da câmara/galeria guardadas por serviço;
- assinatura desenhada no telemóvel e associada ao cliente;
- trabalhos, horas e materiais adicionados à conta;
- resumo final com IVA configurável para apresentar ao cliente;
- stock real de produtos e matérias-primas consultado por API autenticada;
- seleção de artigos do LuGEST para acrescentar à conta;
- sincronização real dos serviços para `servicos_diretos` como rascunhos;
- envio idempotente de fotografias e assinaturas para o arquivo privado no computador;
- pesquisa e filtros de serviços;
- histórico simples por cliente;
- persistência local com `SharedPreferences`;
- fila visual de alterações locais e ação de sincronização.

## Arranque e compilação

O Flutter SDK está instalado neste posto em
`C:\Users\engenharia\develop\flutter`.

```powershell
cd impulse_mobile_app
& 'C:\Users\engenharia\develop\flutter\bin\flutter.bat' pub get
& 'C:\Users\engenharia\develop\flutter\bin\flutter.bat' test
& 'C:\Users\engenharia\develop\flutter\bin\flutter.bat' build apk --release
```

O projeto Android nativo está em `android/`. O APK de teste validado encontra-se
em `output/mobile/LuGEST-Field-0.3.0-remoto-assinado.apk`. A build usa a chave release
própria do LuGEST Field; a chave e as passwords ficam fora do Git e devem ser
guardadas para permitir futuras atualizações da mesma aplicação.

## Ligação remota ao LuGEST

Preparar uma vez o túnel HTTPS e o arranque automático:

```powershell
powershell -ExecutionPolicy Bypass -File .\scripts\setup_cloudflare_tunnel.ps1 -Hostname field.empresa.pt
powershell -ExecutionPolicy Bypass -File .\scripts\install_mobile_remote_autostart.ps1
```

Na app, abrir `Mais > Stock LuGEST`, tocar no ícone de ligação e introduzir o
endereço HTTPS e a chave do dispositivo. O telemóvel pode estar fora da empresa;
o computador precisa apenas de estar ligado e com Internet.

O resumo apresentado ao cliente não substitui uma fatura fiscal; a emissão final
continua no menu Faturação do LuGEST.

## Arquitetura de produção

A app nunca liga diretamente ao MySQL. Usa uma API HTTPS, autenticação por
dispositivo, limite de pedidos e operações idempotentes. A cache local permite
trabalhar sem rede; quando a ligação regressa, envia os serviços e anexos
pendentes. Documentos já confirmados ou faturados no desktop nunca são
sobrescritos pelo telemóvel.

Consultar também `docs/plans/MOBILE_SERVICES_ARCHITECTURE.md`.
