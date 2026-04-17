# Lugest Impulse Mobile App

Base Flutter separada para transformar o `Impulse` numa APK Android.

Esta pasta nao altera o ERP desktop. Consome a API em `../impulse_mobile_api`.

## Estado

- Estrutura base criada manualmente
- Login por token
- Dashboard inicial do Pulse
- Lista de encomendas em preparacao para evoluir depois
- Tema visual industrial ja alinhado com o desktop Qt
- Onboarding de login orientado a instalacao real, sem credenciais demo inseguras

## Requisitos

Instalar Flutter no posto de desenvolvimento.

## Criar o projeto real

Quando o Flutter estiver instalado, podes:

```powershell
cd impulse_mobile_app
flutter pub get
flutter run
```

Para APK:

```powershell
flutter build apk --release
```

## Desenvolvimento sem reinstalar APK

Para nao ter de instalar APK nova a cada alteracao:

1. Liga o telemovel por USB (ou `adb tcpip` para wireless).
2. Arranca em debug:

```powershell
cd impulse_mobile_app
.\arrancar_hot_reload.bat
```

3. No VS Code, guarda os ficheiros (`Ctrl+S`): o app faz hot reload automaticamente.

Notas:
- Isto funciona em modo `debug` (`flutter run`), nao em APK `release`.
- Sempre que quiseres testar versao final, geras nova APK release.

Sem cabo (mesma rede Wi-Fi):

```powershell
cd impulse_mobile_app
.\arrancar_wifi_hot_reload.bat
```

## OTA em producao (sem nova APK)

Foi preparado fluxo de `code push` com Shorebird para alteracoes Dart/UI:

```powershell
cd impulse_mobile_app
.\ota_code_push.bat release   # primeira release base
.\ota_code_push.bat patch     # patches OTA seguintes
```

Guia completo:

- `OTA_SHOREBIRD.md`

## URL da API

No uso normal, o utilizador indica o servidor diretamente no ecra de login.

Importante:
- no telemovel nao usar `127.0.0.1`
- usar o IP ou DNS real do servidor onde a `Mobile API` estiver a correr
- entrar com um utilizador real do LuGEST, criado no desktop

Se quiseres gerar uma APK ja preconfigurada para um cliente especifico:

```powershell
flutter build apk --release --dart-define=LUGEST_DEFAULT_API_HOST=http://IP_DO_SERVIDOR:8050
```

## Proximo passo natural

- detalhe de encomenda
- historico de operacoes
- notificacoes de avaria/desvio
- abrir desenho tecnico da referencia
