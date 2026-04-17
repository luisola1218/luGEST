# OTA sem reinstalar APK (Shorebird)

Este guia permite atualizar a app em producao sem instalar APK nova, para alteracoes de Dart/UI.

## 1) Pre-requisitos

- Flutter instalado
- Dart no PATH (vem com Flutter)
- Conta Shorebird

Instalar CLI:

```powershell
dart pub global activate shorebird_cli
shorebird login
```

## 2) Inicializar Shorebird no projeto (1x)

No terminal, dentro de `impulse_mobile_app`:

```powershell
shorebird init
```

Isto cria a configuracao `shorebird.yaml` com o `app_id`.

## 3) Publicar release base (1x por versao base)

Sem release base, nao existe alvo para patches OTA.

```powershell
cd impulse_mobile_app
.\ota_code_push.bat release
```

## 4) Publicar patch OTA (Dart/UI)

Depois de alterares codigo Flutter:

```powershell
cd impulse_mobile_app
.\ota_code_push.bat patch
```

As apps instaladas com essa release base passam a receber o patch sem reinstalacao de APK.

## Limites importantes

- Alteracoes em `android/`, `ios/`, SDK nativo, assinatura, permissao ou plugins nativos exigem nova release.
- Alteracoes apenas Dart/UI/logic sao elegiveis para OTA patch.

## Fluxo recomendado

1. Desenvolvimento diario: `.\arrancar_wifi_hot_reload.bat` (sem reinstalar, hot reload).
2. Entrega interna/QA: APK release normal.
3. Correcao rapida em producao: `.\ota_code_push.bat patch`.
