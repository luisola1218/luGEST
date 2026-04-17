import 'package:flutter/material.dart';
import 'package:flutter/services.dart';
import 'package:url_launcher/url_launcher.dart';

Future<void> openPdfUrl(BuildContext context, String url) async {
  final messenger = ScaffoldMessenger.maybeOf(context);
  final uri = Uri.tryParse(url);
  if (uri == null) {
    _showOpenPdfError(messenger, 'URL do PDF invalida.');
    return;
  }

  Object? lastError;
  final modes = <LaunchMode>[
    LaunchMode.inAppBrowserView,
    LaunchMode.platformDefault,
    LaunchMode.externalApplication,
  ];

  for (final mode in modes) {
    try {
      final opened = await launchUrl(uri, mode: mode);
      if (opened) {
        return;
      }
    } catch (exc) {
      lastError = exc;
    }
  }

  final message = lastError == null
      ? 'Nao foi possivel abrir o PDF neste dispositivo.'
      : lastError.toString().replaceFirst('Exception: ', '');
  if (!context.mounted) {
    return;
  }
  _showOpenPdfError(messenger, message);
  await _showPdfFallbackDialog(context, url);
}

void _showOpenPdfError(ScaffoldMessengerState? messenger, String message) {
  messenger?.showSnackBar(
    SnackBar(content: Text(message)),
  );
}

Future<void> _showPdfFallbackDialog(BuildContext context, String url) async {
  if (!context.mounted) {
    return;
  }
  await showDialog<void>(
    context: context,
    builder: (dialogContext) {
      return AlertDialog(
        title: const Text('Abrir PDF'),
        content: Column(
          mainAxisSize: MainAxisSize.min,
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            const Text(
              'Nao foi possivel abrir automaticamente o PDF neste dispositivo.',
            ),
            const SizedBox(height: 12),
            SelectableText(url),
          ],
        ),
        actions: [
          TextButton(
            onPressed: () => Navigator.of(dialogContext).pop(),
            child: const Text('Fechar'),
          ),
          FilledButton(
            onPressed: () async {
              await Clipboard.setData(ClipboardData(text: url));
              if (dialogContext.mounted) {
                Navigator.of(dialogContext).pop();
              }
              if (!context.mounted) {
                return;
              }
              ScaffoldMessenger.maybeOf(context)?.showSnackBar(
                const SnackBar(
                  content: Text('Link do PDF copiado para a area de transferencia.'),
                ),
              );
            },
            child: const Text('Copiar link'),
          ),
        ],
      );
    },
  );
}
