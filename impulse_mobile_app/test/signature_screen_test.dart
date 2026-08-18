import 'package:flutter/material.dart';
import 'package:flutter_test/flutter_test.dart';

import 'package:lugest_field/ui/signature_screen.dart';

void main() {
  testWidgets('abre a área e regista o desenho da assinatura', (tester) async {
    await tester.pumpWidget(const MaterialApp(
      home: SignatureScreen(clientName: 'Cliente Teste'),
    ));
    expect(find.text('Assinatura do cliente'), findsOneWidget);

    final canvas = find.byKey(const Key('signature_canvas'));
    await tester.drag(canvas, const Offset(120, 20));
    await tester.drag(canvas, const Offset(-60, 45));
    await tester.pump();

    final clearButton = tester.widget<TextButton>(
      find.widgetWithText(TextButton, 'Limpar'),
    );
    expect(clearButton.onPressed, isNotNull);
    expect(find.text('Guardar assinatura'), findsOneWidget);
  });
}
