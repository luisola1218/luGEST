import 'package:flutter_test/flutter_test.dart';
import 'package:shared_preferences/shared_preferences.dart';

import 'package:lugest_field/app.dart';

void main() {
  testWidgets('abre o painel de serviços', (WidgetTester tester) async {
    SharedPreferences.setMockInitialValues({});
    await tester.pumpWidget(const LugestFieldApp());
    await tester.pumpAndSettle();

    expect(find.text('Bom trabalho, Luís'), findsOneWidget);
    expect(find.text('Novo serviço'), findsOneWidget);
    expect(find.text('Agenda de hoje'), findsOneWidget);
  });
}
