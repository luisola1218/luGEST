import 'package:flutter/material.dart';

import 'src/screens/login_screen.dart';
import 'src/widgets/mobile_ui.dart';

void main() {
  runApp(const LugestImpulseApp());
}

class LugestImpulseApp extends StatelessWidget {
  const LugestImpulseApp({super.key});

  @override
  Widget build(BuildContext context) {
    return MaterialApp(
      title: 'LUGEST Impulse',
      debugShowCheckedModeBanner: false,
      theme: buildLugestMobileTheme(),
      builder: (context, child) {
        final media = MediaQuery.of(context);
        return SelectionContainer.disabled(
          child: MediaQuery(
            data: media.copyWith(
              textScaler: media.textScaler.clamp(
                minScaleFactor: 0.96,
                maxScaleFactor: 1.04,
              ),
              boldText: false,
            ),
            child: child ?? const SizedBox.shrink(),
          ),
        );
      },
      home: const LoginScreen(),
    );
  }
}
