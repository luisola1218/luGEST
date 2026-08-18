import 'package:flutter/material.dart';

abstract final class AppColors {
  static const ink = Color(0xFF222622);
  static const muted = Color(0xFF697069);
  static const canvas = Color(0xFFF4F5F2);
  static const surface = Colors.white;
  static const lime = Color(0xFF78C91C);
  static const limeDark = Color(0xFF315D12);
  static const limeSoft = Color(0xFFEDF8E3);
  static const blue = Color(0xFF315F8A);
  static const blueSoft = Color(0xFFEAF2FA);
  static const amber = Color(0xFFD18A1F);
  static const amberSoft = Color(0xFFFFF6E6);
  static const red = Color(0xFFD92D20);
  static const redSoft = Color(0xFFFFEFED);
  static const border = Color(0xFFD9DDD7);
}

ThemeData buildAppTheme() {
  final scheme = ColorScheme.fromSeed(
    seedColor: AppColors.lime,
    brightness: Brightness.light,
    primary: AppColors.ink,
    secondary: AppColors.lime,
    surface: AppColors.surface,
    error: AppColors.red,
  );
  return ThemeData(
    useMaterial3: true,
    colorScheme: scheme,
    scaffoldBackgroundColor: AppColors.canvas,
    fontFamily: 'sans-serif',
    appBarTheme: const AppBarTheme(
      backgroundColor: AppColors.canvas,
      foregroundColor: AppColors.ink,
      elevation: 0,
      centerTitle: false,
      titleTextStyle: TextStyle(
          fontSize: 22, fontWeight: FontWeight.w900, color: AppColors.ink),
    ),
    cardTheme: CardThemeData(
      elevation: 0,
      margin: EdgeInsets.zero,
      color: AppColors.surface,
      shape: RoundedRectangleBorder(
        side: const BorderSide(color: AppColors.border),
        borderRadius: BorderRadius.circular(20),
      ),
    ),
    inputDecorationTheme: InputDecorationTheme(
      filled: true,
      fillColor: AppColors.surface,
      contentPadding: const EdgeInsets.symmetric(horizontal: 16, vertical: 15),
      border: OutlineInputBorder(
          borderRadius: BorderRadius.circular(14),
          borderSide: const BorderSide(color: AppColors.border)),
      enabledBorder: OutlineInputBorder(
          borderRadius: BorderRadius.circular(14),
          borderSide: const BorderSide(color: AppColors.border)),
      focusedBorder: OutlineInputBorder(
          borderRadius: BorderRadius.circular(14),
          borderSide: const BorderSide(color: AppColors.lime, width: 2)),
    ),
    filledButtonTheme: FilledButtonThemeData(
      style: FilledButton.styleFrom(
        minimumSize: const Size(0, 52),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(16)),
        textStyle: const TextStyle(fontWeight: FontWeight.w800),
      ),
    ),
    navigationBarTheme: NavigationBarThemeData(
      height: 72,
      backgroundColor: AppColors.surface,
      indicatorColor: AppColors.limeSoft,
      labelTextStyle: WidgetStateProperty.resolveWith(
        (states) => TextStyle(
          fontSize: 11,
          fontWeight: states.contains(WidgetState.selected)
              ? FontWeight.w900
              : FontWeight.w700,
          color: states.contains(WidgetState.selected)
              ? AppColors.limeDark
              : AppColors.muted,
        ),
      ),
    ),
  );
}
