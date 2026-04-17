import 'package:flutter/material.dart';

class AppPalette {
  static const ink = Color(0xFF152033);
  static const slate = Color(0xFF66758A);
  static const muted = Color(0xFF98A4B3);
  static const line = Color(0xFFDCE4EF);
  static const lineStrong = Color(0xFFC7D3E3);
  static const surface = Color(0xFFFFFFFF);
  static const surfaceAlt = Color(0xFFF6F8FB);
  static const backgroundTop = Color(0xFFF4F7FB);
  static const backgroundBottom = Color(0xFFEAF0F7);
  static const navy = Color(0xFF183B6B);
  static const cobalt = Color(0xFF2457D6);
  static const green = Color(0xFF16A34A);
  static const amber = Color(0xFFD97706);
  static const orange = Color(0xFFF5A100);
  static const red = Color(0xFFDC2626);
}

class MobileViewSettingsScope extends InheritedWidget {
  const MobileViewSettingsScope({
    super.key,
    required this.uiScale,
    required super.child,
  });

  final double uiScale;

  static MobileViewSettingsScope? maybeOf(BuildContext context) {
    return context.dependOnInheritedWidgetOfExactType<MobileViewSettingsScope>();
  }

  static double uiScaleOf(BuildContext context) {
    return maybeOf(context)?.uiScale ?? 1.0;
  }

  @override
  bool updateShouldNotify(MobileViewSettingsScope oldWidget) {
    return uiScale != oldWidget.uiScale;
  }
}

double mobileUiScale(BuildContext context) {
  final width = MediaQuery.sizeOf(context).width;
  final manual = MobileViewSettingsScope.uiScaleOf(context);
  final adaptive = width >= 900
      ? 1.05
      : width >= 720
          ? 1.02
          : width <= 360
              ? 0.98
              : 1.0;
  return (manual * adaptive).clamp(0.96, 1.08);
}

double ui(BuildContext context, double value, {double? min, double? max}) {
  final scaled = value * mobileUiScale(context);
  if (min == null && max == null) return scaled;
  return scaled.clamp(min ?? scaled, max ?? scaled);
}

bool isWideMobileCanvas(BuildContext context) {
  return MediaQuery.sizeOf(context).width >= 720;
}

int adaptiveGridCount(
  BuildContext context, {
  int compact = 1,
  int wide = 2,
  int ultra = 3,
}) {
  final width = MediaQuery.sizeOf(context).width;
  if (width >= 1120) return ultra;
  if (width >= 720) return wide;
  return compact;
}

ThemeData buildLugestMobileTheme() {
  final base = ThemeData(
    useMaterial3: true,
    colorScheme: ColorScheme.fromSeed(
      seedColor: AppPalette.navy,
      brightness: Brightness.light,
      primary: AppPalette.navy,
      secondary: AppPalette.orange,
      surface: AppPalette.surface,
    ),
  );
  final textTheme = base.textTheme.apply(
    bodyColor: AppPalette.ink,
    displayColor: AppPalette.ink,
  );
  return base.copyWith(
    scaffoldBackgroundColor: AppPalette.backgroundTop,
    textTheme: textTheme.copyWith(
      headlineMedium: textTheme.headlineMedium?.copyWith(
        fontWeight: FontWeight.w800,
        letterSpacing: -0.4,
      ),
      headlineSmall: textTheme.headlineSmall?.copyWith(
        fontWeight: FontWeight.w800,
        letterSpacing: -0.3,
      ),
      titleLarge: textTheme.titleLarge?.copyWith(fontWeight: FontWeight.w800),
      titleMedium: textTheme.titleMedium?.copyWith(fontWeight: FontWeight.w700),
      bodyLarge: textTheme.bodyLarge?.copyWith(height: 1.3),
      bodyMedium: textTheme.bodyMedium?.copyWith(height: 1.3),
      bodySmall: textTheme.bodySmall?.copyWith(height: 1.3),
      labelLarge: textTheme.labelLarge?.copyWith(fontWeight: FontWeight.w700),
      labelMedium: textTheme.labelMedium?.copyWith(fontWeight: FontWeight.w700),
    ),
    appBarTheme: const AppBarTheme(
      elevation: 0,
      scrolledUnderElevation: 0,
      centerTitle: false,
      backgroundColor: Colors.transparent,
      foregroundColor: AppPalette.ink,
    ),
    cardTheme: const CardThemeData(
      color: AppPalette.surface,
      margin: EdgeInsets.zero,
      elevation: 0,
      shape: RoundedRectangleBorder(
        borderRadius: BorderRadius.all(Radius.circular(24)),
        side: BorderSide(color: AppPalette.line, width: 1),
      ),
    ),
    dividerTheme: const DividerThemeData(
      color: AppPalette.line,
      thickness: 1,
      space: 1,
    ),
    inputDecorationTheme: InputDecorationTheme(
      filled: true,
      fillColor: AppPalette.surface,
      contentPadding: const EdgeInsets.symmetric(horizontal: 16, vertical: 14),
      hintStyle: const TextStyle(color: AppPalette.slate),
      border: OutlineInputBorder(
        borderRadius: BorderRadius.circular(18),
        borderSide: const BorderSide(color: AppPalette.line),
      ),
      enabledBorder: OutlineInputBorder(
        borderRadius: BorderRadius.circular(18),
        borderSide: const BorderSide(color: AppPalette.lineStrong),
      ),
      focusedBorder: OutlineInputBorder(
        borderRadius: BorderRadius.circular(18),
        borderSide: const BorderSide(color: AppPalette.navy, width: 1.4),
      ),
    ),
    filledButtonTheme: FilledButtonThemeData(
      style: FilledButton.styleFrom(
        backgroundColor: AppPalette.navy,
        foregroundColor: Colors.white,
        padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 14),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(18)),
      ),
    ),
    outlinedButtonTheme: OutlinedButtonThemeData(
      style: OutlinedButton.styleFrom(
        foregroundColor: AppPalette.navy,
        side: const BorderSide(color: AppPalette.lineStrong),
        padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 14),
        shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(18)),
      ),
    ),
    chipTheme: ChipThemeData(
      backgroundColor: AppPalette.surfaceAlt,
      selectedColor: AppPalette.surfaceAlt,
      disabledColor: AppPalette.surfaceAlt,
      side: const BorderSide(color: AppPalette.line),
      shape: RoundedRectangleBorder(borderRadius: BorderRadius.circular(999)),
      labelStyle: const TextStyle(
        fontWeight: FontWeight.w700,
        color: AppPalette.ink,
      ),
    ),
    textSelectionTheme: const TextSelectionThemeData(
      cursorColor: AppPalette.navy,
      selectionColor: Color(0x22000000),
      selectionHandleColor: AppPalette.navy,
    ),
  );
}

class MobileSurface extends StatelessWidget {
  const MobileSurface({
    super.key,
    required this.child,
    this.padding = const EdgeInsets.fromLTRB(16, 10, 16, 16),
  });

  final Widget child;
  final EdgeInsets padding;

  @override
  Widget build(BuildContext context) {
    return Container(
      decoration: const BoxDecoration(
        gradient: LinearGradient(
          colors: [AppPalette.backgroundTop, AppPalette.backgroundBottom],
          begin: Alignment.topCenter,
          end: Alignment.bottomCenter,
        ),
      ),
      child: SafeArea(
        child: Padding(
          padding: EdgeInsets.fromLTRB(
            ui(context, padding.left),
            ui(context, padding.top),
            ui(context, padding.right),
            ui(context, padding.bottom),
          ),
          child: child,
        ),
      ),
    );
  }
}

class MobilePageHeader extends StatelessWidget {
  const MobilePageHeader({
    super.key,
    required this.title,
    required this.subtitle,
    this.leading,
    this.trailing,
  });

  final String title;
  final String subtitle;
  final Widget? leading;
  final Widget? trailing;

  @override
  Widget build(BuildContext context) {
    return Row(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        if (leading != null) ...[leading!, SizedBox(width: ui(context, 12))],
        Expanded(
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text(title, style: Theme.of(context).textTheme.headlineSmall),
              SizedBox(height: ui(context, 4)),
              Text(
                subtitle,
                style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                      color: AppPalette.slate,
                    ),
              ),
            ],
          ),
        ),
        if (trailing != null) ...[SizedBox(width: ui(context, 12)), trailing!],
      ],
    );
  }
}

class MobileSectionCard extends StatelessWidget {
  const MobileSectionCard({
    super.key,
    this.title,
    this.subtitle,
    this.trailing,
    required this.child,
    this.padding,
  });

  final String? title;
  final String? subtitle;
  final Widget? trailing;
  final Widget child;
  final EdgeInsets? padding;

  @override
  Widget build(BuildContext context) {
    final bodyPadding = padding ??
        EdgeInsets.fromLTRB(
          ui(context, 18),
          ui(context, 18),
          ui(context, 18),
          ui(context, 18),
        );
    return Card(
      child: Padding(
        padding: bodyPadding,
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            if (title != null || subtitle != null || trailing != null)
              Row(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  if (title != null || subtitle != null)
                    Expanded(
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          if (title != null)
                            Text(
                              title!,
                              style: Theme.of(context).textTheme.titleLarge,
                            ),
                          if (subtitle != null) ...[
                            SizedBox(height: ui(context, 4)),
                            Text(
                              subtitle!,
                              style: Theme.of(context)
                                  .textTheme
                                  .bodyMedium
                                  ?.copyWith(color: AppPalette.slate),
                            ),
                          ],
                        ],
                      ),
                    ),
                  if (trailing != null) trailing!,
                ],
              ),
            if (title != null || subtitle != null || trailing != null)
              SizedBox(height: ui(context, 16)),
            child,
          ],
        ),
      ),
    );
  }
}

class MobileHeroCard extends StatelessWidget {
  const MobileHeroCard({
    super.key,
    required this.title,
    required this.subtitle,
    this.actions = const <Widget>[],
    required this.child,
  });

  final String title;
  final String subtitle;
  final List<Widget> actions;
  final Widget child;

  @override
  Widget build(BuildContext context) {
    return Container(
      decoration: BoxDecoration(
        color: AppPalette.navy,
        borderRadius: BorderRadius.circular(28),
      ),
      child: Padding(
        padding: EdgeInsets.all(ui(context, 20)),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Text(
              title,
              style: Theme.of(context).textTheme.headlineSmall?.copyWith(
                    color: Colors.white,
                  ),
            ),
            SizedBox(height: ui(context, 6)),
            Text(
              subtitle,
              style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                    color: Colors.white.withValues(alpha: 0.82),
                  ),
            ),
            if (actions.isNotEmpty) ...[
              SizedBox(height: ui(context, 16)),
              Wrap(spacing: ui(context, 10), runSpacing: ui(context, 10), children: actions),
            ],
            SizedBox(height: ui(context, 18)),
            child,
          ],
        ),
      ),
    );
  }
}

class ModuleActionCard extends StatelessWidget {
  const ModuleActionCard({
    super.key,
    required this.icon,
    required this.color,
    required this.title,
    required this.subtitle,
    this.badge,
    this.onTap,
  });

  final IconData icon;
  final Color color;
  final String title;
  final String subtitle;
  final String? badge;
  final VoidCallback? onTap;

  @override
  Widget build(BuildContext context) {
    return Card(
      child: InkWell(
        borderRadius: BorderRadius.circular(24),
        onTap: onTap,
        child: Padding(
          padding: EdgeInsets.all(ui(context, 18)),
          child: Row(
            children: [
              Container(
                width: ui(context, 50),
                height: ui(context, 50),
                decoration: BoxDecoration(
                  color: color.withValues(alpha: 0.12),
                  borderRadius: BorderRadius.circular(16),
                ),
                child: Icon(icon, color: color),
              ),
              SizedBox(width: ui(context, 14)),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(title, style: Theme.of(context).textTheme.titleLarge),
                    SizedBox(height: ui(context, 4)),
                    Text(
                      subtitle,
                      style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                            color: AppPalette.slate,
                          ),
                    ),
                  ],
                ),
              ),
              if (badge != null) ...[
                SizedBox(width: ui(context, 8)),
                StatusPill(badge!, color: color),
              ],
              SizedBox(width: ui(context, 8)),
              Icon(Icons.chevron_right_rounded, color: AppPalette.slate),
            ],
          ),
        ),
      ),
    );
  }
}

class SoftChip extends StatelessWidget {
  const SoftChip({
    super.key,
    required this.label,
    this.icon,
    this.color = AppPalette.navy,
    this.background = AppPalette.surfaceAlt,
    this.tight = false,
  });

  final String label;
  final IconData? icon;
  final Color color;
  final Color background;
  final bool tight;

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: EdgeInsets.symmetric(
        horizontal: ui(context, tight ? 10 : 12),
        vertical: ui(context, tight ? 7 : 9),
      ),
      decoration: BoxDecoration(
        color: background,
        borderRadius: BorderRadius.circular(999),
        border: Border.all(color: AppPalette.line),
      ),
      child: Row(
        mainAxisSize: MainAxisSize.min,
        children: [
          if (icon != null) ...[
            Icon(icon, size: ui(context, 16), color: color),
            SizedBox(width: ui(context, 6)),
          ],
          Text(
            label,
            style: Theme.of(context).textTheme.labelMedium?.copyWith(
                  color: color,
                ),
          ),
        ],
      ),
    );
  }
}

class StatusPill extends StatelessWidget {
  const StatusPill(this.label, {super.key, this.color});

  final String label;
  final Color? color;

  @override
  Widget build(BuildContext context) {
    final resolved = color ?? statusColor(label);
    return Container(
      padding: EdgeInsets.symmetric(
        horizontal: ui(context, 10),
        vertical: ui(context, 6),
      ),
      decoration: BoxDecoration(
        color: softStatusBackground(label, override: resolved),
        borderRadius: BorderRadius.circular(999),
      ),
      child: Text(
        label,
        style: Theme.of(context).textTheme.labelMedium?.copyWith(
              color: resolved,
            ),
      ),
    );
  }
}

class MetricTile extends StatelessWidget {
  const MetricTile({
    super.key,
    required this.label,
    required this.value,
    this.caption,
    this.color = AppPalette.navy,
  });

  final String label;
  final String value;
  final String? caption;
  final Color color;

  @override
  Widget build(BuildContext context) {
    return Container(
      padding: EdgeInsets.all(ui(context, 16)),
      decoration: BoxDecoration(
        color: AppPalette.surface,
        borderRadius: BorderRadius.circular(20),
        border: Border.all(color: AppPalette.line),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(
            label,
            style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                  color: AppPalette.slate,
                ),
          ),
          SizedBox(height: ui(context, 8)),
          Text(
            value,
            style: Theme.of(context).textTheme.headlineMedium?.copyWith(
                  color: color,
                ),
          ),
          if (caption != null) ...[
            SizedBox(height: ui(context, 6)),
            Text(
              caption!,
              style: Theme.of(context).textTheme.bodySmall?.copyWith(
                    color: AppPalette.slate,
                  ),
            ),
          ],
        ],
      ),
    );
  }
}

class MetricBoardItem {
  const MetricBoardItem({
    required this.label,
    required this.value,
    this.caption,
    this.color = AppPalette.navy,
  });

  final String label;
  final String value;
  final String? caption;
  final Color color;
}

class MetricBoard extends StatelessWidget {
  const MetricBoard({super.key, required this.items});

  final List<MetricBoardItem> items;

  @override
  Widget build(BuildContext context) {
    final columns = adaptiveGridCount(context, compact: 1, wide: 2, ultra: 3);
    return GridView.builder(
      shrinkWrap: true,
      physics: const NeverScrollableScrollPhysics(),
      itemCount: items.length,
      gridDelegate: SliverGridDelegateWithFixedCrossAxisCount(
        crossAxisCount: columns,
        crossAxisSpacing: ui(context, 12),
        mainAxisSpacing: ui(context, 12),
        childAspectRatio: columns == 1 ? 2.6 : 1.45,
      ),
      itemBuilder: (context, index) {
        final item = items[index];
        return MetricTile(
          label: item.label,
          value: item.value,
          caption: item.caption,
          color: item.color,
        );
      },
    );
  }
}

class EmptyStateCard extends StatelessWidget {
  const EmptyStateCard({
    super.key,
    required this.title,
    required this.message,
    this.icon = Icons.inbox_rounded,
  });

  final String title;
  final String message;
  final IconData icon;

  @override
  Widget build(BuildContext context) {
    return Container(
      width: double.infinity,
      padding: EdgeInsets.all(ui(context, 18)),
      decoration: BoxDecoration(
        color: AppPalette.surfaceAlt,
        borderRadius: BorderRadius.circular(20),
        border: Border.all(color: AppPalette.line),
      ),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Icon(icon, color: AppPalette.navy),
          SizedBox(height: ui(context, 10)),
          Text(title, style: Theme.of(context).textTheme.titleMedium),
          SizedBox(height: ui(context, 4)),
          Text(
            message,
            style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                  color: AppPalette.slate,
                ),
          ),
        ],
      ),
    );
  }
}

class DataRowCard extends StatelessWidget {
  const DataRowCard({
    super.key,
    required this.title,
    required this.subtitle,
    this.trailing,
    this.footer,
    this.onTap,
  });

  final String title;
  final String subtitle;
  final Widget? trailing;
  final Widget? footer;
  final VoidCallback? onTap;

  @override
  Widget build(BuildContext context) {
    final body = Padding(
      padding: EdgeInsets.all(ui(context, 16)),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Row(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(title, style: Theme.of(context).textTheme.titleMedium),
                    SizedBox(height: ui(context, 4)),
                    Text(
                      subtitle,
                      style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                            color: AppPalette.slate,
                          ),
                    ),
                  ],
                ),
              ),
              if (trailing != null) ...[SizedBox(width: ui(context, 10)), trailing!],
            ],
          ),
          if (footer != null) ...[
            SizedBox(height: ui(context, 12)),
            footer!,
          ],
        ],
      ),
    );
    return Card(
      child: onTap == null
          ? body
          : InkWell(
              borderRadius: BorderRadius.circular(24),
              onTap: onTap,
              child: body,
            ),
    );
  }
}

class FilterField extends StatelessWidget {
  const FilterField({
    super.key,
    required this.controller,
    required this.hint,
    this.prefixIcon,
    this.onChanged,
  });

  final TextEditingController controller;
  final String hint;
  final IconData? prefixIcon;
  final ValueChanged<String>? onChanged;

  @override
  Widget build(BuildContext context) {
    return TextField(
      controller: controller,
      onChanged: onChanged,
      decoration: InputDecoration(
        hintText: hint,
        prefixIcon: prefixIcon == null ? null : Icon(prefixIcon),
      ),
    );
  }
}

class FilterDropdown extends StatelessWidget {
  const FilterDropdown({
    super.key,
    required this.value,
    required this.items,
    required this.onChanged,
  });

  final String value;
  final List<String> items;
  final ValueChanged<String?> onChanged;

  @override
  Widget build(BuildContext context) {
    return DropdownButtonFormField<String>(
      initialValue:
          items.contains(value) ? value : (items.isEmpty ? null : items.first),
      items: items
          .map((item) => DropdownMenuItem<String>(value: item, child: Text(item)))
          .toList(),
      onChanged: onChanged,
      decoration: const InputDecoration(),
    );
  }
}

Color statusColor(String raw) {
  final txt = raw.toLowerCase();
  if (txt.contains('avaria') || txt.contains('crit')) return AppPalette.red;
  if (txt.contains('curso') || txt.contains('produc')) return AppPalette.cobalt;
  if (txt.contains('pausa') || txt.contains('espera')) return AppPalette.amber;
  if (txt.contains('concl') || txt.contains('ok')) return AppPalette.green;
  if (txt.contains('risco') || txt.contains('atras')) return AppPalette.orange;
  return AppPalette.navy;
}

Color softStatusBackground(String raw, {Color? override}) {
  final color = override ?? statusColor(raw);
  return color.withValues(alpha: 0.12);
}

String shortOperation(String raw) {
  final txt = raw.trim();
  if (txt.isEmpty) return '-';
  return txt
      .replaceAll('Corte Laser', 'Laser')
      .replaceAll('Embalamento', 'Emb.')
      .replaceAll('Quinagem', 'Quina')
      .replaceAll('Soldadura', 'Sold.')
      .replaceAll('Pintura', 'Pint.')
      .replaceAll('Preparacao', 'Prep.');
}
