import 'dart:typed_data';
import 'dart:ui' as ui;

import 'package:flutter/material.dart';
import 'package:flutter/rendering.dart';

import '../core/app_theme.dart';

class SignatureScreen extends StatefulWidget {
  const SignatureScreen({super.key, required this.clientName});

  final String clientName;

  @override
  State<SignatureScreen> createState() => _SignatureScreenState();
}

class _SignatureScreenState extends State<SignatureScreen> {
  final boundaryKey = GlobalKey();
  final points = <Offset?>[];
  bool saving = false;

  Future<void> save() async {
    if (points.whereType<Offset>().length < 4) {
      ScaffoldMessenger.of(context).showSnackBar(
        const SnackBar(content: Text('Peça ao cliente para assinar primeiro.')),
      );
      return;
    }
    setState(() => saving = true);
    try {
      final boundary = boundaryKey.currentContext!.findRenderObject()
          as RenderRepaintBoundary;
      final image = await boundary.toImage(pixelRatio: 2.5);
      final data = await image.toByteData(format: ui.ImageByteFormat.png);
      if (!mounted) return;
      Navigator.pop<Uint8List>(context, data!.buffer.asUint8List());
    } finally {
      if (mounted) setState(() => saving = false);
    }
  }

  @override
  Widget build(BuildContext context) => Scaffold(
        appBar: AppBar(
          title: const Text('Assinatura do cliente'),
          actions: [
            TextButton(
              onPressed: points.isEmpty ? null : () => setState(points.clear),
              child: const Text('Limpar'),
            ),
          ],
        ),
        bottomNavigationBar: SafeArea(
          child: Padding(
            padding: const EdgeInsets.fromLTRB(18, 10, 18, 12),
            child: FilledButton.icon(
              onPressed: saving ? null : save,
              style: FilledButton.styleFrom(
                backgroundColor: AppColors.lime,
                foregroundColor: AppColors.ink,
              ),
              icon: saving
                  ? const SizedBox(
                      width: 18,
                      height: 18,
                      child: CircularProgressIndicator(strokeWidth: 2),
                    )
                  : const Icon(Icons.check_rounded),
              label: const Text('Guardar assinatura'),
            ),
          ),
        ),
        body: Padding(
          padding: const EdgeInsets.fromLTRB(18, 8, 18, 18),
          child: Column(
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text(
                'Confirmação de ${widget.clientName}',
                style:
                    const TextStyle(fontSize: 18, fontWeight: FontWeight.w900),
              ),
              const SizedBox(height: 5),
              const Text(
                'Assine dentro da área abaixo. A assinatura fica associada ao resumo deste serviço.',
                style: TextStyle(color: AppColors.muted, height: 1.4),
              ),
              const SizedBox(height: 14),
              Expanded(
                child: ClipRRect(
                  borderRadius: BorderRadius.circular(20),
                  child: RepaintBoundary(
                    key: boundaryKey,
                    child: ColoredBox(
                      color: Colors.white,
                      child: GestureDetector(
                        key: const Key('signature_canvas'),
                        behavior: HitTestBehavior.opaque,
                        onPanStart: (details) => setState(() {
                          points.add(details.localPosition);
                        }),
                        onPanUpdate: (details) => setState(() {
                          points.add(details.localPosition);
                        }),
                        onPanEnd: (_) => setState(() => points.add(null)),
                        child: CustomPaint(
                          painter: _SignaturePainter(points),
                          child: const SizedBox.expand(),
                        ),
                      ),
                    ),
                  ),
                ),
              ),
              const SizedBox(height: 8),
              const Center(
                child: Text('Assinatura',
                    style: TextStyle(color: AppColors.muted, fontSize: 12)),
              ),
            ],
          ),
        ),
      );
}

class _SignaturePainter extends CustomPainter {
  const _SignaturePainter(this.points);
  final List<Offset?> points;

  @override
  void paint(Canvas canvas, Size size) {
    final guide = Paint()
      ..color = const Color(0xFFD9DDD7)
      ..strokeWidth = 1;
    canvas.drawLine(Offset(24, size.height - 44),
        Offset(size.width - 24, size.height - 44), guide);

    final ink = Paint()
      ..color = const Color(0xFF172019)
      ..strokeWidth = 3
      ..strokeCap = StrokeCap.round
      ..strokeJoin = StrokeJoin.round;
    for (var index = 0; index < points.length - 1; index++) {
      final current = points[index];
      final next = points[index + 1];
      if (current != null && next != null) canvas.drawLine(current, next, ink);
    }
  }

  @override
  bool shouldRepaint(covariant _SignaturePainter oldDelegate) => true;
}
