import 'package:flutter_test/flutter_test.dart';

import 'package:lugest_field/domain/service_job.dart';

void main() {
  test('preserva fotografias, assinatura e IVA ao guardar um serviço', () {
    final original = ServiceJob(
      id: 'SV-TEST-1',
      title: 'Manutenção',
      clientName: 'Cliente',
      clientPhone: '',
      address: 'Local',
      scheduledAt: DateTime(2026, 8, 18, 10),
      status: JobStatus.done,
      photoPaths: const ['files/foto_1.jpg', 'files/foto_2.jpg'],
      signaturePath: 'files/assinatura.png',
      signatureAt: DateTime(2026, 8, 18, 11),
      vatRate: 13,
      lines: const [
        ServiceLine(
          description: 'Trabalho',
          quantity: 2,
          unit: 'H',
          unitPrice: 25,
        ),
      ],
    );

    final restored = ServiceJob.fromJson(original.toJson());

    expect(restored.photoPaths, hasLength(2));
    expect(restored.signaturePath, 'files/assinatura.png');
    expect(restored.vatRate, 13);
    expect(restored.subtotal, 50);
    expect(restored.vatAmount, 6.5);
    expect(restored.total, 56.5);
  });

  test('permite remover uma assinatura já guardada', () {
    final job = ServiceJob(
      id: 'SV-TEST-2',
      title: 'Serviço',
      clientName: 'Cliente',
      clientPhone: '',
      address: 'Local',
      scheduledAt: DateTime(2026, 8, 18),
      status: JobStatus.inProgress,
      signaturePath: 'files/assinatura.png',
      signatureAt: DateTime(2026, 8, 18, 11),
    );

    final cleared = job.copyWith(clearSignature: true);
    expect(cleared.signaturePath, isNull);
    expect(cleared.signatureAt, isNull);
  });
}
