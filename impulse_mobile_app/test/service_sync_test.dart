import 'package:flutter_test/flutter_test.dart';
import 'package:shared_preferences/shared_preferences.dart';

import 'package:lugest_field/data/service_repository.dart';
import 'package:lugest_field/data/stock_repository.dart';
import 'package:lugest_field/domain/service_job.dart';
import 'package:lugest_field/state/service_store.dart';

class _RecordingMobileApi extends StockRepository {
  List<ServiceJob> received = const [];
  final List<String> attachmentsFor = [];

  @override
  Future<JobSyncResult> syncJobs(List<ServiceJob> jobs) async {
    received = [...jobs];
    return JobSyncResult(
      syncedIds: jobs.map((job) => job.id).toSet(),
      lockedIds: const {},
    );
  }

  @override
  Future<int> uploadJobAttachments(ServiceJob job) async {
    attachmentsFor.add(job.id);
    return job.photoPaths.length + (job.signaturePath == null ? 0 : 1);
  }
}

void main() {
  test('não envia dados demonstrativos e sincroniza serviços reais', () async {
    SharedPreferences.setMockInitialValues({});
    final api = _RecordingMobileApi();
    final store = ServiceStore(ServiceRepository(), mobileApi: api);
    await store.load();

    await store.sync();
    expect(api.received, isEmpty);

    await store.upsert(ServiceJob(
      id: 'SV-REAL-1',
      title: 'Trabalho real',
      clientName: 'Cliente real',
      clientPhone: '',
      address: 'Local',
      scheduledAt: DateTime(2026, 8, 18),
      status: JobStatus.readyToInvoice,
      lines: const [
        ServiceLine(
          description: 'Serviço',
          quantity: 1,
          unit: 'SV',
          unitPrice: 50,
        ),
      ],
    ));
    await store.sync();

    expect(api.received.map((job) => job.id), ['SV-REAL-1']);
    expect(api.attachmentsFor, ['SV-REAL-1']);
    expect(store.pendingChanges, 0);
    expect(store.syncError, isNull);
  });
}
