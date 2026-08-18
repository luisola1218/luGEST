import 'package:flutter/foundation.dart';

import '../data/service_repository.dart';
import '../data/stock_repository.dart';
import '../domain/service_job.dart';

class ServiceStore extends ChangeNotifier {
  ServiceStore(this._repository, {StockRepository? mobileApi})
      : _mobileApi = mobileApi ?? StockRepository();

  final ServiceRepository _repository;
  final StockRepository _mobileApi;
  final List<ServiceJob> _jobs = [];
  bool loading = true;
  bool syncing = false;
  DateTime? lastSyncAt;
  String? syncError;

  List<ServiceJob> get jobs => List.unmodifiable(_jobs);
  int get pendingChanges => _jobs.where((job) => job.needsSync).length;

  List<ServiceJob> get todayJobs {
    final now = DateTime.now();
    return _jobs.where((job) => _sameDay(job.scheduledAt, now)).toList()
      ..sort((a, b) => a.scheduledAt.compareTo(b.scheduledAt));
  }

  List<ServiceJob> get readyToInvoice =>
      _jobs.where((job) => job.status == JobStatus.readyToInvoice).toList();

  double get readyToInvoiceTotal =>
      readyToInvoice.fold(0, (sum, job) => sum + job.total);

  Future<void> load() async {
    loading = true;
    notifyListeners();
    _jobs
      ..clear()
      ..addAll(await _repository.load());
    final syncDates = _jobs.map((job) => job.syncedAt).whereType<DateTime>();
    lastSyncAt = syncDates.isEmpty
        ? null
        : syncDates.reduce((a, b) => a.isAfter(b) ? a : b);
    loading = false;
    notifyListeners();
  }

  ServiceJob? byId(String id) {
    for (final job in _jobs) {
      if (job.id == id) return job;
    }
    return null;
  }

  Future<void> upsert(ServiceJob job) async {
    final index = _jobs.indexWhere((item) => item.id == job.id);
    final updated = job.copyWith(
      updatedAt: DateTime.now(),
      clearSyncedAt: true,
    );
    if (index < 0) {
      _jobs.add(updated);
    } else {
      _jobs[index] = updated;
    }
    notifyListeners();
    await _repository.save(_jobs);
  }

  Future<void> setStatus(String id, JobStatus status) async {
    final job = byId(id);
    if (job != null) await upsert(job.copyWith(status: status));
  }

  Future<void> toggleStep(String id, String step) async {
    final job = byId(id);
    if (job == null) return;
    final steps = [...job.completedSteps];
    steps.contains(step) ? steps.remove(step) : steps.add(step);
    await upsert(job.copyWith(completedSteps: steps));
  }

  Future<void> addPhotos(String id, Iterable<String> paths) async {
    final job = byId(id);
    if (job == null) return;
    final merged = <String>{...job.photoPaths, ...paths}
        .where((path) => path.trim().isNotEmpty)
        .toList();
    await upsert(job.copyWith(photoPaths: merged));
  }

  Future<void> removePhoto(String id, String path) async {
    final job = byId(id);
    if (job == null) return;
    await upsert(job.copyWith(
      photoPaths: job.photoPaths.where((item) => item != path).toList(),
    ));
  }

  Future<void> setSignature(String id, String path) async {
    final job = byId(id);
    if (job == null) return;
    final steps = <String>{...job.completedSteps, 'Assinatura do cliente'};
    await upsert(job.copyWith(
      signaturePath: path,
      signatureAt: DateTime.now(),
      completedSteps: steps.toList(),
    ));
  }

  Future<void> clearSignature(String id) async {
    final job = byId(id);
    if (job == null) return;
    final steps = [...job.completedSteps]..remove('Assinatura do cliente');
    await upsert(job.copyWith(clearSignature: true, completedSteps: steps));
  }

  Future<void> setLines(String id, List<ServiceLine> lines) async {
    final job = byId(id);
    if (job != null) await upsert(job.copyWith(lines: lines));
  }

  Future<void> setVatRate(String id, double vatRate) async {
    final job = byId(id);
    if (job != null) await upsert(job.copyWith(vatRate: vatRate));
  }

  Future<void> sync() async {
    if (syncing) return;
    syncing = true;
    syncError = null;
    notifyListeners();
    try {
      final pending = _jobs.where((job) => job.needsSync).toList();
      if (pending.isEmpty) {
        lastSyncAt ??= DateTime.now();
        return;
      }
      final result = await _mobileApi.syncJobs(pending);
      final syncedAt = DateTime.now();
      for (final job in pending) {
        if (result.syncedIds.contains(job.id) &&
            !result.lockedIds.contains(job.id)) {
          await _mobileApi.uploadJobAttachments(job);
        }
      }
      for (var index = 0; index < _jobs.length; index++) {
        final job = _jobs[index];
        if (result.syncedIds.contains(job.id)) {
          _jobs[index] = job.copyWith(syncedAt: syncedAt);
        }
      }
      await _repository.save(_jobs);
      lastSyncAt = syncedAt;
    } on StockConnectionException catch (exception) {
      syncError = exception.message;
    } finally {
      syncing = false;
      notifyListeners();
    }
  }

  static bool _sameDay(DateTime a, DateTime b) =>
      a.year == b.year && a.month == b.month && a.day == b.day;
}
