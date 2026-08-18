import 'dart:convert';

import 'package:shared_preferences/shared_preferences.dart';

import '../domain/service_job.dart';

class ServiceRepository {
  static const _storageKey = 'lugest_field_jobs_v1';

  Future<List<ServiceJob>> load() async {
    final preferences = await SharedPreferences.getInstance();
    final raw = preferences.getString(_storageKey);
    if (raw == null || raw.isEmpty) {
      final jobs = _demoJobs();
      await save(jobs);
      return jobs;
    }
    try {
      final decoded = jsonDecode(raw) as List<dynamic>;
      return decoded.map((item) {
        final payload = Map<String, dynamic>.from(item as Map);
        if (!payload.containsKey('isDemo') &&
            const {
              'Café Central',
              'Marta Ferreira',
              'Oficina Norte',
              'Clínica Horizonte',
            }.contains(payload['clientName'])) {
          payload['isDemo'] = true;
        }
        return ServiceJob.fromJson(payload);
      }).toList();
    } catch (_) {
      return _demoJobs();
    }
  }

  Future<void> save(List<ServiceJob> jobs) async {
    final preferences = await SharedPreferences.getInstance();
    await preferences.setString(
        _storageKey, jsonEncode(jobs.map((job) => job.toJson()).toList()));
  }

  List<ServiceJob> _demoJobs() {
    final now = DateTime.now();
    DateTime at(int dayOffset, int hour, int minute) =>
        DateTime(now.year, now.month, now.day + dayOffset, hour, minute);
    return [
      ServiceJob(
        id: 'SV-${now.year}-0042',
        title: 'Reparação urgente',
        clientName: 'Café Central',
        clientPhone: '912 345 678',
        address: 'Rua do Mercado 18, Braga',
        scheduledAt: at(0, 9, 30),
        status: JobStatus.inProgress,
        isDemo: true,
        notes:
            'Diagnosticar a origem do problema antes de substituir material.',
        completedSteps: const ['Chegada confirmada', 'Diagnóstico inicial'],
        lines: const [
          ServiceLine(
              description: 'Deslocação técnica',
              quantity: 1,
              unit: 'SV',
              unitPrice: 28),
          ServiceLine(
              description: 'Mão de obra especializada',
              quantity: 1.5,
              unit: 'H',
              unitPrice: 32),
        ],
      ),
      ServiceJob(
        id: 'SV-${now.year}-0043',
        title: 'Montagem no local',
        clientName: 'Marta Ferreira',
        clientPhone: '934 222 101',
        address: 'Av. da Liberdade 204, Guimarães',
        scheduledAt: at(0, 14, 0),
        status: JobStatus.scheduled,
        isDemo: true,
        notes: 'Cliente pediu confirmação 30 minutos antes.',
        lines: const [
          ServiceLine(
              description: 'Material de montagem',
              quantity: 3,
              unit: 'UN',
              unitPrice: 14.5),
          ServiceLine(
              description: 'Montagem e verificação',
              quantity: 2,
              unit: 'H',
              unitPrice: 32),
        ],
      ),
      ServiceJob(
        id: 'SV-${now.year}-0039',
        title: 'Manutenção programada',
        clientName: 'Oficina Norte',
        clientPhone: '253 555 470',
        address: 'Zona Industrial, Lote 7',
        scheduledAt: at(-1, 16, 30),
        status: JobStatus.readyToInvoice,
        isDemo: true,
        notes: 'Trabalho validado pelo responsável da oficina.',
        completedSteps: const [
          'Chegada confirmada',
          'Diagnóstico inicial',
          'Trabalho executado',
          'Assinatura do cliente'
        ],
        lines: const [
          ServiceLine(
              description: 'Material aplicado',
              quantity: 2,
              unit: 'UN',
              unitPrice: 67),
          ServiceLine(
              description: 'Mão de obra',
              quantity: 2.5,
              unit: 'H',
              unitPrice: 32),
        ],
      ),
      ServiceJob(
        id: 'SV-${now.year}-0044',
        title: 'Visita técnica preventiva',
        clientName: 'Clínica Horizonte',
        clientPhone: '253 118 822',
        address: 'Praça Nova 5, Braga',
        scheduledAt: at(1, 10, 0),
        status: JobStatus.scheduled,
        isDemo: true,
        lines: const [
          ServiceLine(
              description: 'Inspeção e relatório',
              quantity: 1,
              unit: 'SV',
              unitPrice: 95)
        ],
      ),
    ];
  }
}
