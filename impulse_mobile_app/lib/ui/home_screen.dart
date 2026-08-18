import 'package:flutter/material.dart';

import '../core/app_theme.dart';
import '../core/formatters.dart';
import '../domain/service_job.dart';
import '../state/service_store.dart';
import 'components.dart';

class HomeScreen extends StatelessWidget {
  const HomeScreen(
      {super.key,
      required this.store,
      required this.onOpenJob,
      required this.onNewJob,
      required this.onShowJobs});

  final ServiceStore store;
  final ValueChanged<ServiceJob> onOpenJob;
  final VoidCallback onNewJob;
  final VoidCallback onShowJobs;

  @override
  Widget build(BuildContext context) {
    final now = DateTime.now();
    final active =
        store.todayJobs.where((job) => job.status != JobStatus.invoiced).length;
    return RefreshIndicator(
      onRefresh: store.sync,
      color: AppColors.lime,
      child: ListView(
        physics: const AlwaysScrollableScrollPhysics(),
        padding: const EdgeInsets.fromLTRB(18, 12, 18, 110),
        children: [
          Row(children: [
            Expanded(
              child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text('Bom trabalho, Luís',
                        style: Theme.of(context)
                            .textTheme
                            .headlineSmall
                            ?.copyWith(fontWeight: FontWeight.w900)),
                    const SizedBox(height: 3),
                    Text(fullDate(now),
                        style: const TextStyle(
                            color: AppColors.muted,
                            fontWeight: FontWeight.w700)),
                  ]),
            ),
            SyncButton(store: store),
          ]),
          const SizedBox(height: 20),
          Container(
            padding: const EdgeInsets.all(20),
            decoration: BoxDecoration(
              color: AppColors.ink,
              borderRadius: BorderRadius.circular(24),
              boxShadow: const [
                BoxShadow(
                    color: Color(0x22000000),
                    blurRadius: 20,
                    offset: Offset(0, 10))
              ],
            ),
            child:
                Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
              const Row(children: [
                Icon(Icons.bolt_rounded, color: AppColors.lime),
                SizedBox(width: 8),
                Text('O teu dia no terreno',
                    style: TextStyle(
                        color: Colors.white,
                        fontSize: 18,
                        fontWeight: FontWeight.w900)),
              ]),
              const SizedBox(height: 8),
              Text(
                '$active serviços ativos · ${store.readyToInvoice.length} prontos para faturar',
                style: const TextStyle(color: Color(0xFFD6DBD4), height: 1.4),
              ),
              const SizedBox(height: 18),
              SizedBox(
                width: double.infinity,
                child: FilledButton.icon(
                  onPressed: onNewJob,
                  style: FilledButton.styleFrom(
                      backgroundColor: AppColors.lime,
                      foregroundColor: AppColors.ink),
                  icon: const Icon(Icons.add_rounded),
                  label: const Text('Novo serviço'),
                ),
              ),
            ]),
          ),
          const SizedBox(height: 20),
          Row(children: [
            Expanded(
                child: MetricCard(
                    icon: Icons.calendar_today_rounded,
                    value: '${store.todayJobs.length}',
                    label: 'Hoje',
                    tone: AppColors.blue)),
            const SizedBox(width: 10),
            Expanded(
                child: MetricCard(
                    icon: Icons.receipt_long_rounded,
                    value: '${store.readyToInvoice.length}',
                    label: 'Por faturar',
                    tone: AppColors.amber)),
            const SizedBox(width: 10),
            Expanded(
                child: MetricCard(
                    icon: Icons.euro_rounded,
                    value: money(store.readyToInvoiceTotal),
                    label: 'A emitir',
                    tone: AppColors.limeDark)),
          ]),
          const SizedBox(height: 22),
          SectionHeader('Agenda de hoje',
              action: 'Ver tudo', onAction: onShowJobs),
          const SizedBox(height: 10),
          if (store.todayJobs.isEmpty)
            const _EmptyToday()
          else
            ...store.todayJobs.map((job) => Padding(
                  padding: const EdgeInsets.only(bottom: 10),
                  child: JobCard(
                      job: job, onTap: () => onOpenJob(job), compact: true),
                )),
          if (store.readyToInvoice.isNotEmpty) ...[
            const SizedBox(height: 14),
            const SectionHeader('Fechar o dia'),
            const SizedBox(height: 10),
            Card(
              color: AppColors.amberSoft,
              child: Padding(
                padding: const EdgeInsets.all(18),
                child: Row(children: [
                  const CircleAvatar(
                      backgroundColor: Colors.white,
                      child:
                          Icon(Icons.receipt_rounded, color: AppColors.amber)),
                  const SizedBox(width: 13),
                  Expanded(
                      child: Text(
                          '${store.readyToInvoice.length} serviço(s) concluído(s) aguardam faturação.',
                          style: const TextStyle(fontWeight: FontWeight.w800))),
                  const Icon(Icons.arrow_forward_rounded),
                ]),
              ),
            ),
          ],
        ],
      ),
    );
  }
}

class SyncButton extends StatelessWidget {
  const SyncButton({super.key, required this.store});
  final ServiceStore store;

  @override
  Widget build(BuildContext context) => Tooltip(
        message: store.pendingChanges == 0
            ? (store.syncError ?? 'Tudo sincronizado')
            : '${store.pendingChanges} alterações locais',
        child: InkWell(
          onTap: () => _sync(context),
          borderRadius: BorderRadius.circular(14),
          child: Container(
            padding: const EdgeInsets.symmetric(horizontal: 12, vertical: 9),
            decoration: BoxDecoration(
                color: store.pendingChanges == 0
                    ? AppColors.limeSoft
                    : AppColors.amberSoft,
                borderRadius: BorderRadius.circular(14)),
            child: Row(children: [
              if (store.syncing)
                const SizedBox(
                    width: 16,
                    height: 16,
                    child: CircularProgressIndicator(strokeWidth: 2))
              else
                Icon(
                    store.pendingChanges == 0
                        ? Icons.cloud_done_rounded
                        : Icons.cloud_upload_rounded,
                    size: 18,
                    color: store.pendingChanges == 0
                        ? AppColors.limeDark
                        : AppColors.amber),
              const SizedBox(width: 6),
              Text(
                  store.pendingChanges == 0
                      ? 'Online'
                      : '${store.pendingChanges}',
                  style: const TextStyle(
                      fontWeight: FontWeight.w900, fontSize: 12)),
            ]),
          ),
        ),
      );

  Future<void> _sync(BuildContext context) async {
    await store.sync();
    if (!context.mounted) return;
    ScaffoldMessenger.of(context).showSnackBar(SnackBar(
      content: Text(store.syncError ?? 'Serviços sincronizados com o LuGEST.'),
      backgroundColor: store.syncError == null ? AppColors.ink : AppColors.red,
    ));
  }
}

class _EmptyToday extends StatelessWidget {
  const _EmptyToday();

  @override
  Widget build(BuildContext context) => const Card(
        child: Padding(
          padding: EdgeInsets.all(24),
          child: Column(children: [
            Icon(Icons.event_available_rounded,
                color: AppColors.lime, size: 38),
            SizedBox(height: 10),
            Text('Sem serviços marcados para hoje',
                style: TextStyle(fontWeight: FontWeight.w900)),
          ]),
        ),
      );
}
