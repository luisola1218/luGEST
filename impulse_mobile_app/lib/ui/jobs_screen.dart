import 'package:flutter/material.dart';

import '../core/app_theme.dart';
import '../domain/service_job.dart';
import '../state/service_store.dart';
import 'components.dart';

class JobsScreen extends StatefulWidget {
  const JobsScreen(
      {super.key,
      required this.store,
      required this.onOpenJob,
      required this.onNewJob});

  final ServiceStore store;
  final ValueChanged<ServiceJob> onOpenJob;
  final VoidCallback onNewJob;

  @override
  State<JobsScreen> createState() => _JobsScreenState();
}

class _JobsScreenState extends State<JobsScreen> {
  String query = '';
  JobStatus? status;

  @override
  Widget build(BuildContext context) {
    final normalized = query.trim().toLowerCase();
    final rows = widget.store.jobs.where((job) {
      final matchesText = normalized.isEmpty ||
          '${job.id} ${job.title} ${job.clientName} ${job.address}'
              .toLowerCase()
              .contains(normalized);
      return matchesText && (status == null || job.status == status);
    }).toList()
      ..sort((a, b) => b.scheduledAt.compareTo(a.scheduledAt));
    return Scaffold(
      backgroundColor: Colors.transparent,
      floatingActionButton: FloatingActionButton.extended(
        onPressed: widget.onNewJob,
        backgroundColor: AppColors.lime,
        foregroundColor: AppColors.ink,
        icon: const Icon(Icons.add_rounded),
        label:
            const Text('Novo', style: TextStyle(fontWeight: FontWeight.w900)),
      ),
      body: ListView(
        padding: const EdgeInsets.fromLTRB(18, 12, 18, 110),
        children: [
          Text('Serviços',
              style: Theme.of(context)
                  .textTheme
                  .headlineSmall
                  ?.copyWith(fontWeight: FontWeight.w900)),
          const SizedBox(height: 4),
          const Text('Agenda, execução e faturação num único fluxo.',
              style: TextStyle(color: AppColors.muted)),
          const SizedBox(height: 18),
          TextField(
            onChanged: (value) => setState(() => query = value),
            decoration: const InputDecoration(
                prefixIcon: Icon(Icons.search_rounded),
                hintText: 'Pesquisar cliente, serviço ou local'),
          ),
          const SizedBox(height: 12),
          SizedBox(
            height: 42,
            child: ListView(
              scrollDirection: Axis.horizontal,
              children: [
                _FilterChip(
                    label: 'Todos',
                    selected: status == null,
                    onTap: () => setState(() => status = null)),
                ...[
                  JobStatus.scheduled,
                  JobStatus.inProgress,
                  JobStatus.readyToInvoice,
                  JobStatus.invoiced
                ].map(
                  (value) => _FilterChip(
                      label: value.label,
                      selected: status == value,
                      onTap: () => setState(() => status = value)),
                ),
              ],
            ),
          ),
          const SizedBox(height: 14),
          Text('${rows.length} resultado(s)',
              style: const TextStyle(
                  color: AppColors.muted, fontWeight: FontWeight.w800)),
          const SizedBox(height: 10),
          if (rows.isEmpty)
            const Card(
                child: Padding(
                    padding: EdgeInsets.all(28),
                    child: Center(
                        child:
                            Text('Não existem serviços com estes filtros.'))))
          else
            ...rows.map((job) => Padding(
                  padding: const EdgeInsets.only(bottom: 10),
                  child: JobCard(job: job, onTap: () => widget.onOpenJob(job)),
                )),
        ],
      ),
    );
  }
}

class _FilterChip extends StatelessWidget {
  const _FilterChip(
      {required this.label, required this.selected, required this.onTap});
  final String label;
  final bool selected;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) => Padding(
        padding: const EdgeInsets.only(right: 8),
        child: ChoiceChip(
            label: Text(label), selected: selected, onSelected: (_) => onTap()),
      );
}
