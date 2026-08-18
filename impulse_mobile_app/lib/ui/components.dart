import 'package:flutter/material.dart';

import '../core/app_theme.dart';
import '../core/formatters.dart';
import '../domain/service_job.dart';

class SectionHeader extends StatelessWidget {
  const SectionHeader(this.title, {super.key, this.action, this.onAction});

  final String title;
  final String? action;
  final VoidCallback? onAction;

  @override
  Widget build(BuildContext context) => Row(
        children: [
          Expanded(
              child: Text(title,
                  style: const TextStyle(
                      fontSize: 18, fontWeight: FontWeight.w900))),
          if (action != null)
            TextButton(onPressed: onAction, child: Text(action!)),
        ],
      );
}

class MetricCard extends StatelessWidget {
  const MetricCard(
      {super.key,
      required this.icon,
      required this.value,
      required this.label,
      this.tone = AppColors.lime});

  final IconData icon;
  final String value;
  final String label;
  final Color tone;

  @override
  Widget build(BuildContext context) => Card(
        child: Padding(
          padding: const EdgeInsets.all(15),
          child:
              Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
            Container(
              width: 36,
              height: 36,
              decoration: BoxDecoration(
                  color: tone.withValues(alpha: .13),
                  borderRadius: BorderRadius.circular(11)),
              child: Icon(icon, color: tone, size: 20),
            ),
            const SizedBox(height: 13),
            Text(value,
                style: const TextStyle(
                    fontSize: 21,
                    fontWeight: FontWeight.w900,
                    color: AppColors.ink)),
            const SizedBox(height: 2),
            Text(label,
                style: const TextStyle(
                    fontSize: 12,
                    fontWeight: FontWeight.w700,
                    color: AppColors.muted)),
          ]),
        ),
      );
}

class StatusPill extends StatelessWidget {
  const StatusPill(this.status, {super.key});

  final JobStatus status;

  @override
  Widget build(BuildContext context) {
    final (background, foreground) = switch (status) {
      JobStatus.scheduled => (AppColors.blueSoft, AppColors.blue),
      JobStatus.onRoute => (AppColors.amberSoft, AppColors.amber),
      JobStatus.inProgress => (AppColors.limeSoft, AppColors.limeDark),
      JobStatus.done => (AppColors.limeSoft, AppColors.limeDark),
      JobStatus.readyToInvoice => (
          AppColors.amberSoft,
          const Color(0xFF87550D)
        ),
      JobStatus.invoiced => (const Color(0xFFEDF0ED), AppColors.muted),
    };
    return Container(
      padding: const EdgeInsets.symmetric(horizontal: 10, vertical: 6),
      decoration: BoxDecoration(
          color: background, borderRadius: BorderRadius.circular(20)),
      child: Text(status.label,
          style: TextStyle(
              fontSize: 11, fontWeight: FontWeight.w900, color: foreground)),
    );
  }
}

class JobCard extends StatelessWidget {
  const JobCard(
      {super.key,
      required this.job,
      required this.onTap,
      this.compact = false});

  final ServiceJob job;
  final VoidCallback onTap;
  final bool compact;

  @override
  Widget build(BuildContext context) => Card(
        child: InkWell(
          onTap: onTap,
          borderRadius: BorderRadius.circular(20),
          child: Padding(
            padding: EdgeInsets.all(compact ? 14 : 17),
            child: Row(crossAxisAlignment: CrossAxisAlignment.start, children: [
              Container(
                width: 54,
                padding: const EdgeInsets.symmetric(vertical: 10),
                decoration: BoxDecoration(
                    color: AppColors.ink,
                    borderRadius: BorderRadius.circular(15)),
                child: Column(children: [
                  Text(clock(job.scheduledAt),
                      style: const TextStyle(
                          color: Colors.white, fontWeight: FontWeight.w900)),
                  const SizedBox(height: 2),
                  Text(shortDate(job.scheduledAt),
                      style: const TextStyle(
                          color: Color(0xFFC9CFC7), fontSize: 9)),
                ]),
              ),
              const SizedBox(width: 13),
              Expanded(
                child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(children: [
                        Expanded(
                            child: Text(job.title,
                                style: const TextStyle(
                                    fontSize: 15,
                                    fontWeight: FontWeight.w900))),
                        StatusPill(job.status),
                      ]),
                      const SizedBox(height: 5),
                      Text(job.clientName,
                          style: const TextStyle(
                              fontWeight: FontWeight.w800,
                              color: AppColors.limeDark)),
                      const SizedBox(height: 4),
                      Text(job.address,
                          maxLines: 1,
                          overflow: TextOverflow.ellipsis,
                          style: const TextStyle(
                              color: AppColors.muted, fontSize: 12)),
                      if (!compact) ...[
                        const SizedBox(height: 10),
                        Row(children: [
                          Text(job.id,
                              style: const TextStyle(
                                  color: AppColors.muted,
                                  fontSize: 11,
                                  fontWeight: FontWeight.w700)),
                          const Spacer(),
                          Text(money(job.total),
                              style:
                                  const TextStyle(fontWeight: FontWeight.w900)),
                          const SizedBox(width: 4),
                          const Icon(Icons.chevron_right_rounded,
                              size: 19, color: AppColors.muted),
                        ]),
                      ],
                    ]),
              ),
            ]),
          ),
        ),
      );
}
