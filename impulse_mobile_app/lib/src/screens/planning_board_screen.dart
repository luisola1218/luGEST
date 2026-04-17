import 'package:flutter/material.dart';

import '../api/api_client.dart';
import '../support/pdf_launcher.dart';
import '../widgets/mobile_ui.dart';

class PlanningBoardScreen extends StatefulWidget {
  const PlanningBoardScreen({super.key, required this.api});

  final ApiClient api;

  @override
  State<PlanningBoardScreen> createState() => _PlanningBoardScreenState();
}

class _PlanningBoardScreenState extends State<PlanningBoardScreen> {
  late Future<Map<String, dynamic>> _future;

  @override
  void initState() {
    super.initState();
    _future = widget.api.getPlanningOverview();
  }

  Future<void> _refresh() async {
    setState(() => _future = widget.api.getPlanningOverview());
    await _future;
  }

  Future<void> _openPdf() async {
    final url = widget.api.planningPdfUrl();
    await openPdfUrl(context, url);
  }

  void _openTopic({
    required String title,
    required String subtitle,
    required List<Map<String, dynamic>> items,
  }) {
    Navigator.of(context).push(
      MaterialPageRoute(
        builder: (_) => _PlanningTopicScreen(
          title: title,
          subtitle: subtitle,
          items: items,
        ),
      ),
    );
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: MobileSurface(
        child: FutureBuilder<Map<String, dynamic>>(
          future: _future,
          builder: (context, snapshot) {
            final data = snapshot.data ?? const <String, dynamic>{};
            final summary = data['summary'] as Map<String, dynamic>? ?? const {};
            final active = List<Map<String, dynamic>>.from(
              data['active'] as List<dynamic>? ?? const [],
            );
            final backlog = List<Map<String, dynamic>>.from(
              data['backlog'] as List<dynamic>? ?? const [],
            );
            final history = List<Map<String, dynamic>>.from(
              data['history'] as List<dynamic>? ?? const [],
            );
            final label = summary['week_label']?.toString() ?? '';

            return RefreshIndicator(
              onRefresh: _refresh,
              child: ListView(
                physics: const AlwaysScrollableScrollPhysics(),
                children: [
                  MobilePageHeader(
                    title: 'Planeamento',
                    subtitle:
                        'Semana ativa, backlog e quadro em paginas separadas.',
                    leading: IconButton.filledTonal(
                      onPressed: () => Navigator.of(context).maybePop(),
                      icon: const Icon(Icons.arrow_back_rounded),
                    ),
                    trailing: FilledButton.icon(
                      onPressed: _refresh,
                      icon: const Icon(Icons.refresh_rounded),
                      label: const Text('Atualizar'),
                    ),
                  ),
                  SizedBox(height: ui(context, 14)),
                  if (snapshot.connectionState == ConnectionState.waiting &&
                      !snapshot.hasData)
                    const Center(
                      child: Padding(
                        padding: EdgeInsets.all(24),
                        child: CircularProgressIndicator(),
                      ),
                    )
                  else if (snapshot.hasError)
                    EmptyStateCard(
                      title: 'Falha ao carregar planeamento',
                      message: snapshot.error.toString(),
                    )
                  else ...[
                    MobileHeroCard(
                      title: 'Semana $label',
                      subtitle:
                          'Blocos ativos, backlog e acesso direto ao quadro PDF.',
                      actions: [
                        SoftChip(
                          label: '${summary['blocos_ativos'] ?? 0} blocos',
                          icon: Icons.view_timeline_rounded,
                          color: Colors.white,
                          background: Colors.white.withValues(alpha: 0.14),
                        ),
                        SoftChip(
                          label: '${summary['backlog'] ?? 0} backlog',
                          icon: Icons.backup_table_rounded,
                          color: Colors.white,
                          background: Colors.white.withValues(alpha: 0.14),
                        ),
                      ],
                      child: Column(
                        children: [
                          MetricBoard(
                            items: [
                              MetricBoardItem(
                                label: 'Min ativos',
                                value: ((summary['min_ativos_total'] as num?)
                                            ?.toDouble() ??
                                        0)
                                    .toStringAsFixed(0),
                                caption: 'Carga planeada',
                                color: AppPalette.cobalt,
                              ),
                              MetricBoardItem(
                                label: 'Historico mes',
                                value: (summary['historico_mes'] ?? 0).toString(),
                                caption: 'Blocos fechados',
                                color: AppPalette.green,
                              ),
                            ],
                          ),
                          SizedBox(height: ui(context, 14)),
                          Row(
                            children: [
                              Expanded(
                                child: OutlinedButton.icon(
                                  onPressed: _openPdf,
                                  icon: const Icon(Icons.picture_as_pdf_rounded),
                                  label: const Text('Abrir quadro PDF'),
                                  style: OutlinedButton.styleFrom(
                                    foregroundColor: Colors.white,
                                    side: BorderSide(
                                      color: Colors.white.withValues(alpha: 0.34),
                                    ),
                                  ),
                                ),
                              ),
                            ],
                          ),
                        ],
                      ),
                    ),
                    SizedBox(height: ui(context, 14)),
                    MobileSectionCard(
                      title: 'Abrir por assunto',
                      subtitle: 'Cada tema em janela propria com leitura limpa.',
                      child: Column(
                        children: [
                          ModuleActionCard(
                            icon: Icons.calendar_view_week_rounded,
                            color: AppPalette.cobalt,
                            title: 'Semana ativa',
                            subtitle: 'Blocos planeados para os proximos dias.',
                            badge: '${active.length}',
                            onTap: () => _openTopic(
                              title: 'Semana ativa',
                              subtitle: 'Planeamento atual por bloco.',
                              items: active,
                            ),
                          ),
                          SizedBox(height: ui(context, 12)),
                          ModuleActionCard(
                            icon: Icons.pending_actions_rounded,
                            color: AppPalette.orange,
                            title: 'Backlog',
                            subtitle: 'Encomendas sem bloco ou por planear.',
                            badge: '${backlog.length}',
                            onTap: () => _openTopic(
                              title: 'Backlog',
                              subtitle: 'Carteira que ainda precisa de planeamento.',
                              items: backlog,
                            ),
                          ),
                          SizedBox(height: ui(context, 12)),
                          ModuleActionCard(
                            icon: Icons.history_toggle_off_rounded,
                            color: AppPalette.navy,
                            title: 'Historico',
                            subtitle: 'Blocos fechados e referencia do mes.',
                            badge: '${history.length}',
                            onTap: () => _openTopic(
                              title: 'Historico',
                              subtitle: 'Leitura do historico planeado.',
                              items: history,
                            ),
                          ),
                        ],
                      ),
                    ),
                  ],
                ],
              ),
            );
          },
        ),
      ),
    );
  }
}

class _PlanningTopicScreen extends StatelessWidget {
  const _PlanningTopicScreen({
    required this.title,
    required this.subtitle,
    required this.items,
  });

  final String title;
  final String subtitle;
  final List<Map<String, dynamic>> items;

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: MobileSurface(
        child: ListView(
          children: [
            MobilePageHeader(
              title: title,
              subtitle: subtitle,
              leading: IconButton.filledTonal(
                onPressed: () => Navigator.of(context).pop(),
                icon: const Icon(Icons.arrow_back_rounded),
              ),
            ),
            SizedBox(height: ui(context, 14)),
            if (items.isEmpty)
              const EmptyStateCard(
                title: 'Sem dados',
                message: 'Nao existem registos para este tema.',
              )
            else
              ...items.map(
                (item) => Padding(
                  padding: EdgeInsets.only(bottom: ui(context, 12)),
                  child: DataRowCard(
                    title:
                        '${item['encomenda'] ?? item['numero'] ?? '-'}  ${item['cliente_display'] ?? item['cliente'] ?? '-'}',
                    subtitle: [
                      if ((item['data']?.toString().isNotEmpty ?? false))
                        '${item['data']} ${item['inicio'] ?? ''}',
                      if ((item['material']?.toString().isNotEmpty ?? false))
                        '${item['material']} ${item['espessura'] ?? '-'} mm',
                      if ((item['estado']?.toString().isNotEmpty ?? false))
                        '${item['estado']}',
                    ].join(' | '),
                    trailing: StatusPill(
                      '${item['duracao_min'] ?? item['data_entrega'] ?? '-'}',
                      color: AppPalette.cobalt,
                    ),
                  ),
                ),
              ),
          ],
        ),
      ),
    );
  }
}
