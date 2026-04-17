import 'package:flutter/material.dart';

import '../api/api_client.dart';
import '../models/pulse_models.dart';
import '../widgets/mobile_ui.dart';

class PulseDashboardScreen extends StatefulWidget {
  const PulseDashboardScreen({super.key, required this.api});

  final ApiClient api;

  @override
  State<PulseDashboardScreen> createState() => _PulseDashboardScreenState();
}

class _PulseDashboardScreenState extends State<PulseDashboardScreen> {
  static const List<String> _periods = ['7 dias', '30 dias', 'Tudo'];

  String _period = '7 dias';
  late Future<Map<String, dynamic>> _future;

  @override
  void initState() {
    super.initState();
    _future = _load();
  }

  Future<Map<String, dynamic>> _load() {
    return widget.api.getDashboard(period: _period);
  }

  Future<void> _refresh() async {
    setState(() => _future = _load());
    await _future;
  }

  Future<void> _pickPeriod() async {
    var temp = _period;
    await showModalBottomSheet<void>(
      context: context,
      showDragHandle: true,
      builder: (context) {
        return Padding(
          padding: EdgeInsets.fromLTRB(
            ui(context, 20),
            ui(context, 8),
            ui(context, 20),
            ui(context, 20),
          ),
          child: Column(
            mainAxisSize: MainAxisSize.min,
            crossAxisAlignment: CrossAxisAlignment.start,
            children: [
              Text('Filtros do Pulse', style: Theme.of(context).textTheme.titleLarge),
              SizedBox(height: ui(context, 14)),
              FilterDropdown(
                value: temp,
                items: _periods,
                onChanged: (value) {
                  temp = value ?? temp;
                },
              ),
              SizedBox(height: ui(context, 16)),
              Row(
                children: [
                  Expanded(
                    child: OutlinedButton(
                      onPressed: () => Navigator.of(context).pop(),
                      child: const Text('Cancelar'),
                    ),
                  ),
                  SizedBox(width: ui(context, 10)),
                  Expanded(
                    child: FilledButton(
                      onPressed: () {
                        Navigator.of(context).pop();
                        if (temp != _period) {
                          setState(() {
                            _period = temp;
                            _future = _load();
                          });
                        }
                      },
                      child: const Text('Aplicar'),
                    ),
                  ),
                ],
              ),
            ],
          ),
        );
      },
    );
  }

  void _openTopic({
    required String title,
    required String subtitle,
    required List<dynamic> items,
    required Widget Function(Map<String, dynamic>) builder,
  }) {
    Navigator.of(context).push(
      MaterialPageRoute(
        builder: (_) => _PulseTopicScreen(
          title: title,
          subtitle: subtitle,
          items: items.cast<dynamic>(),
          itemBuilder: builder,
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
            final loading = snapshot.connectionState == ConnectionState.waiting &&
                !snapshot.hasData;
            final data = snapshot.data ?? const <String, dynamic>{};
            final summary = PulseSummary.fromJson(
              data['summary'] as Map<String, dynamic>? ?? const <String, dynamic>{},
            );
            final running = List<Map<String, dynamic>>.from(
              data['running'] as List<dynamic>? ?? const [],
            );
            final history = List<Map<String, dynamic>>.from(
              data['history'] as List<dynamic>? ?? const [],
            );
            final topStops = List<Map<String, dynamic>>.from(
              data['top_stops'] as List<dynamic>? ?? const [],
            );
            final interrupted = List<Map<String, dynamic>>.from(
              data['interrupted'] as List<dynamic>? ?? const [],
            );
            return RefreshIndicator(
              onRefresh: _refresh,
              child: ListView(
                physics: const AlwaysScrollableScrollPhysics(),
                children: [
                  MobilePageHeader(
                    title: 'Production Pulse',
                    subtitle: 'Leitura rapida do turno, desempenho e risco.',
                    leading: IconButton.filledTonal(
                      onPressed: () => Navigator.of(context).maybePop(),
                      icon: const Icon(Icons.arrow_back_rounded),
                    ),
                    trailing: Row(
                      mainAxisSize: MainAxisSize.min,
                      children: [
                        IconButton.outlined(
                          onPressed: _pickPeriod,
                          icon: const Icon(Icons.tune_rounded),
                        ),
                        SizedBox(width: ui(context, 8)),
                        FilledButton.icon(
                          onPressed: _refresh,
                          icon: const Icon(Icons.refresh_rounded),
                          label: const Text('Atualizar'),
                        ),
                      ],
                    ),
                  ),
                  SizedBox(height: ui(context, 14)),
                  MobileHeroCard(
                    title: 'Foco do turno',
                    subtitle:
                        'OEE, carga ativa, desvios e paragens num painel simples.',
                    actions: [
                      SoftChip(
                        label: _period,
                        icon: Icons.calendar_today_rounded,
                        color: Colors.white,
                        background: Colors.white.withValues(alpha: 0.14),
                      ),
                      SoftChip(
                        label: data['updated_at']?.toString().split('T').last.substring(0, 5) ?? '-',
                        icon: Icons.schedule_rounded,
                        color: Colors.white,
                        background: Colors.white.withValues(alpha: 0.14),
                      ),
                    ],
                    child: MetricBoard(
                      items: [
                        MetricBoardItem(
                          label: 'OEE global',
                          value: '${summary.oee.toStringAsFixed(1)}%',
                          caption: 'Visao geral',
                          color: AppPalette.green,
                        ),
                        MetricBoardItem(
                          label: 'Em curso',
                          value: '${summary.pecasEmCurso}',
                          caption: 'Pecas com atividade',
                          color: AppPalette.cobalt,
                        ),
                        MetricBoardItem(
                          label: 'Paragens',
                          value: '${summary.paragensMin.toStringAsFixed(1)}m',
                          caption: 'Tempo acumulado',
                          color: AppPalette.amber,
                        ),
                      ],
                    ),
                  ),
                  SizedBox(height: ui(context, 14)),
                  if (loading)
                    const Center(child: Padding(
                      padding: EdgeInsets.all(24),
                      child: CircularProgressIndicator(),
                    ))
                  else if (snapshot.hasError)
                    EmptyStateCard(
                      title: 'Falha ao carregar',
                      message: snapshot.error.toString(),
                      icon: Icons.error_outline_rounded,
                    )
                  else ...[
                    MobileSectionCard(
                      title: 'Base produtiva',
                      subtitle: 'Indicadores principais antes de entrares no detalhe.',
                      child: MetricBoard(
                        items: [
                          MetricBoardItem(
                            label: 'Disponibilidade',
                            value:
                                '${summary.disponibilidade.toStringAsFixed(1)}%',
                            caption: 'Capacidade real',
                            color: AppPalette.green,
                          ),
                          MetricBoardItem(
                            label: 'Performance',
                            value:
                                '${summary.performance.toStringAsFixed(1)}%',
                            caption: 'Ritmo do plano',
                            color: AppPalette.cobalt,
                          ),
                          MetricBoardItem(
                            label: 'Qualidade',
                            value:
                                '${summary.qualidade.toStringAsFixed(1)}%',
                            caption: summary.qualityScope.isEmpty
                                ? 'Periodo atual'
                                : summary.qualityScope,
                            color: AppPalette.green,
                          ),
                          MetricBoardItem(
                            label: 'Desvio max.',
                            value:
                                '${summary.desvioMaxMin.toStringAsFixed(1)}m',
                            caption: 'Maior atraso',
                            color: summary.desvioMaxMin > 0
                                ? AppPalette.red
                                : AppPalette.green,
                          ),
                        ],
                      ),
                    ),
                    SizedBox(height: ui(context, 14)),
                    MobileSectionCard(
                      title: 'Abrir assuntos',
                      subtitle:
                          'Cada botao abre uma pagina secundaria so com esse tema.',
                      child: Column(
                        children: [
                          ModuleActionCard(
                            icon: Icons.play_circle_outline_rounded,
                            color: AppPalette.cobalt,
                            title: 'Em curso',
                            subtitle: 'Pecas e operacoes a correr agora.',
                            badge: '${running.length}',
                            onTap: () => _openTopic(
                              title: 'Em curso',
                              subtitle: 'Leitura detalhada das pecas ativas.',
                              items: running,
                              builder: _buildRunningRow,
                            ),
                          ),
                          SizedBox(height: ui(context, 12)),
                          ModuleActionCard(
                            icon: Icons.history_rounded,
                            color: AppPalette.navy,
                            title: 'Historico',
                            subtitle: 'Encomendas medidas no periodo filtrado.',
                            badge: '${history.length}',
                            onTap: () => _openTopic(
                              title: 'Historico',
                              subtitle: 'Leitura temporal por encomenda.',
                              items: history,
                              builder: _buildHistoryRow,
                            ),
                          ),
                          SizedBox(height: ui(context, 12)),
                          ModuleActionCard(
                            icon: Icons.pause_circle_outline_rounded,
                            color: AppPalette.amber,
                            title: 'Paragens',
                            subtitle: 'Top de paragens e interrupcoes.',
                            badge: '${topStops.length + interrupted.length}',
                            onTap: () => _openTopic(
                              title: 'Paragens e interrupcoes',
                              subtitle: 'Pontos que podem pedir acao da chefia.',
                              items: [...topStops, ...interrupted],
                              builder: _buildStopRow,
                            ),
                          ),
                        ],
                      ),
                    ),
                    if ((summary.alerts).trim().isNotEmpty) ...[
                      SizedBox(height: ui(context, 14)),
                      MobileSectionCard(
                        title: 'Leitura imediata',
                        subtitle: summary.alerts.trim(),
                        child: StatusPill(
                          summary.alerts.contains('Sem alertas')
                              ? 'Sem alertas criticos'
                              : 'Acompanhar',
                          color: summary.alerts.contains('Sem alertas')
                              ? AppPalette.green
                              : AppPalette.red,
                        ),
                      ),
                    ],
                  ],
                ],
              ),
            );
          },
        ),
      ),
    );
  }

  Widget _buildRunningRow(Map<String, dynamic> row) {
    final title =
        '${row['encomenda'] ?? '-'}  ${row['ref_interna'] ?? row['peca'] ?? '-'}';
    final subtitle =
        '${row['operacao'] ?? row['estado'] ?? '-'} | ${row['elapsed_min'] ?? row['tempo_min'] ?? '-'} min';
    return DataRowCard(
      title: title,
      subtitle: subtitle,
      trailing: StatusPill(row['fora'] == true ? 'Com desvio' : 'Normal'),
      footer: Wrap(
        spacing: 8,
        runSpacing: 8,
        children: [
          SoftChip(label: '${row['cliente'] ?? '-'}', icon: Icons.person_outline_rounded),
          SoftChip(label: '${row['plan_min'] ?? '-'} min plano', icon: Icons.schedule_rounded),
        ],
      ),
    );
  }

  Widget _buildHistoryRow(Map<String, dynamic> row) {
    final delta = (row['delta_min'] as num?)?.toDouble() ?? 0;
    return DataRowCard(
      title: '${row['encomenda'] ?? '-'}',
      subtitle:
          '${row['ops'] ?? 0} ops | ${row['elapsed_min'] ?? 0} min reais | ${row['plan_min'] ?? 0} min plano',
      trailing: StatusPill(
        delta > 0 ? 'Atraso' : 'Dentro do plano',
        color: delta > 0 ? AppPalette.red : AppPalette.green,
      ),
    );
  }

  Widget _buildStopRow(Map<String, dynamic> row) {
    return DataRowCard(
      title: '${row['encomenda'] ?? '-'}',
      subtitle:
          '${row['motivo'] ?? row['title'] ?? 'Paragem'} | ${row['duracao_txt'] ?? row['duracao_min'] ?? '-'}',
      trailing: StatusPill('${row['estado'] ?? 'Paragem'}'),
    );
  }
}

class _PulseTopicScreen extends StatelessWidget {
  const _PulseTopicScreen({
    required this.title,
    required this.subtitle,
    required this.items,
    required this.itemBuilder,
  });

  final String title;
  final String subtitle;
  final List<dynamic> items;
  final Widget Function(Map<String, dynamic>) itemBuilder;

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
                title: 'Sem registos',
                message: 'Nao existem elementos para este assunto.',
              )
            else
              ...items
                  .whereType<Map<String, dynamic>>()
                  .map((row) => Padding(
                        padding: EdgeInsets.only(bottom: ui(context, 12)),
                        child: itemBuilder(row),
                      )),
          ],
        ),
      ),
    );
  }
}
