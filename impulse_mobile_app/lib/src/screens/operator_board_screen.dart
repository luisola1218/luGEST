import 'package:flutter/material.dart';

import '../api/api_client.dart';
import '../widgets/mobile_ui.dart';

class OperatorBoardScreen extends StatefulWidget {
  const OperatorBoardScreen({super.key, required this.api});

  final ApiClient api;

  @override
  State<OperatorBoardScreen> createState() => _OperatorBoardScreenState();
}

class _OperatorBoardScreenState extends State<OperatorBoardScreen> {
  late Future<Map<String, dynamic>> _future;

  @override
  void initState() {
    super.initState();
    _future = widget.api.getOperatorBoard();
  }

  Future<void> _refresh() async {
    setState(() => _future = widget.api.getOperatorBoard());
    await _future;
  }

  void _openList({
    required String title,
    required String subtitle,
    required List<Map<String, dynamic>> items,
  }) {
    Navigator.of(context).push(
      MaterialPageRoute(
        builder: (_) => _OperatorTopicScreen(
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
            final items = List<Map<String, dynamic>>.from(
              data['items'] as List<dynamic>? ?? const [],
            );
            final deliveryItems = items
                .where((row) =>
                    (row['data_entrega']?.toString().trim().isNotEmpty ?? false))
                .toList();
            final cutLaserItems = items
                .where((row) => row['tem_corte_laser'] == true)
                .toList();
            final riskItems = items
                .where((row) =>
                    ((row['pecas_em_avaria'] as num?)?.toInt() ?? 0) > 0 ||
                    ((row['desvio_min'] as num?)?.toDouble() ?? 0) > 0)
                .toList();

            return RefreshIndicator(
              onRefresh: _refresh,
              child: ListView(
                physics: const AlwaysScrollableScrollPhysics(),
                children: [
                  MobilePageHeader(
                    title: 'Operador',
                    subtitle:
                        'Entrar por assunto e aprofundar cada grupo numa janela secundaria.',
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
                      title: 'Falha ao carregar operador',
                      message: snapshot.error.toString(),
                      icon: Icons.error_outline_rounded,
                    )
                  else ...[
                    MobileHeroCard(
                      title: 'Resumo operacional',
                      subtitle:
                          'Grupos ativos, progresso geral e carga em risco.',
                      actions: [
                        SoftChip(
                          label: '${summary['encomendas_ativas'] ?? 0} encomendas',
                          icon: Icons.inventory_2_rounded,
                          color: Colors.white,
                          background: Colors.white.withValues(alpha: 0.14),
                        ),
                        SoftChip(
                          label: '${summary['grupos'] ?? 0} grupos',
                          icon: Icons.view_stream_rounded,
                          color: Colors.white,
                          background: Colors.white.withValues(alpha: 0.14),
                        ),
                      ],
                      child: MetricBoard(
                        items: [
                          MetricBoardItem(
                            label: 'Progresso',
                            value:
                                '${((summary['progress_pct'] as num?)?.toDouble() ?? 0).toStringAsFixed(1)}%',
                            caption: 'Carga produzida',
                            color: AppPalette.green,
                          ),
                          MetricBoardItem(
                            label: 'Pecas em curso',
                            value: '${summary['pecas_em_curso'] ?? 0}',
                            caption: 'Operacao aberta',
                            color: AppPalette.cobalt,
                          ),
                          MetricBoardItem(
                            label: 'Avarias',
                            value: '${summary['pecas_em_avaria'] ?? 0}',
                            caption: 'Pontos de bloqueio',
                            color: AppPalette.red,
                          ),
                        ],
                      ),
                    ),
                    SizedBox(height: ui(context, 14)),
                    MobileSectionCard(
                      title: 'Abrir por assunto',
                      subtitle:
                          'Sem misturar tudo no mesmo ecrã. Escolhes o tema e entras.',
                      child: Column(
                        children: [
                          ModuleActionCard(
                            icon: Icons.widgets_outlined,
                            color: AppPalette.navy,
                            title: 'Grupos ativos',
                            subtitle:
                                'Todos os grupos com cliente, material, progresso e estado.',
                            badge: '${items.length}',
                            onTap: () => _openList(
                              title: 'Grupos ativos',
                              subtitle:
                                  'Visao geral dos grupos ativos e concluidos.',
                              items: items,
                            ),
                          ),
                          SizedBox(height: ui(context, 12)),
                          ModuleActionCard(
                            icon: Icons.local_shipping_outlined,
                            color: AppPalette.orange,
                            title: 'Entregas',
                            subtitle:
                                'Itens com data de entrega ou carga a acompanhar.',
                            badge: '${deliveryItems.length}',
                            onTap: () => _openList(
                              title: 'Entregas',
                              subtitle: 'Grupos com prazo e leitura de entrega.',
                              items: deliveryItems,
                            ),
                          ),
                          SizedBox(height: ui(context, 12)),
                          ModuleActionCard(
                            icon: Icons.cut_rounded,
                            color: AppPalette.cobalt,
                            title: 'Corte / Laser',
                            subtitle:
                                'Grupos com prazo de corte/laser destacado.',
                            badge: '${cutLaserItems.length}',
                            onTap: () => _openList(
                              title: 'Corte / Laser',
                              subtitle: 'Prazos planeados e grupos associados.',
                              items: cutLaserItems,
                            ),
                          ),
                          SizedBox(height: ui(context, 12)),
                          ModuleActionCard(
                            icon: Icons.priority_high_rounded,
                            color: AppPalette.red,
                            title: 'Riscos',
                            subtitle:
                                'Avarias e desvios positivos que pedem resposta.',
                            badge: '${riskItems.length}',
                            onTap: () => _openList(
                              title: 'Riscos',
                              subtitle: 'Grupos com sinais de bloqueio ou atraso.',
                              items: riskItems,
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

class _OperatorTopicScreen extends StatelessWidget {
  const _OperatorTopicScreen({
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
                title: 'Sem grupos para mostrar',
                message: 'Nao existem grupos neste assunto.',
              )
            else
              ...items.map(
                (item) => Padding(
                  padding: EdgeInsets.only(bottom: ui(context, 12)),
                  child: _OperatorGroupCard(item: item),
                ),
              ),
          ],
        ),
      ),
    );
  }
}

class _OperatorGroupCard extends StatelessWidget {
  const _OperatorGroupCard({required this.item});

  final Map<String, dynamic> item;

  @override
  Widget build(BuildContext context) {
    final pieces = List<Map<String, dynamic>>.from(
      item['pieces'] as List<dynamic>? ?? const [],
    );
    final progress = ((item['progress_pct'] as num?)?.toDouble() ?? 0);
    final prazoData = item['prazo_corte_laser_data']?.toString().trim() ?? '';
    final prazoInicio = item['prazo_corte_laser_inicio']?.toString().trim() ?? '';
    return DataRowCard(
      title:
          '${item['encomenda'] ?? '-'}  ${item['cliente_display'] ?? item['cliente'] ?? '-'}',
      subtitle:
          '${item['material'] ?? '-'} ${item['espessura'] ?? '-'} mm | ${item['estado_espessura'] ?? '-'}',
      trailing: StatusPill('${item['estado_espessura'] ?? '-'}'),
      footer: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Wrap(
            spacing: 8,
            runSpacing: 8,
            children: [
              SoftChip(
                label: 'Progresso ${progress.toStringAsFixed(0)}%',
                icon: Icons.donut_large_rounded,
              ),
              SoftChip(
                label:
                    '${item['produzido_total'] ?? 0}/${item['planeado_total'] ?? 0}',
                icon: Icons.checklist_rounded,
              ),
              if (prazoData.isNotEmpty)
                SoftChip(
                  label: prazoInicio.isEmpty ? prazoData : '$prazoData $prazoInicio',
                  icon: Icons.schedule_rounded,
                  color: AppPalette.cobalt,
                ),
            ],
          ),
          SizedBox(height: ui(context, 12)),
          if (pieces.isEmpty)
            const EmptyStateCard(
              title: 'Sem pecas',
              message: 'Este grupo nao tem pecas visiveis.',
            )
          else
            ...pieces.take(4).map(
              (piece) => Padding(
                padding: EdgeInsets.only(bottom: ui(context, 8)),
                child: Container(
                  padding: EdgeInsets.all(ui(context, 14)),
                  decoration: BoxDecoration(
                    color: AppPalette.surfaceAlt,
                    borderRadius: BorderRadius.circular(18),
                    border: Border.all(color: AppPalette.line),
                  ),
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(
                        '${piece['ref_interna'] ?? '-'}',
                        style: Theme.of(context).textTheme.titleMedium,
                      ),
                      SizedBox(height: ui(context, 4)),
                      Text(
                        '${piece['ref_externa'] ?? '-'}',
                        style: Theme.of(context).textTheme.bodySmall?.copyWith(
                              color: AppPalette.slate,
                            ),
                      ),
                      SizedBox(height: ui(context, 10)),
                      Wrap(
                        spacing: 8,
                        runSpacing: 8,
                        children: [
                          StatusPill('${piece['estado'] ?? '-'}'),
                          SoftChip(
                            label: shortOperation(
                              piece['operacao_atual']?.toString() ?? '-',
                            ),
                            icon: Icons.play_circle_outline_rounded,
                          ),
                          SoftChip(
                            label:
                                '${piece['produzido'] ?? 0}/${piece['planeado'] ?? 0}',
                            icon: Icons.functions_rounded,
                          ),
                        ],
                      ),
                    ],
                  ),
                ),
              ),
            ),
        ],
      ),
    );
  }
}
