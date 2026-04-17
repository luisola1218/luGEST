import 'package:flutter/material.dart';

import '../api/api_client.dart';
import '../widgets/mobile_ui.dart';

class AvariasScreen extends StatefulWidget {
  const AvariasScreen({super.key, required this.api});

  final ApiClient api;

  @override
  State<AvariasScreen> createState() => _AvariasScreenState();
}

class _AvariasScreenState extends State<AvariasScreen> {
  late Future<Map<String, dynamic>> _future;

  @override
  void initState() {
    super.initState();
    _future = widget.api.getAvarias();
  }

  Future<void> _refresh() async {
    setState(() => _future = widget.api.getAvarias());
    await _future;
  }

  void _openTopic({
    required String title,
    required String subtitle,
    required List<Map<String, dynamic>> items,
  }) {
    Navigator.of(context).push(
      MaterialPageRoute(
        builder: (_) => _AvariasTopicScreen(
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
            final open = List<Map<String, dynamic>>.from(
              data['open'] as List<dynamic>? ?? const [],
            );
            final history = List<Map<String, dynamic>>.from(
              data['history'] as List<dynamic>? ?? const [],
            );

            return RefreshIndicator(
              onRefresh: _refresh,
              child: ListView(
                physics: const AlwaysScrollableScrollPhysics(),
                children: [
                  MobilePageHeader(
                    title: 'Avarias',
                    subtitle: 'Abertas agora ou resolvidas no historico recente.',
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
                      title: 'Falha ao carregar avarias',
                      message: snapshot.error.toString(),
                    )
                  else ...[
                    MobileHeroCard(
                      title: 'Radar de avarias',
                      subtitle: (summary['headline']?.toString().trim().isNotEmpty ??
                              false)
                          ? summary['headline'].toString()
                          : 'Sem bloqueios criticos no momento.',
                      actions: [
                        SoftChip(
                          label: '${summary['open_count'] ?? 0} abertas',
                          icon: Icons.warning_amber_rounded,
                          color: Colors.white,
                          background: Colors.white.withValues(alpha: 0.14),
                        ),
                        SoftChip(
                          label: '${summary['history_count'] ?? 0} historico',
                          icon: Icons.history_rounded,
                          color: Colors.white,
                          background: Colors.white.withValues(alpha: 0.14),
                        ),
                      ],
                      child: MetricBoard(
                        items: [
                          MetricBoardItem(
                            label: 'Abertas',
                            value: '${summary['open_count'] ?? 0}',
                            caption: 'A rever',
                            color: AppPalette.red,
                          ),
                          MetricBoardItem(
                            label: 'Historico',
                            value: '${summary['history_count'] ?? 0}',
                            caption: 'Registos fechados',
                            color: AppPalette.cobalt,
                          ),
                        ],
                      ),
                    ),
                    SizedBox(height: ui(context, 14)),
                    MobileSectionCard(
                      title: 'Abrir por assunto',
                      subtitle: 'Separado por abertas e historico.',
                      child: Column(
                        children: [
                          ModuleActionCard(
                            icon: Icons.error_outline_rounded,
                            color: AppPalette.red,
                            title: 'Abertas',
                            subtitle: 'Avarias ainda em curso.',
                            badge: '${open.length}',
                            onTap: () => _openTopic(
                              title: 'Avarias abertas',
                              subtitle: 'Ocorrencias que ainda precisam de resposta.',
                              items: open,
                            ),
                          ),
                          SizedBox(height: ui(context, 12)),
                          ModuleActionCard(
                            icon: Icons.history_toggle_off_rounded,
                            color: AppPalette.navy,
                            title: 'Historico',
                            subtitle: 'Ultimos registos fechados.',
                            badge: '${history.length}',
                            onTap: () => _openTopic(
                              title: 'Historico de avarias',
                              subtitle: 'Ocorrencias encerradas recentemente.',
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

class _AvariasTopicScreen extends StatelessWidget {
  const _AvariasTopicScreen({
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
                title: 'Sem registos',
                message: 'Nao existem elementos neste tema.',
              )
            else
              ...items.map(
                (item) => Padding(
                  padding: EdgeInsets.only(bottom: ui(context, 12)),
                  child: DataRowCard(
                    title:
                        '${item['encomenda'] ?? '-'}  ${item['ref_interna'] ?? item['peca_id'] ?? '-'}',
                    subtitle:
                        '${item['posto'] ?? '-'} | ${item['motivo'] ?? 'Avaria'} | ${item['duracao_txt'] ?? '-'}',
                    trailing: StatusPill('${item['estado'] ?? '-'}'),
                    footer: Wrap(
                      spacing: 8,
                      runSpacing: 8,
                      children: [
                        if ((item['operador']?.toString().trim().isNotEmpty ?? false))
                          SoftChip(
                            label: '${item['operador']}',
                            icon: Icons.person_outline_rounded,
                          ),
                        if ((item['material']?.toString().trim().isNotEmpty ?? false))
                          SoftChip(
                            label:
                                '${item['material']} ${item['espessura'] ?? '-'} mm',
                            icon: Icons.layers_outlined,
                          ),
                      ],
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
