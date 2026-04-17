import 'package:flutter/material.dart';

import '../api/api_client.dart';
import '../widgets/mobile_ui.dart';

class EncomendaDetailScreen extends StatefulWidget {
  const EncomendaDetailScreen({
    super.key,
    required this.api,
    required this.numero,
  });

  final ApiClient api;
  final String numero;

  @override
  State<EncomendaDetailScreen> createState() => _EncomendaDetailScreenState();
}

class _EncomendaDetailScreenState extends State<EncomendaDetailScreen> {
  late Future<Map<String, dynamic>> _future;

  @override
  void initState() {
    super.initState();
    _future = widget.api.getEncomenda(widget.numero);
  }

  Future<void> _refresh() async {
    setState(() => _future = widget.api.getEncomenda(widget.numero));
    await _future;
  }

  @override
  Widget build(BuildContext context) {
    return Scaffold(
      body: MobileSurface(
        child: FutureBuilder<Map<String, dynamic>>(
          future: _future,
          builder: (context, snapshot) {
            final item = snapshot.data ?? const <String, dynamic>{};
            final materials = List<Map<String, dynamic>>.from(
              item['materiais'] as List<dynamic>? ?? const [],
            );
            return RefreshIndicator(
              onRefresh: _refresh,
              child: ListView(
                physics: const AlwaysScrollableScrollPhysics(),
                children: [
                  MobilePageHeader(
                    title: widget.numero,
                    subtitle:
                        'Detalhe da encomenda com resumo, materiais e pecas.',
                    leading: IconButton.filledTonal(
                      onPressed: () => Navigator.of(context).pop(),
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
                      title: 'Falha ao carregar encomenda',
                      message: snapshot.error.toString(),
                    )
                  else ...[
                    MobileHeroCard(
                      title: '${item['cliente'] ?? '-'}',
                      subtitle:
                          (item['nota_cliente']?.toString().trim().isNotEmpty ??
                                  false)
                              ? item['nota_cliente'].toString()
                              : 'Resumo rapido da encomenda.',
                      actions: [
                        StatusPill('${item['estado'] ?? '-'}'),
                        if ((item['data_entrega']?.toString().trim().isNotEmpty ??
                            false))
                          SoftChip(
                            label: 'Entrega ${item['data_entrega']}',
                            icon: Icons.local_shipping_outlined,
                            color: Colors.white,
                            background: Colors.white.withValues(alpha: 0.14),
                          ),
                      ],
                      child: MetricBoard(
                        items: [
                          MetricBoardItem(
                            label: 'Tempo real',
                            value:
                                '${((item['tempo'] as num?)?.toDouble() ?? 0).toStringAsFixed(1)} h',
                            caption: 'Executado',
                            color: AppPalette.cobalt,
                          ),
                          MetricBoardItem(
                            label: 'Tempo estimado',
                            value:
                                '${((item['tempo_estimado'] as num?)?.toDouble() ?? 0).toStringAsFixed(1)} m',
                            caption: 'Planeado',
                            color: AppPalette.orange,
                          ),
                          MetricBoardItem(
                            label: 'Pronta expedicao',
                            value: '${item['qtd_pronta_expedicao'] ?? 0}',
                            caption: 'Disponivel',
                            color: AppPalette.green,
                          ),
                        ],
                      ),
                    ),
                    SizedBox(height: ui(context, 14)),
                    MobileSectionCard(
                      title: 'Materiais',
                      subtitle: 'Leitura por material, espessura e pecas.',
                      child: materials.isEmpty
                          ? const EmptyStateCard(
                              title: 'Sem materiais',
                              message: 'A encomenda nao tem materiais visiveis.',
                            )
                          : Column(
                              children: materials.map((material) {
                                final espessuras = List<Map<String, dynamic>>.from(
                                  material['espessuras'] as List<dynamic>? ?? const [],
                                );
                                return Padding(
                                  padding: EdgeInsets.only(bottom: ui(context, 12)),
                                  child: _MaterialCard(
                                    material: material,
                                    espessuras: espessuras,
                                  ),
                                );
                              }).toList(),
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

class _MaterialCard extends StatelessWidget {
  const _MaterialCard({
    required this.material,
    required this.espessuras,
  });

  final Map<String, dynamic> material;
  final List<Map<String, dynamic>> espessuras;

  @override
  Widget build(BuildContext context) {
    return Container(
      decoration: BoxDecoration(
        color: AppPalette.surfaceAlt,
        borderRadius: BorderRadius.circular(22),
        border: Border.all(color: AppPalette.line),
      ),
      padding: EdgeInsets.all(ui(context, 16)),
      child: Column(
        crossAxisAlignment: CrossAxisAlignment.start,
        children: [
          Text(
            '${material['material'] ?? '-'}',
            style: Theme.of(context).textTheme.titleLarge,
          ),
          SizedBox(height: ui(context, 10)),
          ...espessuras.map(
            (esp) {
              final pieces = List<Map<String, dynamic>>.from(
                esp['pecas'] as List<dynamic>? ?? const [],
              );
              return Padding(
                padding: EdgeInsets.only(bottom: ui(context, 12)),
                child: Container(
                  padding: EdgeInsets.all(ui(context, 14)),
                  decoration: BoxDecoration(
                    color: Colors.white,
                    borderRadius: BorderRadius.circular(18),
                    border: Border.all(color: AppPalette.line),
                  ),
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Row(
                        children: [
                          Expanded(
                            child: Text(
                              'Esp. ${esp['espessura'] ?? '-'} mm',
                              style: Theme.of(context).textTheme.titleMedium,
                            ),
                          ),
                          StatusPill('${esp['estado'] ?? '-'}'),
                        ],
                      ),
                      SizedBox(height: ui(context, 8)),
                      Wrap(
                        spacing: 8,
                        runSpacing: 8,
                        children: [
                          SoftChip(
                            label:
                                '${((esp['tempo_min'] as num?)?.toDouble() ?? 0).toStringAsFixed(0)} min',
                            icon: Icons.schedule_rounded,
                          ),
                          SoftChip(
                            label: '${pieces.length} pecas',
                            icon: Icons.widgets_outlined,
                          ),
                        ],
                      ),
                      SizedBox(height: ui(context, 10)),
                      ...pieces.map(
                        (piece) => Padding(
                          padding: EdgeInsets.only(bottom: ui(context, 8)),
                          child: Container(
                            padding: EdgeInsets.all(ui(context, 12)),
                            decoration: BoxDecoration(
                              color: AppPalette.surfaceAlt,
                              borderRadius: BorderRadius.circular(16),
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
                                  style: Theme.of(context)
                                      .textTheme
                                      .bodySmall
                                      ?.copyWith(color: AppPalette.slate),
                                ),
                                SizedBox(height: ui(context, 10)),
                                Wrap(
                                  spacing: 8,
                                  runSpacing: 8,
                                  children: [
                                    StatusPill('${piece['estado'] ?? '-'}'),
                                    SoftChip(
                                      label:
                                          '${piece['produzido_ok'] ?? piece['produzido'] ?? 0}/${piece['quantidade_pedida'] ?? piece['planeado'] ?? 0}',
                                      icon: Icons.checklist_rounded,
                                    ),
                                    SoftChip(
                                      label: shortOperation(
                                        _firstPendingOrCurrent(piece),
                                      ),
                                      icon: Icons.play_circle_outline_rounded,
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
                ),
              );
            },
          ),
        ],
      ),
    );
  }
}

String _firstPendingOrCurrent(Map<String, dynamic> piece) {
  final atual = piece['operacao_atual']?.toString().trim() ?? '';
  if (atual.isNotEmpty && atual != '-') return atual;
  final fluxo = List<Map<String, dynamic>>.from(
    piece['operacoes_fluxo'] as List<dynamic>? ?? const [],
  );
  for (final op in fluxo) {
    final estado = op['estado']?.toString().toLowerCase() ?? '';
    if (!estado.contains('concl')) {
      return op['nome']?.toString() ?? '-';
    }
  }
  return fluxo.isEmpty ? '-' : fluxo.last['nome']?.toString() ?? '-';
}
