import 'package:flutter/material.dart';

import '../api/api_client.dart';
import '../models/pulse_models.dart';
import '../widgets/mobile_ui.dart';
import 'encomenda_detail_screen.dart';

class OrdersScreen extends StatefulWidget {
  const OrdersScreen({super.key, required this.api});

  final ApiClient api;

  @override
  State<OrdersScreen> createState() => _OrdersScreenState();
}

class _OrdersScreenState extends State<OrdersScreen> {
  late Future<List<EncomendaItem>> _future;

  @override
  void initState() {
    super.initState();
    _future = _load();
  }

  Future<List<EncomendaItem>> _load() async {
    final rows = await widget.api.getEncomendas();
    return rows
        .whereType<Map<String, dynamic>>()
        .map(EncomendaItem.fromJson)
        .toList();
  }

  Future<void> _refresh() async {
    setState(() => _future = _load());
    await _future;
  }

  void _openTopic({
    required String title,
    required String subtitle,
    required List<EncomendaItem> items,
  }) {
    Navigator.of(context).push(
      MaterialPageRoute(
        builder: (_) => _OrdersTopicScreen(
          api: widget.api,
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
        child: FutureBuilder<List<EncomendaItem>>(
          future: _future,
          builder: (context, snapshot) {
            final items = snapshot.data ?? const <EncomendaItem>[];
            final delivering =
                items.where((item) => item.dataEntrega.trim().isNotEmpty).toList();
            final inProgress = items
                .where((item) =>
                    item.pecasEmCurso > 0 ||
                    item.pecasEmPausa > 0 ||
                    item.pecasEmAvaria > 0)
                .toList();
            final risk = items
                .where((item) => item.pecasEmAvaria > 0 || item.estado.toLowerCase().contains('incomplet'))
                .toList();

            return RefreshIndicator(
              onRefresh: _refresh,
              child: ListView(
                physics: const AlwaysScrollableScrollPhysics(),
                children: [
                  MobilePageHeader(
                    title: 'Encomendas',
                    subtitle:
                        'Consulta rapida por cliente, prazo e estado da producao.',
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
                      title: 'Falha ao carregar encomendas',
                      message: snapshot.error.toString(),
                    )
                  else ...[
                    MobileHeroCard(
                      title: 'Carteira ativa',
                      subtitle:
                          'Numero, cliente e estado num ponto unico de consulta.',
                      actions: [
                        SoftChip(
                          label: '${items.length} encomendas',
                          icon: Icons.inventory_2_rounded,
                          color: Colors.white,
                          background: Colors.white.withValues(alpha: 0.14),
                        ),
                        SoftChip(
                          label: '${delivering.length} com entrega',
                          icon: Icons.local_shipping_outlined,
                          color: Colors.white,
                          background: Colors.white.withValues(alpha: 0.14),
                        ),
                      ],
                      child: MetricBoard(
                        items: [
                          MetricBoardItem(
                            label: 'Em curso',
                            value: '${inProgress.length}',
                            caption: 'Com atividade produtiva',
                            color: AppPalette.cobalt,
                          ),
                          MetricBoardItem(
                            label: 'Em risco',
                            value: '${risk.length}',
                            caption: 'Acompanhar',
                            color: AppPalette.red,
                          ),
                        ],
                      ),
                    ),
                    SizedBox(height: ui(context, 14)),
                    MobileSectionCard(
                      title: 'Abrir por assunto',
                      subtitle:
                          'Entrar pelo tema e depois abrir a encomenda certa.',
                      child: Column(
                        children: [
                          ModuleActionCard(
                            icon: Icons.list_alt_rounded,
                            color: AppPalette.navy,
                            title: 'Todas as ativas',
                            subtitle: 'Lista geral de encomendas disponiveis.',
                            badge: '${items.length}',
                            onTap: () => _openTopic(
                              title: 'Encomendas ativas',
                              subtitle: 'Consulta geral por encomenda e cliente.',
                              items: items,
                            ),
                          ),
                          SizedBox(height: ui(context, 12)),
                          ModuleActionCard(
                            icon: Icons.local_shipping_outlined,
                            color: AppPalette.orange,
                            title: 'Entregas',
                            subtitle: 'Encomendas com data de entrega preenchida.',
                            badge: '${delivering.length}',
                            onTap: () => _openTopic(
                              title: 'Entregas',
                              subtitle: 'Encomendas com prazo de entrega.',
                              items: delivering,
                            ),
                          ),
                          SizedBox(height: ui(context, 12)),
                          ModuleActionCard(
                            icon: Icons.play_circle_outline_rounded,
                            color: AppPalette.cobalt,
                            title: 'Em curso',
                            subtitle:
                                'Encomendas com pecas em curso, pausa ou avaria.',
                            badge: '${inProgress.length}',
                            onTap: () => _openTopic(
                              title: 'Em curso',
                              subtitle: 'Estado vivo da producao.',
                              items: inProgress,
                            ),
                          ),
                          SizedBox(height: ui(context, 12)),
                          ModuleActionCard(
                            icon: Icons.priority_high_rounded,
                            color: AppPalette.red,
                            title: 'Em risco',
                            subtitle: 'Itens a acompanhar mais de perto.',
                            badge: '${risk.length}',
                            onTap: () => _openTopic(
                              title: 'Em risco',
                              subtitle: 'Encomendas com sinais de alerta.',
                              items: risk,
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

class _OrdersTopicScreen extends StatefulWidget {
  const _OrdersTopicScreen({
    required this.api,
    required this.title,
    required this.subtitle,
    required this.items,
  });

  final ApiClient api;
  final String title;
  final String subtitle;
  final List<EncomendaItem> items;

  @override
  State<_OrdersTopicScreen> createState() => _OrdersTopicScreenState();
}

class _OrdersTopicScreenState extends State<_OrdersTopicScreen> {
  final _searchCtrl = TextEditingController();
  String _query = '';

  @override
  void dispose() {
    _searchCtrl.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) {
    final filtered = widget.items.where((item) {
      final q = _query.trim().toLowerCase();
      if (q.isEmpty) return true;
      return item.numero.toLowerCase().contains(q) ||
          item.cliente.toLowerCase().contains(q) ||
          item.clienteNome.toLowerCase().contains(q) ||
          item.notaCliente.toLowerCase().contains(q);
    }).toList();

    return Scaffold(
      body: MobileSurface(
        child: ListView(
          children: [
            MobilePageHeader(
              title: widget.title,
              subtitle: widget.subtitle,
              leading: IconButton.filledTonal(
                onPressed: () => Navigator.of(context).pop(),
                icon: const Icon(Icons.arrow_back_rounded),
              ),
            ),
            SizedBox(height: ui(context, 14)),
            FilterField(
              controller: _searchCtrl,
              hint: 'Pesquisar encomenda, cliente ou nome...',
              prefixIcon: Icons.search_rounded,
              onChanged: (value) => setState(() => _query = value),
            ),
            SizedBox(height: ui(context, 14)),
            if (filtered.isEmpty)
              const EmptyStateCard(
                title: 'Sem encomendas',
                message: 'Nao existem registos para este filtro.',
              )
            else
              ...filtered.map(
                (item) => Padding(
                  padding: EdgeInsets.only(bottom: ui(context, 12)),
                  child: DataRowCard(
                    title: '${item.numero}  ${item.cliente} | ${item.clienteNome}',
                    subtitle: item.dataEntrega.trim().isEmpty
                        ? item.producaoResumo
                        : 'Entrega ${item.dataEntrega} | ${item.producaoResumo}',
                    trailing: StatusPill(item.estado),
                    footer: Wrap(
                      spacing: 8,
                      runSpacing: 8,
                      children: [
                        if (item.pecasEmCurso > 0)
                          SoftChip(
                            label: '${item.pecasEmCurso} em curso',
                            icon: Icons.play_circle_outline_rounded,
                            color: AppPalette.cobalt,
                          ),
                        if (item.pecasEmPausa > 0)
                          SoftChip(
                            label: '${item.pecasEmPausa} em pausa',
                            icon: Icons.pause_circle_outline_rounded,
                            color: AppPalette.amber,
                          ),
                        if (item.pecasEmAvaria > 0)
                          SoftChip(
                            label: '${item.pecasEmAvaria} em avaria',
                            icon: Icons.warning_amber_rounded,
                            color: AppPalette.red,
                          ),
                        OutlinedButton(
                          onPressed: () {
                            Navigator.of(context).push(
                              MaterialPageRoute(
                                builder: (_) => EncomendaDetailScreen(
                                  api: widget.api,
                                  numero: item.numero,
                                ),
                              ),
                            );
                          },
                          child: const Text('Abrir detalhe'),
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
