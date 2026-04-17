import 'package:flutter/material.dart';

import '../api/api_client.dart';
import '../widgets/mobile_ui.dart';
import 'avarias_screen.dart';
import 'login_screen.dart';
import 'material_separation_screen.dart';
import 'operator_board_screen.dart';
import 'orders_screen.dart';
import 'planning_board_screen.dart';
import 'pulse_dashboard_screen.dart';

class MobileHomeShell extends StatefulWidget {
  const MobileHomeShell({super.key, required this.api});

  final ApiClient api;

  @override
  State<MobileHomeShell> createState() => _MobileHomeShellState();
}

class _MobileHomeShellState extends State<MobileHomeShell> {
  late Future<Map<String, dynamic>> _alertsFuture;

  @override
  void initState() {
    super.initState();
    _alertsFuture = widget.api.getMobileAlerts();
  }

  Future<void> _refresh() async {
    setState(() {
      _alertsFuture = widget.api.getMobileAlerts();
    });
    await _alertsFuture;
  }

  void _openModule(Widget page) {
    Navigator.of(context).push(MaterialPageRoute(builder: (_) => page));
  }

  @override
  Widget build(BuildContext context) {
    return const MobileViewSettingsScope(
      uiScale: 1.0,
      child: _HomeBody(),
    );
  }
}

class _HomeBody extends StatelessWidget {
  const _HomeBody();

  @override
  Widget build(BuildContext context) {
    final state = context.findAncestorStateOfType<_MobileHomeShellState>()!;
    return Scaffold(
      body: MobileSurface(
        child: RefreshIndicator(
          onRefresh: state._refresh,
          child: ListView(
            physics: const AlwaysScrollableScrollPhysics(),
            children: [
              MobilePageHeader(
                title: 'LuGEST Impulse',
                subtitle:
                    'Janela principal da operacao movel. Entrar num modulo, analisar e voltar sem ruido.',
                trailing: PopupMenuButton<String>(
                  onSelected: (value) {
                    if (value == 'logout') {
                      Navigator.of(context).pushAndRemoveUntil(
                        MaterialPageRoute(builder: (_) => const LoginScreen()),
                        (route) => false,
                      );
                    }
                  },
                  itemBuilder: (context) => const [
                    PopupMenuItem<String>(
                      value: 'logout',
                      child: Row(
                        children: [
                          Icon(Icons.logout_rounded),
                          SizedBox(width: 8),
                          Text('Terminar sessao'),
                        ],
                      ),
                    ),
                  ],
                ),
              ),
              SizedBox(height: ui(context, 14)),
              MobileHeroCard(
                title: 'Operacao simplificada',
                subtitle:
                    'Cada modulo abre numa pagina propria. Dentro de cada modulo entras por assuntos e vens para tras sem confusao.',
                actions: [
                  SoftChip(
                    label: state.widget.api.username.isEmpty
                        ? 'Sem utilizador'
                        : state.widget.api.username,
                    icon: Icons.person_outline_rounded,
                    color: Colors.white,
                    background: Colors.white.withValues(alpha: 0.14),
                  ),
                  SoftChip(
                    label: state.widget.api.role.isEmpty
                        ? 'Perfil'
                        : state.widget.api.role,
                    icon: Icons.badge_outlined,
                    color: Colors.white,
                    background: Colors.white.withValues(alpha: 0.14),
                  ),
                  SoftChip(
                    label: _serverHostLabel(state.widget.api.baseUrl),
                    icon: Icons.dns_outlined,
                    color: Colors.white,
                    background: Colors.white.withValues(alpha: 0.14),
                  ),
                ],
                child: Row(
                  children: [
                    Expanded(
                      child: FilledButton.icon(
                        onPressed: state._refresh,
                        style: FilledButton.styleFrom(
                          backgroundColor: Colors.white,
                          foregroundColor: AppPalette.navy,
                        ),
                        icon: const Icon(Icons.refresh_rounded),
                        label: const Text('Atualizar'),
                      ),
                    ),
                  ],
                ),
              ),
              SizedBox(height: ui(context, 14)),
              FutureBuilder<Map<String, dynamic>>(
                future: state._alertsFuture,
                builder: (context, snapshot) {
                  final data = snapshot.data ?? const <String, dynamic>{};
                  final count = (data['count'] as num?)?.toInt() ?? 0;
                  final critical = data['critical'] == true;
                  final banner = data['banner']?.toString().trim() ?? '';
                  return MobileSectionCard(
                    title: 'Atenção imediata',
                    subtitle: critical
                        ? (banner.isEmpty ? 'Existem alertas abertos.' : banner)
                        : 'Sem alertas criticos no momento.',
                    trailing: StatusPill(
                      critical ? '$count alertas' : 'Sem bloqueios',
                      color: critical ? AppPalette.red : AppPalette.green,
                    ),
                    child: MetricBoard(
                      items: [
                        MetricBoardItem(
                          label: 'Alertas',
                          value: '$count',
                          caption: critical ? 'A rever' : 'Sem urgencias',
                          color: critical ? AppPalette.red : AppPalette.green,
                        ),
                        MetricBoardItem(
                          label: 'Servidor',
                          value: _serverHostLabel(state.widget.api.baseUrl),
                          caption: 'API mobile',
                          color: AppPalette.cobalt,
                        ),
                      ],
                    ),
                  );
                },
              ),
              SizedBox(height: ui(context, 14)),
              MobileSectionCard(
                title: 'Modulos principais',
                subtitle:
                    'Entrar por assunto e aprofundar numa pagina secundaria com retroceder.',
                child: Column(
                  children: [
                    ModuleActionCard(
                      icon: Icons.insights_rounded,
                      color: AppPalette.cobalt,
                      title: 'Production Pulse',
                      subtitle: 'Ritmo, desempenho, historico e riscos.',
                      onTap: () => state._openModule(
                        PulseDashboardScreen(api: state.widget.api),
                      ),
                    ),
                    SizedBox(height: ui(context, 12)),
                    ModuleActionCard(
                      icon: Icons.layers_rounded,
                      color: AppPalette.amber,
                      title: 'Separacao MP',
                      subtitle: 'Ver o que separar no posto e marcar visto rapido.',
                      onTap: () => state._openModule(
                        MaterialSeparationScreen(api: state.widget.api),
                      ),
                    ),
                    SizedBox(height: ui(context, 12)),
                    ModuleActionCard(
                      icon: Icons.precision_manufacturing_rounded,
                      color: AppPalette.navy,
                      title: 'Operador',
                      subtitle:
                          'Grupos ativos, entregas, corte/laser e prioridades.',
                      onTap: () => state._openModule(
                        OperatorBoardScreen(api: state.widget.api),
                      ),
                    ),
                    SizedBox(height: ui(context, 12)),
                    ModuleActionCard(
                      icon: Icons.view_timeline_rounded,
                      color: AppPalette.orange,
                      title: 'Planeamento',
                      subtitle: 'Semana ativa, backlog e quadro.',
                      onTap: () => state._openModule(
                        PlanningBoardScreen(api: state.widget.api),
                      ),
                    ),
                    SizedBox(height: ui(context, 12)),
                    ModuleActionCard(
                      icon: Icons.inventory_2_rounded,
                      color: AppPalette.green,
                      title: 'Encomendas',
                      subtitle: 'Consulta por cliente, prazo e estado.',
                      onTap: () => state._openModule(
                        OrdersScreen(api: state.widget.api),
                      ),
                    ),
                    SizedBox(height: ui(context, 12)),
                    ModuleActionCard(
                      icon: Icons.warning_amber_rounded,
                      color: AppPalette.red,
                      title: 'Avarias',
                      subtitle: 'Abertas, historico e prioridade de resposta.',
                      onTap: () => state._openModule(
                        AvariasScreen(api: state.widget.api),
                      ),
                    ),
                  ],
                ),
              ),
            ],
          ),
        ),
      ),
    );
  }
}

String _serverHostLabel(String baseUrl) {
  final uri = Uri.tryParse(baseUrl);
  if (uri == null) return baseUrl;
  final host = uri.host.isEmpty ? baseUrl : uri.host;
  final port = uri.hasPort ? ':${uri.port}' : '';
  return '$host$port';
}
