import 'package:flutter/material.dart';

import 'core/app_theme.dart';
import 'data/service_repository.dart';
import 'data/stock_repository.dart';
import 'domain/service_job.dart';
import 'state/service_store.dart';
import 'ui/home_screen.dart';
import 'ui/job_detail_screen.dart';
import 'ui/job_form_screen.dart';
import 'ui/jobs_screen.dart';
import 'ui/stock_screen.dart';

class LugestFieldApp extends StatefulWidget {
  const LugestFieldApp({super.key});

  @override
  State<LugestFieldApp> createState() => _LugestFieldAppState();
}

class _LugestFieldAppState extends State<LugestFieldApp> {
  late final ServiceStore store;

  @override
  void initState() {
    super.initState();
    store = ServiceStore(ServiceRepository(), mobileApi: StockRepository())
      ..load();
  }

  @override
  void dispose() {
    store.dispose();
    super.dispose();
  }

  @override
  Widget build(BuildContext context) => MaterialApp(
        debugShowCheckedModeBanner: false,
        title: 'LuGEST Field',
        theme: buildAppTheme(),
        home: AnimatedBuilder(
          animation: store,
          builder: (context, _) =>
              store.loading ? const _LaunchScreen() : AppShell(store: store),
        ),
      );
}

class _LaunchScreen extends StatelessWidget {
  const _LaunchScreen();

  @override
  Widget build(BuildContext context) => const Scaffold(
        body: Center(
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            CircleAvatar(
                radius: 32,
                backgroundColor: AppColors.ink,
                child:
                    Icon(Icons.bolt_rounded, color: AppColors.lime, size: 34)),
            SizedBox(height: 16),
            Text('LuGEST Field',
                style: TextStyle(fontSize: 22, fontWeight: FontWeight.w900)),
            SizedBox(height: 16),
            SizedBox(
                width: 28,
                height: 28,
                child: CircularProgressIndicator(
                    strokeWidth: 3, color: AppColors.lime)),
          ]),
        ),
      );
}

class AppShell extends StatefulWidget {
  const AppShell({super.key, required this.store});
  final ServiceStore store;

  @override
  State<AppShell> createState() => _AppShellState();
}

class _AppShellState extends State<AppShell> {
  int index = 0;

  Future<void> openJob(ServiceJob job) async {
    await Navigator.push(
        context,
        MaterialPageRoute<void>(
            builder: (_) =>
                JobDetailScreen(store: widget.store, jobId: job.id)));
  }

  Future<void> newJob() async {
    final job = await Navigator.push<ServiceJob>(context,
        MaterialPageRoute<ServiceJob>(builder: (_) => const JobFormScreen()));
    if (job == null) return;
    await widget.store.upsert(job);
    if (mounted) await openJob(job);
  }

  @override
  Widget build(BuildContext context) {
    final screens = [
      HomeScreen(
          store: widget.store,
          onOpenJob: openJob,
          onNewJob: newJob,
          onShowJobs: () => setState(() => index = 1)),
      JobsScreen(store: widget.store, onOpenJob: openJob, onNewJob: newJob),
      ClientsScreen(store: widget.store, onOpenJob: openJob),
      MoreScreen(store: widget.store),
    ];
    return Scaffold(
      body: SafeArea(child: IndexedStack(index: index, children: screens)),
      bottomNavigationBar: NavigationBar(
        selectedIndex: index,
        onDestinationSelected: (value) => setState(() => index = value),
        destinations: const [
          NavigationDestination(
              icon: Icon(Icons.today_outlined),
              selectedIcon: Icon(Icons.today_rounded),
              label: 'Hoje'),
          NavigationDestination(
              icon: Icon(Icons.handyman_outlined),
              selectedIcon: Icon(Icons.handyman_rounded),
              label: 'Serviços'),
          NavigationDestination(
              icon: Icon(Icons.people_outline_rounded),
              selectedIcon: Icon(Icons.people_rounded),
              label: 'Clientes'),
          NavigationDestination(
              icon: Icon(Icons.grid_view_outlined),
              selectedIcon: Icon(Icons.grid_view_rounded),
              label: 'Mais'),
        ],
      ),
    );
  }
}

class ClientsScreen extends StatelessWidget {
  const ClientsScreen(
      {super.key, required this.store, required this.onOpenJob});
  final ServiceStore store;
  final ValueChanged<ServiceJob> onOpenJob;

  @override
  Widget build(BuildContext context) {
    final clients = <String, List<ServiceJob>>{};
    for (final job in store.jobs) {
      clients.putIfAbsent(job.clientName, () => []).add(job);
    }
    final names = clients.keys.toList()..sort();
    return ListView(
      padding: const EdgeInsets.fromLTRB(18, 12, 18, 110),
      children: [
        Text('Clientes',
            style: Theme.of(context)
                .textTheme
                .headlineSmall
                ?.copyWith(fontWeight: FontWeight.w900)),
        const SizedBox(height: 4),
        const Text('Histórico e contexto antes de chegar ao local.',
            style: TextStyle(color: AppColors.muted)),
        const SizedBox(height: 18),
        const TextField(
            decoration: InputDecoration(
                prefixIcon: Icon(Icons.search_rounded),
                hintText: 'Pesquisar cliente')),
        const SizedBox(height: 14),
        ...names.map((name) {
          final jobs = clients[name]!
            ..sort((a, b) => b.scheduledAt.compareTo(a.scheduledAt));
          final latest = jobs.first;
          return Padding(
            padding: const EdgeInsets.only(bottom: 10),
            child: Card(
              child: ListTile(
                onTap: () => onOpenJob(latest),
                contentPadding:
                    const EdgeInsets.symmetric(horizontal: 17, vertical: 8),
                leading: CircleAvatar(
                    backgroundColor: AppColors.limeSoft,
                    child: Text(name.substring(0, 1).toUpperCase(),
                        style: const TextStyle(
                            fontWeight: FontWeight.w900,
                            color: AppColors.limeDark))),
                title: Text(name,
                    style: const TextStyle(fontWeight: FontWeight.w900)),
                subtitle: Text(
                    '${jobs.length} serviço(s) · ${latest.clientPhone}',
                    style: const TextStyle(color: AppColors.muted)),
                trailing: const Icon(Icons.chevron_right_rounded),
              ),
            ),
          );
        }),
      ],
    );
  }
}

class MoreScreen extends StatelessWidget {
  const MoreScreen({super.key, required this.store});
  final ServiceStore store;

  @override
  Widget build(BuildContext context) => ListView(
        padding: const EdgeInsets.fromLTRB(18, 12, 18, 110),
        children: [
          Text('Mais',
              style: Theme.of(context)
                  .textTheme
                  .headlineSmall
                  ?.copyWith(fontWeight: FontWeight.w900)),
          const SizedBox(height: 18),
          Card(
            child: Column(children: [
              ListTile(
                  leading: const CircleAvatar(
                      backgroundColor: AppColors.ink,
                      child: Icon(Icons.person_rounded, color: AppColors.lime)),
                  title: const Text('Luís Oliveira',
                      style: TextStyle(fontWeight: FontWeight.w900)),
                  subtitle:
                      const Text('Profissional independente · Administrador'),
                  trailing: const Icon(Icons.chevron_right_rounded)),
              const Divider(height: 1),
              _MenuTile(
                  icon: Icons.inventory_2_outlined,
                  title: 'Stock LuGEST',
                  subtitle: 'Produtos e matérias-primas em tempo real',
                  onTap: () => Navigator.push<void>(
                        context,
                        MaterialPageRoute<void>(
                          builder: (_) =>
                              StockScreen(repository: StockRepository()),
                        ),
                      )),
              _MenuTile(
                  icon: Icons.receipt_long_outlined,
                  title: 'Faturação',
                  subtitle:
                      '${store.readyToInvoice.length} serviço(s) pendente(s)',
                  onTap: () => _message(context,
                      'Os serviços prontos serão enviados para o menu Faturação do desktop.')),
              _MenuTile(
                  icon: Icons.cloud_sync_outlined,
                  title: 'Sincronização',
                  subtitle: store.pendingChanges == 0
                      ? (store.syncError ?? 'Tudo sincronizado')
                      : '${store.pendingChanges} alteração(ões) por enviar',
                  onTap: () async {
                    await store.sync();
                    if (!context.mounted) return;
                    _message(
                        context,
                        store.syncError ??
                            'Serviços sincronizados com o LuGEST.');
                  }),
              _MenuTile(
                  icon: Icons.settings_outlined,
                  title: 'Definições',
                  subtitle: 'Empresa, impostos e preferências',
                  onTap: () => _message(context,
                      'Definições protegidas pelo perfil da empresa.')),
            ]),
          ),
          const SizedBox(height: 14),
          const Center(
              child: Text('LuGEST Field · versão de testes 0.2.0',
                  style: TextStyle(color: AppColors.muted, fontSize: 11))),
        ],
      );

  static void _message(BuildContext context, String text) =>
      ScaffoldMessenger.of(context).showSnackBar(SnackBar(content: Text(text)));
}

class _MenuTile extends StatelessWidget {
  const _MenuTile(
      {required this.icon,
      required this.title,
      required this.subtitle,
      required this.onTap});
  final IconData icon;
  final String title;
  final String subtitle;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) => ListTile(
        onTap: onTap,
        leading: Icon(icon, color: AppColors.limeDark),
        title: Text(title, style: const TextStyle(fontWeight: FontWeight.w900)),
        subtitle: Text(subtitle),
        trailing: const Icon(Icons.chevron_right_rounded),
      );
}
