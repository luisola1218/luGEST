import 'dart:async';

import 'package:flutter/material.dart';

import '../core/app_theme.dart';
import '../core/formatters.dart';
import '../data/stock_repository.dart';
import '../domain/stock_item.dart';

class StockScreen extends StatefulWidget {
  const StockScreen({
    super.key,
    required this.repository,
    this.pickMode = false,
  });

  final StockRepository repository;
  final bool pickMode;

  @override
  State<StockScreen> createState() => _StockScreenState();
}

class _StockScreenState extends State<StockScreen>
    with SingleTickerProviderStateMixin {
  late final TabController tabs = TabController(length: 2, vsync: this);
  final search = TextEditingController();
  Timer? debounce;
  bool loading = true;
  bool configured = false;
  String? error;
  StockSummary? summary;
  List<StockItem> products = const [];
  List<StockItem> materials = const [];

  @override
  void initState() {
    super.initState();
    _load();
  }

  @override
  void dispose() {
    debounce?.cancel();
    search.dispose();
    tabs.dispose();
    super.dispose();
  }

  Future<void> _load() async {
    setState(() {
      loading = true;
      error = null;
    });
    final settings = await widget.repository.loadSettings();
    configured = settings.isConfigured;
    if (!configured) {
      if (mounted) setState(() => loading = false);
      return;
    }
    try {
      final results = await Future.wait([
        widget.repository.summary(),
        widget.repository.items(kind: 'product', query: search.text),
        widget.repository.items(kind: 'material', query: search.text),
      ]);
      if (!mounted) return;
      setState(() {
        summary = results[0] as StockSummary;
        products = results[1] as List<StockItem>;
        materials = results[2] as List<StockItem>;
        loading = false;
      });
    } on StockConnectionException catch (exception) {
      if (mounted) {
        setState(() {
          error = exception.message;
          loading = false;
        });
      }
    }
  }

  void _searchChanged(String _) {
    debounce?.cancel();
    debounce = Timer(const Duration(milliseconds: 350), _load);
  }

  Future<void> _configure() async {
    final current = await widget.repository.loadSettings();
    if (!mounted) return;
    final result = await showDialog<StockApiSettings>(
      context: context,
      builder: (_) => _ConnectionDialog(
        repository: widget.repository,
        initial: current,
      ),
    );
    if (result != null) await _load();
  }

  @override
  Widget build(BuildContext context) => Scaffold(
        appBar: AppBar(
          title: Text(widget.pickMode ? 'Escolher do stock' : 'Stock LuGEST'),
          actions: [
            IconButton(
              onPressed: _configure,
              tooltip: 'Configurar ligação',
              icon: const Icon(Icons.settings_ethernet_rounded),
            ),
          ],
          bottom: configured
              ? TabBar(
                  controller: tabs,
                  tabs: const [
                    Tab(text: 'Produtos'),
                    Tab(text: 'Matérias-primas'),
                  ],
                )
              : null,
        ),
        body: !configured
            ? _NotConfigured(onConfigure: _configure)
            : Column(children: [
                if (summary != null) _SummaryBar(summary!),
                Padding(
                  padding: const EdgeInsets.fromLTRB(18, 12, 18, 10),
                  child: TextField(
                    controller: search,
                    onChanged: _searchChanged,
                    decoration: const InputDecoration(
                      prefixIcon: Icon(Icons.search_rounded),
                      hintText: 'Código, descrição, categoria ou localização',
                    ),
                  ),
                ),
                Expanded(
                  child: loading
                      ? const Center(child: CircularProgressIndicator())
                      : error != null
                          ? _LoadError(message: error!, onRetry: _load)
                          : TabBarView(
                              controller: tabs,
                              children: [
                                _StockList(
                                    items: products,
                                    pickMode: widget.pickMode,
                                    onRefresh: _load,
                                    onSelected: _select),
                                _StockList(
                                    items: materials,
                                    pickMode: widget.pickMode,
                                    onRefresh: _load,
                                    onSelected: _select),
                              ],
                            ),
                ),
              ]),
      );

  void _select(StockItem item) {
    if (widget.pickMode) Navigator.pop(context, item);
  }
}

class _SummaryBar extends StatelessWidget {
  const _SummaryBar(this.summary);
  final StockSummary summary;

  @override
  Widget build(BuildContext context) => Container(
        width: double.infinity,
        color: AppColors.ink,
        padding: const EdgeInsets.symmetric(horizontal: 18, vertical: 11),
        child: Text(
          '${summary.products} produtos · ${summary.materials} lotes · ${summary.critical} alertas',
          style:
              const TextStyle(color: Colors.white, fontWeight: FontWeight.w800),
        ),
      );
}

class _StockList extends StatelessWidget {
  const _StockList(
      {required this.items,
      required this.pickMode,
      required this.onRefresh,
      required this.onSelected});
  final List<StockItem> items;
  final bool pickMode;
  final Future<void> Function() onRefresh;
  final ValueChanged<StockItem> onSelected;

  @override
  Widget build(BuildContext context) {
    if (items.isEmpty) {
      return const Center(child: Text('Nenhum artigo encontrado.'));
    }
    return RefreshIndicator(
      onRefresh: onRefresh,
      child: ListView.separated(
        padding: const EdgeInsets.fromLTRB(18, 4, 18, 100),
        itemCount: items.length,
        separatorBuilder: (_, __) => const SizedBox(height: 9),
        itemBuilder: (context, index) {
          final item = items[index];
          return Card(
            child: ListTile(
              onTap: pickMode ? () => onSelected(item) : null,
              contentPadding:
                  const EdgeInsets.symmetric(horizontal: 15, vertical: 7),
              leading: CircleAvatar(
                backgroundColor:
                    item.isCritical ? AppColors.redSoft : AppColors.limeSoft,
                child: Icon(
                  item.kind == 'material'
                      ? Icons.layers_rounded
                      : Icons.inventory_2_rounded,
                  color: item.isCritical ? AppColors.red : AppColors.limeDark,
                ),
              ),
              title: Text(item.description,
                  maxLines: 2,
                  overflow: TextOverflow.ellipsis,
                  style: const TextStyle(fontWeight: FontWeight.w900)),
              subtitle: Text([
                item.code,
                if (item.location.isNotEmpty) item.location,
                if (item.salePrice > 0) money(item.salePrice),
              ].join(' · ')),
              trailing: Column(
                mainAxisAlignment: MainAxisAlignment.center,
                crossAxisAlignment: CrossAxisAlignment.end,
                children: [
                  Text('${item.available.toStringAsFixed(2)} ${item.unit}',
                      style: TextStyle(
                          fontWeight: FontWeight.w900,
                          color: item.isCritical
                              ? AppColors.red
                              : AppColors.limeDark)),
                  if (pickMode)
                    const Text('Selecionar',
                        style: TextStyle(fontSize: 10, color: AppColors.muted)),
                ],
              ),
            ),
          );
        },
      ),
    );
  }
}

class _NotConfigured extends StatelessWidget {
  const _NotConfigured({required this.onConfigure});
  final VoidCallback onConfigure;

  @override
  Widget build(BuildContext context) => Center(
        child: Padding(
          padding: const EdgeInsets.all(30),
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            const CircleAvatar(
              radius: 34,
              backgroundColor: AppColors.limeSoft,
              child: Icon(Icons.cloud_off_rounded,
                  color: AppColors.limeDark, size: 34),
            ),
            const SizedBox(height: 16),
            const Text('Ligar ao stock do LuGEST',
                textAlign: TextAlign.center,
                style: TextStyle(fontSize: 21, fontWeight: FontWeight.w900)),
            const SizedBox(height: 7),
            const Text(
              'Introduza o endereço indicado pelo computador onde o LuGEST está instalado e a chave de acesso móvel.',
              textAlign: TextAlign.center,
              style: TextStyle(color: AppColors.muted, height: 1.45),
            ),
            const SizedBox(height: 18),
            FilledButton.icon(
              onPressed: onConfigure,
              icon: const Icon(Icons.link_rounded),
              label: const Text('Configurar ligação'),
            ),
          ]),
        ),
      );
}

class _LoadError extends StatelessWidget {
  const _LoadError({required this.message, required this.onRetry});
  final String message;
  final VoidCallback onRetry;

  @override
  Widget build(BuildContext context) => Center(
        child: Padding(
          padding: const EdgeInsets.all(26),
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            const Icon(Icons.wifi_off_rounded, color: AppColors.red, size: 38),
            const SizedBox(height: 10),
            Text(message, textAlign: TextAlign.center),
            const SizedBox(height: 14),
            OutlinedButton.icon(
                onPressed: onRetry,
                icon: const Icon(Icons.refresh_rounded),
                label: const Text('Tentar novamente')),
          ]),
        ),
      );
}

class _ConnectionDialog extends StatefulWidget {
  const _ConnectionDialog({required this.repository, required this.initial});
  final StockRepository repository;
  final StockApiSettings initial;

  @override
  State<_ConnectionDialog> createState() => _ConnectionDialogState();
}

class _ConnectionDialogState extends State<_ConnectionDialog> {
  late final url = TextEditingController(text: widget.initial.baseUrl);
  late final token = TextEditingController(text: widget.initial.token);
  bool testing = false;
  String? error;

  @override
  void dispose() {
    url.dispose();
    token.dispose();
    super.dispose();
  }

  Future<void> submit() async {
    final settings = StockApiSettings(baseUrl: url.text, token: token.text);
    if (!settings.isConfigured) {
      setState(() => error = 'Preencha o endereço e a chave de acesso.');
      return;
    }
    setState(() {
      testing = true;
      error = null;
    });
    try {
      await widget.repository.testConnection(settings);
      if (mounted) Navigator.pop(context, settings);
    } on StockConnectionException catch (exception) {
      if (mounted) setState(() => error = exception.message);
    } finally {
      if (mounted) setState(() => testing = false);
    }
  }

  @override
  Widget build(BuildContext context) => AlertDialog(
        title: const Text('Ligação ao LuGEST'),
        content: SingleChildScrollView(
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            TextField(
              controller: url,
              keyboardType: TextInputType.url,
              autocorrect: false,
              decoration: const InputDecoration(
                labelText: 'Endereço do computador',
                hintText: 'https://field.seudominio.pt',
              ),
            ),
            const SizedBox(height: 10),
            TextField(
              controller: token,
              obscureText: true,
              autocorrect: false,
              decoration: const InputDecoration(labelText: 'Chave de acesso'),
            ),
            if (error != null) ...[
              const SizedBox(height: 10),
              Text(error!, style: const TextStyle(color: AppColors.red)),
            ],
          ]),
        ),
        actions: [
          TextButton(
              onPressed: testing ? null : () => Navigator.pop(context),
              child: const Text('Cancelar')),
          FilledButton(
            onPressed: testing ? null : submit,
            child: Text(testing ? 'A testar…' : 'Testar e guardar'),
          ),
        ],
      );
}
