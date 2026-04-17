import 'package:flutter/material.dart';

import '../api/api_client.dart';
import '../support/pdf_launcher.dart';
import '../widgets/mobile_ui.dart';

class MaterialSeparationScreen extends StatefulWidget {
  const MaterialSeparationScreen({super.key, required this.api});

  final ApiClient api;

  @override
  State<MaterialSeparationScreen> createState() => _MaterialSeparationScreenState();
}

class _MaterialSeparationScreenState extends State<MaterialSeparationScreen> {
  static const int _horizonDays = 4;

  bool _loading = true;
  String _error = '';
  String _activePosto = '';
  Map<String, dynamic> _data = const <String, dynamic>{};
  final Set<String> _savingKeys = <String>{};
  final TextEditingController _searchCtrl = TextEditingController();

  @override
  void initState() {
    super.initState();
    _load();
  }

  @override
  void dispose() {
    _searchCtrl.dispose();
    super.dispose();
  }

  Future<void> _openPdf() async {
    final url = widget.api.materialSeparationPdfUrl(horizonDays: _horizonDays);
    await openPdfUrl(context, url);
  }

  Future<void> _load() async {
    setState(() {
      _loading = true;
      _error = '';
    });
    try {
      final data = await widget.api.getMaterialSeparation(
        horizonDays: _horizonDays,
      );
      final postos = List<Map<String, dynamic>>.from(
        data['postos'] as List<dynamic>? ?? const [],
      );
      final stillExists = _activePosto.isNotEmpty &&
          postos.any((row) => row['posto_trabalho']?.toString() == _activePosto);
      if (!mounted) return;
      setState(() {
        _data = data;
        _activePosto = stillExists ? _activePosto : '';
        _loading = false;
      });
    } catch (exc) {
      if (!mounted) return;
      setState(() {
        _error = exc.toString().replaceFirst('Exception: ', '');
        _loading = false;
      });
    }
  }

  Future<void> _toggleCheck(Map<String, dynamic> row, bool checked) async {
    final checkKey = row['check_key']?.toString().trim() ?? '';
    if (checkKey.isEmpty || _savingKeys.contains(checkKey)) {
      return;
    }
    final previous = row['visto_sep_checked'] == true;
    setState(() {
      _savingKeys.add(checkKey);
      _applyCheckLocally(checkKey, checked);
    });
    try {
      await widget.api.setMaterialSeparationCheck(
        checkKey: checkKey,
        checked: checked,
      );
    } catch (exc) {
      if (!mounted) return;
      setState(() {
        _applyCheckLocally(checkKey, previous);
        _error = exc.toString().replaceFirst('Exception: ', '');
      });
    } finally {
      if (mounted) {
        setState(() => _savingKeys.remove(checkKey));
      }
    }
  }

  void _applyCheckLocally(String checkKey, bool checked) {
    final summary = _summary;
    for (final posto in _postos) {
      final rows = List<Map<String, dynamic>>.from(
        posto['rows'] as List<dynamic>? ?? const [],
      );
      for (final row in rows) {
        if ((row['check_key']?.toString().trim() ?? '') != checkKey) {
          continue;
        }
        final previousChecked = row['visto_sep_checked'] == true;
        if (previousChecked == checked) {
          return;
        }
        row['visto_sep_checked'] = checked;
        final delta = checked ? 1 : -1;
        posto['checked_sep'] = ((posto['checked_sep'] as num?)?.toInt() ?? 0) + delta;
        summary['checked_sep'] = ((summary['checked_sep'] as num?)?.toInt() ?? 0) + delta;
        summary['pending_sep'] = ((summary['pending_sep'] as num?)?.toInt() ?? 0) - delta;
        return;
      }
    }
  }

  List<Map<String, dynamic>> get _postos => List<Map<String, dynamic>>.from(
        _data['postos'] as List<dynamic>? ?? const [],
      );

  Map<String, dynamic> get _summary =>
      _data['summary'] as Map<String, dynamic>? ?? <String, dynamic>{};

  String get _horizonLabel =>
      _data['horizon_label']?.toString().trim().isNotEmpty == true
          ? _data['horizon_label']?.toString() ?? ''
          : '$_horizonDays dias uteis';

  List<Map<String, dynamic>> _allRowsForPosto(String postoTxt) {
    for (final posto in _postos) {
      if ((posto['posto_trabalho']?.toString() ?? '') == postoTxt) {
        return List<Map<String, dynamic>>.from(
          posto['rows'] as List<dynamic>? ?? const [],
        );
      }
    }
    return const <Map<String, dynamic>>[];
  }

  bool _matchesSearch(Map<String, dynamic> row) {
    final query = _searchCtrl.text.trim().toLowerCase();
    if (query.isEmpty) return true;
    final haystack = [
      row['numero'],
      row['cliente'],
      row['material'],
      row['espessura'],
      row['lote'],
      row['dimensao'],
      row['planeado_dia'],
      row['planeado_hora'],
    ].map((value) => value?.toString().toLowerCase() ?? '').join(' | ');
    return haystack.contains(query);
  }

  List<Map<String, dynamic>> _visibleRowsForPosto(String postoTxt) {
    return _allRowsForPosto(postoTxt).where(_matchesSearch).toList();
  }

  void _openPosto(String postoTxt) {
    setState(() {
      _activePosto = postoTxt;
      _error = '';
      _searchCtrl.clear();
    });
  }

  void _closePosto() {
    setState(() {
      _activePosto = '';
      _error = '';
      _searchCtrl.clear();
    });
  }

  @override
  Widget build(BuildContext context) {
    final activeRows = _activePosto.isEmpty
        ? const <Map<String, dynamic>>[]
        : _visibleRowsForPosto(_activePosto);
    final activeChecked =
        activeRows.where((row) => row['visto_sep_checked'] == true).length;
    return Scaffold(
      body: MobileSurface(
        child: RefreshIndicator(
          onRefresh: _load,
          child: ListView(
            physics: const AlwaysScrollableScrollPhysics(),
            children: [
              _SeparationHeader(
                title: _activePosto.isEmpty ? 'Separacao MP' : _activePosto,
                subtitle: _activePosto.isEmpty
                    ? 'Materiais por separar nos proximos $_horizonDays dias uteis.'
                    : '${activeRows.length} linhas visiveis | $activeChecked com visto verde.',
                onBack: _activePosto.isEmpty
                    ? () => Navigator.of(context).maybePop()
                    : _closePosto,
                onPdf: _openPdf,
                onRefresh: _load,
              ),
              SizedBox(height: ui(context, 14)),
              if (_loading)
                const Center(
                  child: Padding(
                    padding: EdgeInsets.all(24),
                    child: CircularProgressIndicator(),
                  ),
                )
              else if (_error.isNotEmpty && _data.isEmpty)
                EmptyStateCard(
                  title: 'Falha ao carregar separacao',
                  message: _error,
                  icon: Icons.error_outline_rounded,
                )
              else if (_activePosto.isEmpty)
                _OverviewBody(
                  postos: _postos,
                  summary: _summary,
                  horizonLabel: _horizonLabel,
                  onOpenPosto: _openPosto,
                )
              else
                _PostoDetailBody(
                  posto: _activePosto,
                  rows: activeRows,
                  error: _error,
                  searchCtrl: _searchCtrl,
                  onSearchChanged: () => setState(() {}),
                  savingKeys: _savingKeys,
                  onToggleCheck: _toggleCheck,
                ),
            ],
          ),
        ),
      ),
    );
  }
}

class _SeparationHeader extends StatelessWidget {
  const _SeparationHeader({
    required this.title,
    required this.subtitle,
    required this.onBack,
    required this.onPdf,
    required this.onRefresh,
  });

  final String title;
  final String subtitle;
  final VoidCallback onBack;
  final VoidCallback onPdf;
  final VoidCallback onRefresh;

  @override
  Widget build(BuildContext context) {
    return Column(
      crossAxisAlignment: CrossAxisAlignment.start,
      children: [
        Row(
          children: [
            IconButton.filledTonal(
              onPressed: onBack,
              icon: const Icon(Icons.arrow_back_rounded),
            ),
            SizedBox(width: ui(context, 12)),
            Expanded(
              child: Text(
                title,
                maxLines: 1,
                overflow: TextOverflow.ellipsis,
                style: Theme.of(context).textTheme.headlineSmall,
              ),
            ),
          ],
        ),
        SizedBox(height: ui(context, 8)),
        Text(
          subtitle,
          style: Theme.of(context).textTheme.bodyMedium?.copyWith(
                color: AppPalette.slate,
              ),
        ),
        SizedBox(height: ui(context, 14)),
        Row(
          children: [
            Expanded(
              child: OutlinedButton.icon(
                onPressed: onPdf,
                icon: const Icon(Icons.picture_as_pdf_rounded),
                label: const Text('PDF'),
              ),
            ),
            SizedBox(width: ui(context, 10)),
            Expanded(
              child: FilledButton.icon(
                onPressed: onRefresh,
                icon: const Icon(Icons.refresh_rounded),
                label: const Text('Atualizar'),
              ),
            ),
          ],
        ),
      ],
    );
  }
}

class _OverviewBody extends StatelessWidget {
  const _OverviewBody({
    required this.postos,
    required this.summary,
    required this.horizonLabel,
    required this.onOpenPosto,
  });

  final List<Map<String, dynamic>> postos;
  final Map<String, dynamic> summary;
  final String horizonLabel;
  final ValueChanged<String> onOpenPosto;

  @override
  Widget build(BuildContext context) {
    return Column(
      children: [
        MobileSectionCard(
          title: 'Resumo',
          subtitle: 'Visao rapida do que esta por separar.',
          child: Wrap(
            spacing: ui(context, 8),
            runSpacing: ui(context, 8),
            children: [
              SoftChip(
                label: '${summary['rows'] ?? 0} linhas',
                icon: Icons.format_list_bulleted_rounded,
              ),
              SoftChip(
                label: '${summary['checked_sep'] ?? 0} com visto',
                icon: Icons.check_circle_outline_rounded,
                color: AppPalette.green,
              ),
              SoftChip(
                label: '${summary['pending_sep'] ?? 0} pendentes',
                icon: Icons.pending_actions_rounded,
                color: AppPalette.orange,
              ),
              SoftChip(
                label: horizonLabel,
                icon: Icons.schedule_rounded,
                color: AppPalette.cobalt,
              ),
            ],
          ),
        ),
        SizedBox(height: ui(context, 12)),
        MobileSectionCard(
          title: 'Maquinas',
          subtitle: 'Entra na maquina para ver o material a separar.',
          child: postos.isEmpty
              ? const EmptyStateCard(
                  title: 'Sem linhas de separacao',
                  message: 'Nao existem necessidades ativas neste horizonte.',
                )
              : Column(
                  children: postos.map((posto) {
                    final postoTxt =
                        posto['posto_trabalho']?.toString() ?? 'Sem posto';
                    final total = (posto['rows'] as List<dynamic>? ?? const []).length;
                    final checked = ((posto['checked_sep'] as num?)?.toInt() ?? 0);
                    final pending = (total - checked).clamp(0, total);
                    return Padding(
                      padding: EdgeInsets.only(bottom: ui(context, 10)),
                      child: _PostoEntryCard(
                        posto: postoTxt,
                        total: total,
                        checked: checked,
                        pending: pending,
                        onTap: () => onOpenPosto(postoTxt),
                      ),
                    );
                  }).toList(),
                ),
        ),
      ],
    );
  }
}

class _PostoDetailBody extends StatelessWidget {
  const _PostoDetailBody({
    required this.posto,
    required this.rows,
    required this.error,
    required this.searchCtrl,
    required this.onSearchChanged,
    required this.savingKeys,
    required this.onToggleCheck,
  });

  final String posto;
  final List<Map<String, dynamic>> rows;
  final String error;
  final TextEditingController searchCtrl;
  final VoidCallback onSearchChanged;
  final Set<String> savingKeys;
  final Future<void> Function(Map<String, dynamic> row, bool checked) onToggleCheck;

  @override
  Widget build(BuildContext context) {
    final checkedVisible =
        rows.where((row) => row['visto_sep_checked'] == true).length;
    return Column(
      children: [
        MobileSectionCard(
          title: posto,
          subtitle: '${rows.length} linha(s) visiveis | $checkedVisible com visto verde.',
          child: Column(
            children: [
              Wrap(
                spacing: ui(context, 8),
                runSpacing: ui(context, 8),
                children: [
                  SoftChip(
                    label: '${rows.length} linhas',
                    icon: Icons.view_list_rounded,
                  ),
                  SoftChip(
                    label: '$checkedVisible vistas',
                    icon: Icons.check_circle_outline_rounded,
                    color: AppPalette.green,
                  ),
                ],
              ),
              SizedBox(height: ui(context, 14)),
              FilterField(
                controller: searchCtrl,
                hint: 'Pesquisar encomenda, lote, material ou formato...',
                prefixIcon: Icons.search_rounded,
                onChanged: (_) => onSearchChanged(),
              ),
            ],
          ),
        ),
        if (error.isNotEmpty)
          Padding(
            padding: EdgeInsets.only(top: ui(context, 12)),
            child: EmptyStateCard(
              title: 'Aviso',
              message: error,
              icon: Icons.info_outline_rounded,
            ),
          ),
        SizedBox(height: ui(context, 12)),
        if (rows.isEmpty)
          const EmptyStateCard(
            title: 'Sem linhas neste posto',
            message: 'Nao existem linhas para o posto e pesquisa escolhidos.',
          )
        else
          Column(
            children: rows.map((row) {
              final checkKey = row['check_key']?.toString().trim() ?? '';
              final saving = savingKeys.contains(checkKey);
              final checked = row['visto_sep_checked'] == true;
              return Padding(
                padding: EdgeInsets.only(bottom: ui(context, 10)),
                child: _SeparationRowCard(
                  row: row,
                  checked: checked,
                  saving: saving,
                  onToggle: (value) => onToggleCheck(row, value),
                ),
              );
            }).toList(),
          ),
      ],
    );
  }
}

class _PostoEntryCard extends StatelessWidget {
  const _PostoEntryCard({
    required this.posto,
    required this.total,
    required this.checked,
    required this.pending,
    required this.onTap,
  });

  final String posto;
  final int total;
  final int checked;
  final int pending;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) {
    return InkWell(
      borderRadius: BorderRadius.circular(22),
      onTap: onTap,
      child: Ink(
        decoration: BoxDecoration(
          color: AppPalette.surface,
          borderRadius: BorderRadius.circular(22),
          border: Border.all(color: AppPalette.line),
        ),
        child: Padding(
          padding: EdgeInsets.all(ui(context, 16)),
          child: Row(
            children: [
              Container(
                width: ui(context, 46),
                height: ui(context, 46),
                decoration: BoxDecoration(
                  color: AppPalette.navy.withValues(alpha: 0.08),
                  borderRadius: BorderRadius.circular(16),
                ),
                child: const Icon(
                  Icons.precision_manufacturing_rounded,
                  color: AppPalette.navy,
                ),
              ),
              SizedBox(width: ui(context, 14)),
              Expanded(
                child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                    Text(
                      posto,
                      style: Theme.of(context).textTheme.titleMedium?.copyWith(
                            fontWeight: FontWeight.w800,
                          ),
                    ),
                    SizedBox(height: ui(context, 4)),
                    Text(
                      '$pending pendentes | $checked com visto | $total linhas',
                      style: Theme.of(context).textTheme.bodySmall?.copyWith(
                            color: AppPalette.slate,
                          ),
                    ),
                  ],
                ),
              ),
              const Icon(
                Icons.chevron_right_rounded,
                color: AppPalette.slate,
              ),
            ],
          ),
        ),
      ),
    );
  }
}

class _SeparationRowCard extends StatelessWidget {
  const _SeparationRowCard({
    required this.row,
    required this.checked,
    required this.saving,
    required this.onToggle,
  });

  final Map<String, dynamic> row;
  final bool checked;
  final bool saving;
  final ValueChanged<bool> onToggle;

  @override
  Widget build(BuildContext context) {
    final priority = row['priority_label']?.toString() ?? 'Media';
    final priorityColor = _priorityColor(priority);
    final lotTxt = row['lote']?.toString().trim().isNotEmpty == true
        ? row['lote']?.toString() ?? '-'
        : '-';
    final planeadoTxt =
        '${row['planeado_dia'] ?? '-'} ${row['planeado_hora'] ?? '-'}'.trim();
    return Container(
      decoration: BoxDecoration(
        color: AppPalette.surface,
        borderRadius: BorderRadius.circular(20),
        border: Border.all(color: AppPalette.line),
      ),
      child: Padding(
        padding: EdgeInsets.all(ui(context, 14)),
        child: Column(
          crossAxisAlignment: CrossAxisAlignment.start,
          children: [
            Row(
              crossAxisAlignment: CrossAxisAlignment.start,
              children: [
                Expanded(
                  child: Column(
                    crossAxisAlignment: CrossAxisAlignment.start,
                    children: [
                      Text(
                        '${row['numero'] ?? '-'}',
                        style: Theme.of(context).textTheme.titleMedium?.copyWith(
                              fontWeight: FontWeight.w800,
                            ),
                      ),
                      SizedBox(height: ui(context, 3)),
                      Text(
                        '${row['cliente'] ?? '-'}',
                        style: Theme.of(context).textTheme.bodySmall?.copyWith(
                              color: AppPalette.slate,
                            ),
                      ),
                    ],
                  ),
                ),
                Column(
                  children: [
                    Checkbox(
                      value: checked,
                      onChanged: saving ? null : (value) => onToggle(value == true),
                      fillColor: WidgetStateProperty.resolveWith(
                        (_) => checked ? AppPalette.green : AppPalette.surface,
                      ),
                      checkColor: Colors.white,
                      side: BorderSide(
                        color: checked ? AppPalette.green : AppPalette.lineStrong,
                        width: 1.4,
                      ),
                    ),
                    Text(
                      saving ? 'A guardar' : 'Separado',
                      style: Theme.of(context).textTheme.labelSmall?.copyWith(
                            color: checked ? AppPalette.green : AppPalette.slate,
                            fontWeight: checked ? FontWeight.w700 : FontWeight.w600,
                          ),
                    ),
                  ],
                ),
              ],
            ),
            SizedBox(height: ui(context, 10)),
            Text(
              '${row['material'] ?? '-'} ${row['espessura'] ?? '-'} mm',
              style: Theme.of(context).textTheme.titleSmall?.copyWith(
                    fontWeight: FontWeight.w700,
                  ),
            ),
            SizedBox(height: ui(context, 10)),
            Wrap(
              spacing: ui(context, 8),
              runSpacing: ui(context, 8),
              children: [
                StatusPill(priority, color: priorityColor),
                SoftChip(
                  label: 'Lote $lotTxt',
                  icon: Icons.sell_outlined,
                ),
                SoftChip(
                  label: 'Formato ${row['dimensao'] ?? '-'}',
                  icon: Icons.aspect_ratio_rounded,
                ),
                SoftChip(
                  label: 'Qtd. ${_fmtQty(row['quantidade'])}',
                  icon: Icons.functions_rounded,
                ),
                SoftChip(
                  label: 'Planeado $planeadoTxt',
                  icon: Icons.schedule_rounded,
                  color: AppPalette.cobalt,
                ),
              ],
            ),
            SizedBox(height: ui(context, 10)),
            Text(
              row['acao_sugerida']?.toString() ?? '-',
              style: Theme.of(context).textTheme.bodyMedium,
            ),
            if ((row['alerta_texto']?.toString().trim().isNotEmpty ?? false)) ...[
              SizedBox(height: ui(context, 8)),
              Container(
                width: double.infinity,
                padding: EdgeInsets.all(ui(context, 10)),
                decoration: BoxDecoration(
                  color: AppPalette.orange.withValues(alpha: 0.08),
                  borderRadius: BorderRadius.circular(14),
                  border: Border.all(
                    color: AppPalette.orange.withValues(alpha: 0.18),
                  ),
                ),
                child: Text(
                  row['alerta_texto']?.toString() ?? '',
                  style: Theme.of(context).textTheme.bodySmall?.copyWith(
                        color: AppPalette.amber,
                      ),
                ),
              ),
            ],
          ],
        ),
      ),
    );
  }
}

Color _priorityColor(String priority) {
  final text = priority.trim().toLowerCase();
  if (text.contains('crit')) return AppPalette.red;
  if (text.contains('alta')) return AppPalette.orange;
  if (text.contains('media')) return AppPalette.cobalt;
  return AppPalette.green;
}

String _fmtQty(dynamic value) {
  final number = (value as num?)?.toDouble() ?? 0;
  if ((number - number.roundToDouble()).abs() < 0.001) {
    return number.toStringAsFixed(0);
  }
  return number.toStringAsFixed(2);
}
