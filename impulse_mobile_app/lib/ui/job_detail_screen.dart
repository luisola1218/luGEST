import 'dart:io';
import 'dart:typed_data';

import 'package:flutter/material.dart';

import '../core/app_theme.dart';
import '../core/formatters.dart';
import '../data/attachment_repository.dart';
import '../data/stock_repository.dart';
import '../domain/service_job.dart';
import '../domain/stock_item.dart';
import '../state/service_store.dart';
import 'components.dart';
import 'signature_screen.dart';
import 'stock_screen.dart';

class JobDetailScreen extends StatefulWidget {
  const JobDetailScreen({super.key, required this.store, required this.jobId});

  final ServiceStore store;
  final String jobId;

  @override
  State<JobDetailScreen> createState() => _JobDetailScreenState();
}

class _JobDetailScreenState extends State<JobDetailScreen> {
  final attachments = AttachmentRepository();
  bool handlingAttachment = false;

  static const steps = [
    'Chegada confirmada',
    'Diagnóstico inicial',
    'Trabalho executado',
    'Assinatura do cliente'
  ];

  @override
  void initState() {
    super.initState();
    WidgetsBinding.instance.addPostFrameCallback((_) => _recoverLostPhotos());
  }

  Future<void> _recoverLostPhotos() async {
    try {
      final paths = await attachments.recoverLostPhotos(widget.jobId);
      if (paths.isNotEmpty) await widget.store.addPhotos(widget.jobId, paths);
    } catch (_) {
      // Não interrompe a abertura do serviço se não houver dados perdidos.
    }
  }

  @override
  Widget build(BuildContext context) => AnimatedBuilder(
        animation: widget.store,
        builder: (context, _) {
          final job = widget.store.byId(widget.jobId);
          if (job == null) {
            return const Scaffold(
                body: Center(child: Text('Serviço não encontrado.')));
          }
          return Scaffold(
            appBar: AppBar(
              title: Text(job.id),
              actions: [
                IconButton(
                    onPressed: () => _message(context,
                        'Edição disponível enquanto o serviço não estiver faturado.'),
                    icon: const Icon(Icons.more_horiz_rounded))
              ],
            ),
            bottomNavigationBar: _BottomAction(
              job: job,
              store: widget.store,
              onShowAccount: () => _showAccount(job),
            ),
            body: ListView(
              padding: const EdgeInsets.fromLTRB(18, 4, 18, 120),
              children: [
                Row(children: [
                  Expanded(
                      child: Text(job.title,
                          style: const TextStyle(
                              fontSize: 24, fontWeight: FontWeight.w900))),
                  StatusPill(job.status)
                ]),
                const SizedBox(height: 16),
                Card(
                  child: Padding(
                    padding: const EdgeInsets.all(17),
                    child: Column(children: [
                      _InfoRow(
                          icon: Icons.person_rounded,
                          title: job.clientName,
                          subtitle: job.clientPhone,
                          onTap: () => _message(context,
                              'A iniciar chamada para ${job.clientPhone}')),
                      const Divider(height: 24),
                      _InfoRow(
                          icon: Icons.location_on_rounded,
                          title: job.address,
                          subtitle: 'Abrir rota',
                          onTap: () => _message(context,
                              'A abrir navegação para ${job.address}')),
                      const Divider(height: 24),
                      _InfoRow(
                          icon: Icons.event_rounded,
                          title:
                              '${fullDate(job.scheduledAt)} · ${clock(job.scheduledAt)}',
                          subtitle: 'Data e hora do serviço'),
                    ]),
                  ),
                ),
                const SizedBox(height: 20),
                const SectionHeader('Execução no terreno'),
                const SizedBox(height: 10),
                Card(
                  child: Padding(
                    padding: const EdgeInsets.symmetric(vertical: 8),
                    child: Column(
                      children: steps.map((step) {
                        final checked = job.completedSteps.contains(step);
                        return CheckboxListTile(
                          value: checked,
                          activeColor: AppColors.limeDark,
                          controlAffinity: ListTileControlAffinity.leading,
                          title: Text(step,
                              style: TextStyle(
                                  fontWeight: FontWeight.w800,
                                  decoration: checked
                                      ? TextDecoration.lineThrough
                                      : null)),
                          onChanged: job.status == JobStatus.invoiced
                              ? null
                              : (_) {
                                  if (step == 'Assinatura do cliente') {
                                    job.signaturePath == null
                                        ? _captureSignature(job)
                                        : _removeSignature(job);
                                  } else {
                                    widget.store.toggleStep(job.id, step);
                                  }
                                },
                        );
                      }).toList(),
                    ),
                  ),
                ),
                const SizedBox(height: 12),
                Row(children: [
                  Expanded(
                      child: OutlinedButton.icon(
                          onPressed: handlingAttachment
                              ? null
                              : () => _choosePhotoSource(job),
                          icon: const Icon(Icons.photo_camera_rounded),
                          label: Text(job.photoPaths.isEmpty
                              ? 'Fotografias'
                              : 'Fotos · ${job.photoPaths.length}'))),
                  const SizedBox(width: 10),
                  Expanded(
                      child: OutlinedButton.icon(
                          onPressed: handlingAttachment
                              ? null
                              : () => _captureSignature(job),
                          icon: const Icon(Icons.draw_rounded),
                          label: Text(job.signaturePath == null
                              ? 'Assinatura'
                              : 'Assinado'))),
                ]),
                if (handlingAttachment) ...[
                  const SizedBox(height: 10),
                  const LinearProgressIndicator(),
                ],
                if (job.photoPaths.isNotEmpty) ...[
                  const SizedBox(height: 14),
                  _PhotoStrip(
                    paths: job.photoPaths,
                    onRemove: (path) => _removePhoto(job, path),
                  ),
                ],
                if (job.signaturePath != null) ...[
                  const SizedBox(height: 14),
                  _SignatureCard(
                    path: job.signaturePath!,
                    signedAt: job.signatureAt,
                    onReplace: () => _captureSignature(job),
                    onRemove: () => _removeSignature(job),
                  ),
                ],
                const SizedBox(height: 22),
                SectionHeader(
                  'Trabalho e materiais',
                  action: job.status == JobStatus.invoiced ? null : 'Adicionar',
                  onAction: job.status == JobStatus.invoiced
                      ? null
                      : () => _addLine(job),
                ),
                const SizedBox(height: 10),
                Card(
                  child: Padding(
                    padding: const EdgeInsets.all(17),
                    child: Column(children: [
                      ...job.lines.asMap().entries.map((entry) => Padding(
                            padding: const EdgeInsets.only(bottom: 13),
                            child: Row(
                                crossAxisAlignment: CrossAxisAlignment.start,
                                children: [
                                  if (job.status != JobStatus.invoiced)
                                    IconButton(
                                      visualDensity: VisualDensity.compact,
                                      tooltip: 'Remover linha',
                                      onPressed: () =>
                                          _removeLine(job, entry.key),
                                      icon: const Icon(
                                          Icons.remove_circle_outline,
                                          color: AppColors.muted,
                                          size: 19),
                                    )
                                  else ...[
                                    Container(
                                        width: 8,
                                        height: 8,
                                        margin: const EdgeInsets.only(top: 5),
                                        decoration: const BoxDecoration(
                                            color: AppColors.lime,
                                            shape: BoxShape.circle)),
                                    const SizedBox(width: 10),
                                  ],
                                  Expanded(
                                      child: Column(
                                          crossAxisAlignment:
                                              CrossAxisAlignment.start,
                                          children: [
                                        Text(entry.value.description,
                                            style: const TextStyle(
                                                fontWeight: FontWeight.w800)),
                                        Text(
                                            '${entry.value.quantity} ${entry.value.unit} × ${money(entry.value.unitPrice)}',
                                            style: const TextStyle(
                                                fontSize: 12,
                                                color: AppColors.muted)),
                                      ])),
                                  Text(money(entry.value.total),
                                      style: const TextStyle(
                                          fontWeight: FontWeight.w900)),
                                ]),
                          )),
                      const Divider(),
                      _Amount(label: 'Base', value: money(job.subtotal)),
                      _Amount(
                          label:
                              'IVA ${job.vatRate.toStringAsFixed(job.vatRate % 1 == 0 ? 0 : 1)}%',
                          value: money(job.vatAmount)),
                      const SizedBox(height: 7),
                      _Amount(
                          label: 'Total',
                          value: money(job.total),
                          strong: true),
                    ]),
                  ),
                ),
                if (job.notes.isNotEmpty) ...[
                  const SizedBox(height: 20),
                  const SectionHeader('Notas'),
                  const SizedBox(height: 8),
                  Card(
                      child: Padding(
                          padding: const EdgeInsets.all(17),
                          child: Align(
                              alignment: Alignment.centerLeft,
                              child: Text(job.notes,
                                  style: const TextStyle(height: 1.45))))),
                ],
              ],
            ),
          );
        },
      );

  Future<void> _choosePhotoSource(ServiceJob job) async {
    final source = await showModalBottomSheet<String>(
      context: context,
      showDragHandle: true,
      builder: (context) => SafeArea(
        child: Padding(
          padding: const EdgeInsets.fromLTRB(18, 0, 18, 18),
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            const Align(
              alignment: Alignment.centerLeft,
              child: Text('Adicionar fotografias',
                  style: TextStyle(fontSize: 20, fontWeight: FontWeight.w900)),
            ),
            const SizedBox(height: 10),
            ListTile(
              leading: const Icon(Icons.photo_camera_rounded),
              title: const Text('Tirar fotografia'),
              subtitle: const Text('Abrir a câmara do telemóvel'),
              onTap: () => Navigator.pop(context, 'camera'),
            ),
            ListTile(
              leading: const Icon(Icons.photo_library_rounded),
              title: const Text('Escolher da galeria'),
              subtitle: const Text('Pode selecionar várias imagens'),
              onTap: () => Navigator.pop(context, 'gallery'),
            ),
          ]),
        ),
      ),
    );
    if (source == null || !mounted) return;
    setState(() => handlingAttachment = true);
    try {
      final paths = source == 'camera'
          ? [if (await attachments.capturePhoto(job.id) case final path?) path]
          : await attachments.selectPhotos(job.id);
      if (paths.isNotEmpty) {
        await widget.store.addPhotos(job.id, paths);
        if (mounted) {
          _message(context, '${paths.length} fotografia(s) guardada(s).');
        }
      }
    } catch (error) {
      if (mounted) {
        _message(context, 'Não foi possível abrir as fotografias: $error');
      }
    } finally {
      if (mounted) setState(() => handlingAttachment = false);
    }
  }

  Future<void> _captureSignature(ServiceJob job) async {
    final bytes = await Navigator.push<Uint8List>(
      context,
      MaterialPageRoute<Uint8List>(
        builder: (_) => SignatureScreen(clientName: job.clientName),
      ),
    );
    if (bytes == null || !mounted) return;
    setState(() => handlingAttachment = true);
    try {
      final oldPath = job.signaturePath;
      final path = await attachments.saveSignature(job.id, bytes);
      await widget.store.setSignature(job.id, path);
      if (oldPath != null && oldPath != path) await attachments.delete(oldPath);
      if (mounted) _message(context, 'Assinatura guardada no serviço.');
    } finally {
      if (mounted) setState(() => handlingAttachment = false);
    }
  }

  Future<void> _removePhoto(ServiceJob job, String path) async {
    final confirmed = await _confirm(
      'Remover fotografia?',
      'A fotografia será retirada definitivamente deste serviço.',
    );
    if (!confirmed) return;
    await widget.store.removePhoto(job.id, path);
    await attachments.delete(path);
  }

  Future<void> _removeSignature(ServiceJob job) async {
    final confirmed = await _confirm(
      'Remover assinatura?',
      'Será necessário pedir uma nova assinatura ao cliente.',
    );
    if (!confirmed) return;
    final path = job.signaturePath;
    await widget.store.clearSignature(job.id);
    await attachments.delete(path);
  }

  Future<bool> _confirm(String title, String message) async =>
      await showDialog<bool>(
        context: context,
        builder: (context) => AlertDialog(
          title: Text(title),
          content: Text(message),
          actions: [
            TextButton(
                onPressed: () => Navigator.pop(context, false),
                child: const Text('Cancelar')),
            FilledButton(
                onPressed: () => Navigator.pop(context, true),
                child: const Text('Remover')),
          ],
        ),
      ) ??
      false;

  Future<void> _addLine(ServiceJob job) async {
    final source = await showModalBottomSheet<String>(
      context: context,
      showDragHandle: true,
      builder: (context) => SafeArea(
        child: Padding(
          padding: const EdgeInsets.fromLTRB(18, 0, 18, 18),
          child: Column(mainAxisSize: MainAxisSize.min, children: [
            const Align(
              alignment: Alignment.centerLeft,
              child: Text('Adicionar à conta',
                  style: TextStyle(fontSize: 20, fontWeight: FontWeight.w900)),
            ),
            const SizedBox(height: 8),
            ListTile(
              leading: const Icon(Icons.edit_note_rounded),
              title: const Text('Inserir manualmente'),
              subtitle:
                  const Text('Trabalho, deslocação, horas ou outro custo'),
              onTap: () => Navigator.pop(context, 'manual'),
            ),
            ListTile(
              leading: const Icon(Icons.inventory_2_rounded),
              title: const Text('Escolher do stock LuGEST'),
              subtitle: const Text('Usar descrição e preço do software'),
              onTap: () => Navigator.pop(context, 'stock'),
            ),
          ]),
        ),
      ),
    );
    if (source == null || !mounted) return;
    ServiceLine? line;
    if (source == 'stock') {
      final item = await Navigator.push<StockItem>(
        context,
        MaterialPageRoute<StockItem>(
          builder: (_) => StockScreen(
            repository: StockRepository(),
            pickMode: true,
          ),
        ),
      );
      if (item != null && mounted) {
        line = await showDialog<ServiceLine>(
          context: context,
          builder: (_) => _LineDialog(stockItem: item),
        );
      }
    } else {
      line = await showDialog<ServiceLine>(
        context: context,
        builder: (_) => const _LineDialog(),
      );
    }
    if (line != null) await widget.store.setLines(job.id, [...job.lines, line]);
  }

  Future<void> _showAccount(ServiceJob job) async {
    await showModalBottomSheet<void>(
      context: context,
      isScrollControlled: true,
      showDragHandle: true,
      builder: (context) => _AccountSheet(
        job: job,
        onVatChanged: (rate) => widget.store.setVatRate(job.id, rate),
      ),
    );
  }

  Future<void> _removeLine(ServiceJob job, int index) async {
    final confirmed = await _confirm(
      'Remover da conta?',
      job.lines[index].description,
    );
    if (!confirmed) return;
    final lines = [...job.lines]..removeAt(index);
    await widget.store.setLines(job.id, lines);
  }

  static void _message(BuildContext context, String message) =>
      ScaffoldMessenger.of(context)
          .showSnackBar(SnackBar(content: Text(message)));
}

class _BottomAction extends StatelessWidget {
  const _BottomAction(
      {required this.job, required this.store, required this.onShowAccount});
  final ServiceJob job;
  final ServiceStore store;
  final VoidCallback onShowAccount;

  @override
  Widget build(BuildContext context) {
    final (label, icon, next) = switch (job.status) {
      JobStatus.scheduled => (
          'Iniciar deslocação',
          Icons.directions_car_rounded,
          JobStatus.onRoute
        ),
      JobStatus.onRoute => (
          'Cheguei · iniciar serviço',
          Icons.play_arrow_rounded,
          JobStatus.inProgress
        ),
      JobStatus.inProgress => (
          'Concluir trabalho',
          Icons.task_alt_rounded,
          JobStatus.done
        ),
      JobStatus.done => (
          'Pronto para faturar',
          Icons.receipt_long_rounded,
          JobStatus.readyToInvoice
        ),
      JobStatus.readyToInvoice => (
          'Sincronizar com LuGEST',
          Icons.cloud_upload_rounded,
          JobStatus.readyToInvoice
        ),
      JobStatus.invoiced => (
          'Serviço faturado',
          Icons.verified_rounded,
          JobStatus.invoiced
        ),
    };
    return SafeArea(
      child: Container(
        padding: const EdgeInsets.fromLTRB(18, 12, 18, 12),
        decoration: const BoxDecoration(
            color: Colors.white,
            border: Border(top: BorderSide(color: AppColors.border))),
        child: Row(children: [
          IconButton.filledTonal(
            onPressed: onShowAccount,
            tooltip: 'Mostrar conta ao cliente',
            icon: const Icon(Icons.request_quote_rounded),
          ),
          const SizedBox(width: 10),
          Expanded(
            child: FilledButton.icon(
              onPressed: job.status == JobStatus.invoiced
                  ? onShowAccount
                  : job.status == JobStatus.readyToInvoice
                      ? () async {
                          await store.sync();
                          if (!context.mounted) return;
                          ScaffoldMessenger.of(context).showSnackBar(SnackBar(
                            content: Text(store.syncError ??
                                'Serviço enviado ao LuGEST como rascunho.'),
                            backgroundColor: store.syncError == null
                                ? AppColors.ink
                                : AppColors.red,
                          ));
                        }
                      : () => store.setStatus(job.id, next),
              style: FilledButton.styleFrom(
                  backgroundColor: AppColors.lime,
                  foregroundColor: AppColors.ink),
              icon: Icon(job.status == JobStatus.invoiced
                  ? Icons.request_quote_rounded
                  : icon),
              label: Text(job.status == JobStatus.invoiced
                  ? 'Ver conta do serviço'
                  : label),
            ),
          ),
        ]),
      ),
    );
  }
}

class _PhotoStrip extends StatelessWidget {
  const _PhotoStrip({required this.paths, required this.onRemove});
  final List<String> paths;
  final ValueChanged<String> onRemove;

  @override
  Widget build(BuildContext context) => SizedBox(
        height: 104,
        child: ListView.separated(
          scrollDirection: Axis.horizontal,
          itemCount: paths.length,
          separatorBuilder: (_, __) => const SizedBox(width: 9),
          itemBuilder: (context, index) {
            final path = paths[index];
            return Stack(children: [
              GestureDetector(
                onTap: () => showDialog<void>(
                  context: context,
                  builder: (_) => Dialog(
                    insetPadding: const EdgeInsets.all(16),
                    child: InteractiveViewer(child: Image.file(File(path))),
                  ),
                ),
                child: ClipRRect(
                  borderRadius: BorderRadius.circular(14),
                  child: Image.file(
                    File(path),
                    width: 112,
                    height: 104,
                    fit: BoxFit.cover,
                    errorBuilder: (_, __, ___) => const SizedBox(
                      width: 112,
                      child: ColoredBox(
                        color: AppColors.redSoft,
                        child: Icon(Icons.broken_image_outlined),
                      ),
                    ),
                  ),
                ),
              ),
              Positioned(
                right: 4,
                top: 4,
                child: IconButton.filled(
                  visualDensity: VisualDensity.compact,
                  iconSize: 16,
                  onPressed: () => onRemove(path),
                  icon: const Icon(Icons.close_rounded),
                ),
              ),
            ]);
          },
        ),
      );
}

class _SignatureCard extends StatelessWidget {
  const _SignatureCard({
    required this.path,
    required this.signedAt,
    required this.onReplace,
    required this.onRemove,
  });
  final String path;
  final DateTime? signedAt;
  final VoidCallback onReplace;
  final VoidCallback onRemove;

  @override
  Widget build(BuildContext context) => Card(
        color: AppColors.limeSoft,
        child: Padding(
          padding: const EdgeInsets.all(12),
          child: Row(children: [
            ClipRRect(
              borderRadius: BorderRadius.circular(10),
              child: Image.file(
                File(path),
                width: 92,
                height: 58,
                fit: BoxFit.contain,
                errorBuilder: (_, __, ___) => const SizedBox(
                    width: 92, height: 58, child: Icon(Icons.draw_outlined)),
              ),
            ),
            const SizedBox(width: 12),
            Expanded(
              child: Column(
                crossAxisAlignment: CrossAxisAlignment.start,
                children: [
                  const Text('Cliente assinou',
                      style: TextStyle(fontWeight: FontWeight.w900)),
                  Text(
                    signedAt == null
                        ? 'Confirmação guardada'
                        : '${fullDate(signedAt!)} · ${clock(signedAt!)}',
                    style:
                        const TextStyle(color: AppColors.muted, fontSize: 12),
                  ),
                ],
              ),
            ),
            PopupMenuButton<String>(
              onSelected: (value) =>
                  value == 'replace' ? onReplace() : onRemove(),
              itemBuilder: (_) => const [
                PopupMenuItem(value: 'replace', child: Text('Substituir')),
                PopupMenuItem(value: 'remove', child: Text('Remover')),
              ],
            ),
          ]),
        ),
      );
}

class _LineDialog extends StatefulWidget {
  const _LineDialog({this.stockItem});
  final StockItem? stockItem;

  @override
  State<_LineDialog> createState() => _LineDialogState();
}

class _LineDialogState extends State<_LineDialog> {
  final formKey = GlobalKey<FormState>();
  late final description =
      TextEditingController(text: widget.stockItem?.description ?? '');
  final quantity = TextEditingController(text: '1');
  late final price = TextEditingController(
      text: widget.stockItem == null
          ? ''
          : widget.stockItem!.salePrice.toStringAsFixed(2));
  late String unit = widget.stockItem?.unit ?? 'UN';

  @override
  void dispose() {
    description.dispose();
    quantity.dispose();
    price.dispose();
    super.dispose();
  }

  double _number(String value) =>
      double.tryParse(value.trim().replaceAll(',', '.')) ?? 0;

  void submit() {
    if (!formKey.currentState!.validate()) return;
    Navigator.pop(
      context,
      ServiceLine(
        description: description.text.trim(),
        quantity: _number(quantity.text),
        unit: unit,
        unitPrice: _number(price.text),
        kind: widget.stockItem?.kind ?? 'service',
        ref: widget.stockItem?.code ?? '',
      ),
    );
  }

  @override
  Widget build(BuildContext context) => AlertDialog(
        title: const Text('Adicionar à conta'),
        content: Form(
          key: formKey,
          child: SingleChildScrollView(
            child: Column(mainAxisSize: MainAxisSize.min, children: [
              TextFormField(
                controller: description,
                autofocus: true,
                textCapitalization: TextCapitalization.sentences,
                decoration: const InputDecoration(labelText: 'Descrição'),
                validator: (value) => (value ?? '').trim().isEmpty
                    ? 'Indique o trabalho ou material'
                    : null,
              ),
              const SizedBox(height: 10),
              Row(children: [
                Expanded(
                  child: TextFormField(
                    controller: quantity,
                    keyboardType:
                        const TextInputType.numberWithOptions(decimal: true),
                    decoration: const InputDecoration(labelText: 'Quantidade'),
                    validator: (value) =>
                        _number(value ?? '') <= 0 ? 'Valor inválido' : null,
                  ),
                ),
                const SizedBox(width: 8),
                Expanded(
                  child: DropdownButtonFormField<String>(
                    initialValue: unit,
                    decoration: const InputDecoration(labelText: 'Unidade'),
                    items: const ['UN', 'H', 'SV', 'M', 'M2', 'KG']
                        .map((value) => DropdownMenuItem(
                              value: value,
                              child: Text(value),
                            ))
                        .toList(),
                    onChanged: (value) => unit = value ?? 'UN',
                  ),
                ),
              ]),
              const SizedBox(height: 10),
              TextFormField(
                controller: price,
                keyboardType:
                    const TextInputType.numberWithOptions(decimal: true),
                decoration: const InputDecoration(
                    labelText: 'Preço unitário', suffixText: 'EUR'),
                validator: (value) =>
                    _number(value ?? '') < 0 ? 'Valor inválido' : null,
              ),
            ]),
          ),
        ),
        actions: [
          TextButton(
              onPressed: () => Navigator.pop(context),
              child: const Text('Cancelar')),
          FilledButton(onPressed: submit, child: const Text('Adicionar')),
        ],
      );
}

class _AccountSheet extends StatefulWidget {
  const _AccountSheet({required this.job, required this.onVatChanged});
  final ServiceJob job;
  final ValueChanged<double> onVatChanged;

  @override
  State<_AccountSheet> createState() => _AccountSheetState();
}

class _AccountSheetState extends State<_AccountSheet> {
  late double vatRate = widget.job.vatRate;

  @override
  Widget build(BuildContext context) {
    final vat = widget.job.subtotal * vatRate / 100;
    final total = widget.job.subtotal + vat;
    return SafeArea(
      child: Padding(
        padding: const EdgeInsets.fromLTRB(20, 0, 20, 20),
        child: SingleChildScrollView(
          child:
              Column(crossAxisAlignment: CrossAxisAlignment.start, children: [
            const Text('Resumo para o cliente',
                style: TextStyle(fontSize: 24, fontWeight: FontWeight.w900)),
            const SizedBox(height: 4),
            Text('${widget.job.clientName} · ${widget.job.id}',
                style: const TextStyle(color: AppColors.muted)),
            const SizedBox(height: 18),
            ...widget.job.lines.map((line) => Padding(
                  padding: const EdgeInsets.only(bottom: 10),
                  child: Row(children: [
                    Expanded(
                      child: Column(
                        crossAxisAlignment: CrossAxisAlignment.start,
                        children: [
                          Text(line.description,
                              style:
                                  const TextStyle(fontWeight: FontWeight.w800)),
                          Text('${line.quantity} ${line.unit}',
                              style: const TextStyle(
                                  color: AppColors.muted, fontSize: 12)),
                        ],
                      ),
                    ),
                    Text(money(line.total),
                        style: const TextStyle(fontWeight: FontWeight.w900)),
                  ]),
                )),
            const Divider(height: 26),
            Row(children: [
              const Text('Taxa de IVA',
                  style: TextStyle(fontWeight: FontWeight.w800)),
              const Spacer(),
              DropdownButton<double>(
                value: vatRate,
                items: const [0.0, 6.0, 13.0, 23.0]
                    .map((rate) => DropdownMenuItem(
                        value: rate,
                        child: Text('${rate.toStringAsFixed(0)}%')))
                    .toList(),
                onChanged: (value) {
                  if (value == null) return;
                  setState(() => vatRate = value);
                  widget.onVatChanged(value);
                },
              ),
            ]),
            _Amount(label: 'Subtotal', value: money(widget.job.subtotal)),
            _Amount(label: 'IVA', value: money(vat)),
            const SizedBox(height: 6),
            Container(
              padding: const EdgeInsets.all(16),
              decoration: BoxDecoration(
                color: AppColors.ink,
                borderRadius: BorderRadius.circular(18),
              ),
              child: Row(children: [
                const Text('TOTAL A PAGAR',
                    style: TextStyle(
                        color: Colors.white, fontWeight: FontWeight.w900)),
                const Spacer(),
                Text(money(total),
                    style: const TextStyle(
                        color: AppColors.lime,
                        fontSize: 25,
                        fontWeight: FontWeight.w900)),
              ]),
            ),
            const SizedBox(height: 12),
            const Text(
              'Resumo de trabalho — não substitui a fatura fiscal. A emissão final é feita no LuGEST.',
              style:
                  TextStyle(color: AppColors.muted, fontSize: 11, height: 1.4),
            ),
            const SizedBox(height: 14),
            SizedBox(
              width: double.infinity,
              child: FilledButton.icon(
                onPressed: () => Navigator.pop(context),
                style: FilledButton.styleFrom(
                    backgroundColor: AppColors.lime,
                    foregroundColor: AppColors.ink),
                icon: const Icon(Icons.check_circle_rounded),
                label: const Text('Confirmar com o cliente'),
              ),
            ),
          ]),
        ),
      ),
    );
  }
}

class _InfoRow extends StatelessWidget {
  const _InfoRow(
      {required this.icon,
      required this.title,
      required this.subtitle,
      this.onTap});
  final IconData icon;
  final String title;
  final String subtitle;
  final VoidCallback? onTap;

  @override
  Widget build(BuildContext context) => InkWell(
        onTap: onTap,
        child: Row(children: [
          CircleAvatar(
              backgroundColor: AppColors.limeSoft,
              child: Icon(icon, color: AppColors.limeDark, size: 20)),
          const SizedBox(width: 12),
          Expanded(
              child: Column(
                  crossAxisAlignment: CrossAxisAlignment.start,
                  children: [
                Text(title,
                    style: const TextStyle(fontWeight: FontWeight.w900)),
                const SizedBox(height: 2),
                Text(subtitle,
                    style:
                        const TextStyle(color: AppColors.muted, fontSize: 12))
              ])),
          if (onTap != null)
            const Icon(Icons.chevron_right_rounded, color: AppColors.muted),
        ]),
      );
}

class _Amount extends StatelessWidget {
  const _Amount(
      {required this.label, required this.value, this.strong = false});
  final String label;
  final String value;
  final bool strong;

  @override
  Widget build(BuildContext context) => Padding(
        padding: const EdgeInsets.symmetric(vertical: 3),
        child: Row(children: [
          Text(label,
              style: TextStyle(
                  color: strong ? AppColors.ink : AppColors.muted,
                  fontWeight: strong ? FontWeight.w900 : FontWeight.w600)),
          const Spacer(),
          Text(value,
              style: TextStyle(
                  fontSize: strong ? 19 : 14,
                  fontWeight: FontWeight.w900,
                  color: strong ? AppColors.limeDark : AppColors.ink))
        ]),
      );
}
