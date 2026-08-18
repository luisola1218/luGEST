import 'package:flutter/material.dart';

import '../core/app_theme.dart';
import '../core/formatters.dart';
import '../domain/service_job.dart';

class JobFormScreen extends StatefulWidget {
  const JobFormScreen({super.key});

  @override
  State<JobFormScreen> createState() => _JobFormScreenState();
}

class _JobFormScreenState extends State<JobFormScreen> {
  final formKey = GlobalKey<FormState>();
  final title = TextEditingController();
  final client = TextEditingController();
  final phone = TextEditingController();
  final address = TextEditingController();
  final notes = TextEditingController();
  final price = TextEditingController(text: '45');
  DateTime date = DateTime.now().add(const Duration(hours: 2));

  @override
  void dispose() {
    for (final controller in [title, client, phone, address, notes, price]) {
      controller.dispose();
    }
    super.dispose();
  }

  Future<void> pickDate() async {
    final selected = await showDatePicker(
      context: context,
      initialDate: date,
      firstDate: DateTime.now().subtract(const Duration(days: 30)),
      lastDate: DateTime.now().add(const Duration(days: 730)),
      helpText: 'DATA DO SERVIÇO',
      cancelText: 'Cancelar',
      confirmText: 'Escolher',
    );
    if (selected != null && mounted) {
      setState(() => date = DateTime(
          selected.year, selected.month, selected.day, date.hour, date.minute));
    }
  }

  Future<void> pickTime() async {
    final selected = await showTimePicker(
        context: context,
        initialTime: TimeOfDay.fromDateTime(date),
        helpText: 'HORA DO SERVIÇO');
    if (selected != null && mounted) {
      setState(() => date = DateTime(
          date.year, date.month, date.day, selected.hour, selected.minute));
    }
  }

  void save() {
    if (!formKey.currentState!.validate()) return;
    final amount = double.tryParse(price.text.replaceAll(',', '.')) ?? 0;
    final now = DateTime.now();
    Navigator.pop(
      context,
      ServiceJob(
        id: 'SV-${now.year}-${now.microsecondsSinceEpoch.toString().substring(8)}',
        title: title.text.trim(),
        clientName: client.text.trim(),
        clientPhone: phone.text.trim(),
        address: address.text.trim(),
        scheduledAt: date,
        status: JobStatus.scheduled,
        notes: notes.text.trim(),
        lines: [
          ServiceLine(
              description: 'Serviço técnico',
              quantity: 1,
              unit: 'SV',
              unitPrice: amount)
        ],
      ),
    );
  }

  @override
  Widget build(BuildContext context) => Scaffold(
        appBar: AppBar(title: const Text('Novo serviço')),
        bottomNavigationBar: SafeArea(
          child: Padding(
            padding: const EdgeInsets.fromLTRB(18, 10, 18, 12),
            child: FilledButton.icon(
              onPressed: save,
              style: FilledButton.styleFrom(
                  backgroundColor: AppColors.lime,
                  foregroundColor: AppColors.ink),
              icon: const Icon(Icons.check_rounded),
              label: const Text('Guardar e agendar'),
            ),
          ),
        ),
        body: Form(
          key: formKey,
          child: ListView(
            padding: const EdgeInsets.fromLTRB(18, 4, 18, 28),
            children: [
              const Text('Informação essencial',
                  style: TextStyle(fontSize: 18, fontWeight: FontWeight.w900)),
              const SizedBox(height: 5),
              const Text(
                  'Crie o serviço em poucos segundos; pode completar materiais e fotografias no local.',
                  style: TextStyle(color: AppColors.muted, height: 1.4)),
              const SizedBox(height: 18),
              TextFormField(
                  controller: title,
                  textCapitalization: TextCapitalization.sentences,
                  decoration: const InputDecoration(
                      labelText: 'Trabalho a realizar',
                      prefixIcon: Icon(Icons.handyman_rounded)),
                  validator: _required),
              const SizedBox(height: 12),
              TextFormField(
                  controller: client,
                  textCapitalization: TextCapitalization.words,
                  decoration: const InputDecoration(
                      labelText: 'Cliente',
                      prefixIcon: Icon(Icons.person_rounded)),
                  validator: _required),
              const SizedBox(height: 12),
              TextFormField(
                  controller: phone,
                  keyboardType: TextInputType.phone,
                  decoration: const InputDecoration(
                      labelText: 'Telefone',
                      prefixIcon: Icon(Icons.phone_rounded))),
              const SizedBox(height: 12),
              TextFormField(
                  controller: address,
                  textCapitalization: TextCapitalization.words,
                  decoration: const InputDecoration(
                      labelText: 'Morada do serviço',
                      prefixIcon: Icon(Icons.location_on_rounded)),
                  validator: _required),
              const SizedBox(height: 12),
              Row(children: [
                Expanded(
                    child: _PickerField(
                        label: 'Data',
                        value: fullDate(date),
                        icon: Icons.calendar_month_rounded,
                        onTap: pickDate)),
                const SizedBox(width: 10),
                Expanded(
                    child: _PickerField(
                        label: 'Hora',
                        value: clock(date),
                        icon: Icons.schedule_rounded,
                        onTap: pickTime)),
              ]),
              const SizedBox(height: 12),
              TextFormField(
                  controller: price,
                  keyboardType:
                      const TextInputType.numberWithOptions(decimal: true),
                  decoration: const InputDecoration(
                      labelText: 'Valor previsto sem IVA',
                      prefixIcon: Icon(Icons.euro_rounded),
                      suffixText: 'EUR')),
              const SizedBox(height: 12),
              TextFormField(
                  controller: notes,
                  maxLines: 4,
                  textCapitalization: TextCapitalization.sentences,
                  decoration: const InputDecoration(
                      labelText: 'Notas e instruções',
                      alignLabelWithHint: true)),
            ],
          ),
        ),
      );

  static String? _required(String? value) =>
      (value ?? '').trim().isEmpty ? 'Campo obrigatório' : null;
}

class _PickerField extends StatelessWidget {
  const _PickerField(
      {required this.label,
      required this.value,
      required this.icon,
      required this.onTap});
  final String label;
  final String value;
  final IconData icon;
  final VoidCallback onTap;

  @override
  Widget build(BuildContext context) => Semantics(
        button: true,
        label: '$label, $value. Toque para alterar.',
        child: InkWell(
          onTap: onTap,
          borderRadius: BorderRadius.circular(14),
          child: InputDecorator(
            decoration: InputDecoration(
                labelText: label,
                prefixIcon: Icon(icon),
                suffixIcon: IconButton(
                    onPressed: onTap,
                    icon: const Icon(Icons.expand_more_rounded))),
            child: Text(value,
                maxLines: 1,
                overflow: TextOverflow.ellipsis,
                style: const TextStyle(fontWeight: FontWeight.w800)),
          ),
        ),
      );
}
