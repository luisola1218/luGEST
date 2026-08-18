enum JobStatus {
  scheduled,
  onRoute,
  inProgress,
  done,
  readyToInvoice,
  invoiced
}

extension JobStatusLabel on JobStatus {
  String get label => switch (this) {
        JobStatus.scheduled => 'Agendado',
        JobStatus.onRoute => 'A caminho',
        JobStatus.inProgress => 'Em execução',
        JobStatus.done => 'Concluído',
        JobStatus.readyToInvoice => 'Por faturar',
        JobStatus.invoiced => 'Faturado',
      };
}

class ServiceLine {
  const ServiceLine(
      {required this.description,
      required this.quantity,
      required this.unit,
      required this.unitPrice,
      this.kind = 'service',
      this.ref = ''});

  final String description;
  final double quantity;
  final String unit;
  final double unitPrice;
  final String kind;
  final String ref;

  double get total => quantity * unitPrice;

  Map<String, Object> toJson() => {
        'description': description,
        'quantity': quantity,
        'unit': unit,
        'unitPrice': unitPrice,
        'kind': kind,
        'ref': ref,
      };

  factory ServiceLine.fromJson(Map<String, dynamic> json) => ServiceLine(
        description: json['description'] as String? ?? '',
        quantity: (json['quantity'] as num?)?.toDouble() ?? 0,
        unit: json['unit'] as String? ?? 'UN',
        unitPrice: (json['unitPrice'] as num?)?.toDouble() ?? 0,
        kind: json['kind'] as String? ?? 'service',
        ref: json['ref'] as String? ?? '',
      );
}

class ServiceJob {
  const ServiceJob({
    required this.id,
    required this.title,
    required this.clientName,
    required this.clientPhone,
    required this.address,
    required this.scheduledAt,
    required this.status,
    this.notes = '',
    this.lines = const [],
    this.completedSteps = const [],
    this.photoPaths = const [],
    this.signaturePath,
    this.signatureAt,
    this.vatRate = 23,
    this.isDemo = false,
    this.updatedAt,
    this.syncedAt,
  });

  final String id;
  final String title;
  final String clientName;
  final String clientPhone;
  final String address;
  final DateTime scheduledAt;
  final JobStatus status;
  final String notes;
  final List<ServiceLine> lines;
  final List<String> completedSteps;
  final List<String> photoPaths;
  final String? signaturePath;
  final DateTime? signatureAt;
  final double vatRate;
  final bool isDemo;
  final DateTime? updatedAt;
  final DateTime? syncedAt;

  bool get needsSync =>
      !isDemo &&
      (syncedAt == null ||
          (updatedAt != null && updatedAt!.isAfter(syncedAt!)));

  double get subtotal => lines.fold(0, (sum, line) => sum + line.total);
  double get vatAmount => subtotal * vatRate / 100;
  double get total => subtotal + vatAmount;

  ServiceJob copyWith({
    String? title,
    String? clientName,
    String? clientPhone,
    String? address,
    DateTime? scheduledAt,
    JobStatus? status,
    String? notes,
    List<ServiceLine>? lines,
    List<String>? completedSteps,
    List<String>? photoPaths,
    String? signaturePath,
    DateTime? signatureAt,
    bool clearSignature = false,
    double? vatRate,
    bool? isDemo,
    DateTime? updatedAt,
    DateTime? syncedAt,
    bool clearSyncedAt = false,
  }) =>
      ServiceJob(
        id: id,
        title: title ?? this.title,
        clientName: clientName ?? this.clientName,
        clientPhone: clientPhone ?? this.clientPhone,
        address: address ?? this.address,
        scheduledAt: scheduledAt ?? this.scheduledAt,
        status: status ?? this.status,
        notes: notes ?? this.notes,
        lines: lines ?? this.lines,
        completedSteps: completedSteps ?? this.completedSteps,
        photoPaths: photoPaths ?? this.photoPaths,
        signaturePath:
            clearSignature ? null : signaturePath ?? this.signaturePath,
        signatureAt: clearSignature ? null : signatureAt ?? this.signatureAt,
        vatRate: vatRate ?? this.vatRate,
        isDemo: isDemo ?? this.isDemo,
        updatedAt: updatedAt ?? this.updatedAt,
        syncedAt: clearSyncedAt ? null : syncedAt ?? this.syncedAt,
      );

  Map<String, Object?> toJson() => {
        'id': id,
        'title': title,
        'clientName': clientName,
        'clientPhone': clientPhone,
        'address': address,
        'scheduledAt': scheduledAt.toIso8601String(),
        'status': status.name,
        'notes': notes,
        'lines': lines.map((line) => line.toJson()).toList(),
        'completedSteps': completedSteps,
        'photoPaths': photoPaths,
        'signaturePath': signaturePath,
        'signatureAt': signatureAt?.toIso8601String(),
        'vatRate': vatRate,
        'isDemo': isDemo,
        'updatedAt': updatedAt?.toIso8601String(),
        'syncedAt': syncedAt?.toIso8601String(),
      };

  factory ServiceJob.fromJson(Map<String, dynamic> json) => ServiceJob(
        id: json['id'] as String? ?? '',
        title: json['title'] as String? ?? '',
        clientName: json['clientName'] as String? ?? '',
        clientPhone: json['clientPhone'] as String? ?? '',
        address: json['address'] as String? ?? '',
        scheduledAt: DateTime.tryParse(json['scheduledAt'] as String? ?? '') ??
            DateTime.now(),
        status: JobStatus.values.firstWhere(
          (value) => value.name == json['status'],
          orElse: () => JobStatus.scheduled,
        ),
        notes: json['notes'] as String? ?? '',
        lines: (json['lines'] as List<dynamic>? ?? [])
            .map((item) => ServiceLine.fromJson(item as Map<String, dynamic>))
            .toList(),
        completedSteps: List<String>.from(
            json['completedSteps'] as List<dynamic>? ?? const []),
        photoPaths:
            List<String>.from(json['photoPaths'] as List<dynamic>? ?? const []),
        signaturePath: json['signaturePath'] as String?,
        signatureAt: DateTime.tryParse(json['signatureAt'] as String? ?? ''),
        vatRate: (json['vatRate'] as num?)?.toDouble() ?? 23,
        isDemo: json['isDemo'] as bool? ?? false,
        updatedAt: DateTime.tryParse(json['updatedAt'] as String? ?? ''),
        syncedAt: DateTime.tryParse(json['syncedAt'] as String? ?? ''),
      );
}
