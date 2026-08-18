class StockItem {
  const StockItem({
    required this.kind,
    required this.code,
    required this.description,
    required this.available,
    required this.unit,
    required this.salePrice,
    this.location = '',
    this.category = '',
    this.alertLevel = 0,
  });

  final String kind;
  final String code;
  final String description;
  final double available;
  final String unit;
  final double salePrice;
  final String location;
  final String category;
  final double alertLevel;

  bool get isCritical =>
      available <= 0 || (alertLevel > 0 && available <= alertLevel);

  factory StockItem.fromJson(Map<String, dynamic> json) => StockItem(
        kind: json['kind'] as String? ?? 'product',
        code: json['code'] as String? ?? '',
        description: json['description'] as String? ?? '',
        available: (json['available'] as num?)?.toDouble() ?? 0,
        unit: json['unit'] as String? ?? 'UN',
        salePrice: (json['sale_price'] as num?)?.toDouble() ?? 0,
        location: json['location'] as String? ?? '',
        category: json['category'] as String? ?? '',
        alertLevel: (json['alert_level'] as num?)?.toDouble() ?? 0,
      );
}

class StockSummary {
  const StockSummary({
    required this.products,
    required this.materials,
    required this.critical,
    required this.updatedAt,
  });

  final int products;
  final int materials;
  final int critical;
  final DateTime? updatedAt;

  factory StockSummary.fromJson(Map<String, dynamic> json) => StockSummary(
        products: (json['products'] as num?)?.toInt() ?? 0,
        materials: (json['materials'] as num?)?.toInt() ?? 0,
        critical: (json['critical'] as num?)?.toInt() ?? 0,
        updatedAt: DateTime.tryParse(json['updated_at'] as String? ?? ''),
      );
}
