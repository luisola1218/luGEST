class PulseSummary {
  PulseSummary({
    required this.oee,
    required this.disponibilidade,
    required this.performance,
    required this.qualidade,
    required this.paragensMin,
    required this.pecasEmCurso,
    required this.pecasForaTempo,
    required this.desvioMaxMin,
    required this.alerts,
    required this.andonProd,
    required this.andonSetup,
    required this.andonEspera,
    required this.andonStop,
    required this.qualityScope,
  });

  final double oee;
  final double disponibilidade;
  final double performance;
  final double qualidade;
  final double paragensMin;
  final int pecasEmCurso;
  final int pecasForaTempo;
  final double desvioMaxMin;
  final String alerts;
  final int andonProd;
  final int andonSetup;
  final int andonEspera;
  final int andonStop;
  final String qualityScope;

  factory PulseSummary.fromJson(Map<String, dynamic> json) {
    final andon = json['andon'] as Map<String, dynamic>? ?? const <String, dynamic>{};
    return PulseSummary(
      oee: (json['oee'] as num?)?.toDouble() ?? 0,
      disponibilidade: (json['disponibilidade'] as num?)?.toDouble() ?? 0,
      performance: (json['performance'] as num?)?.toDouble() ?? 0,
      qualidade: (json['qualidade'] as num?)?.toDouble() ?? 0,
      paragensMin: (json['paragens_min'] as num?)?.toDouble() ?? 0,
      pecasEmCurso: (json['pecas_em_curso'] as num?)?.toInt() ?? 0,
      pecasForaTempo: (json['pecas_fora_tempo'] as num?)?.toInt() ?? 0,
      desvioMaxMin: (json['desvio_max_min'] as num?)?.toDouble() ?? 0,
      alerts: json['alerts']?.toString() ?? '',
      andonProd: (andon['prod'] as num?)?.toInt() ?? 0,
      andonSetup: (andon['setup'] as num?)?.toInt() ?? 0,
      andonEspera: (andon['espera'] as num?)?.toInt() ?? 0,
      andonStop: (andon['stop'] as num?)?.toInt() ?? 0,
      qualityScope: json['quality_scope']?.toString() ?? '',
    );
  }
}

class EncomendaItem {
  EncomendaItem({
    required this.numero,
    required this.cliente,
    required this.clienteNome,
    required this.estado,
    required this.dataEntrega,
    required this.notaCliente,
    required this.tempoH,
    required this.producaoResumo,
    required this.pecasEmCurso,
    required this.pecasEmPausa,
    required this.pecasEmAvaria,
    required this.opsAtivas,
    required this.opsPausadas,
    required this.opsAvaria,
  });

  final String numero;
  final String cliente;
  final String clienteNome;
  final String estado;
  final String dataEntrega;
  final String notaCliente;
  final double tempoH;
  final String producaoResumo;
  final int pecasEmCurso;
  final int pecasEmPausa;
  final int pecasEmAvaria;
  final String opsAtivas;
  final String opsPausadas;
  final String opsAvaria;

  factory EncomendaItem.fromJson(Map<String, dynamic> json) {
    return EncomendaItem(
      numero: json['numero']?.toString() ?? '',
      cliente: json['cliente']?.toString() ?? '',
      clienteNome: json['cliente_nome']?.toString() ?? '',
      estado: json['estado']?.toString() ?? '',
      dataEntrega: json['data_entrega']?.toString() ?? '',
      notaCliente: json['nota_cliente']?.toString() ?? '',
      tempoH: (json['tempo_h'] as num?)?.toDouble() ?? 0,
      producaoResumo: json['producao_resumo']?.toString() ?? '',
      pecasEmCurso: (json['pecas_em_curso'] as num?)?.toInt() ?? 0,
      pecasEmPausa: (json['pecas_em_pausa'] as num?)?.toInt() ?? 0,
      pecasEmAvaria: (json['pecas_em_avaria'] as num?)?.toInt() ?? 0,
      opsAtivas: json['ops_ativas']?.toString() ?? '',
      opsPausadas: json['ops_pausadas']?.toString() ?? '',
      opsAvaria: json['ops_avaria']?.toString() ?? '',
    );
  }
}
