import 'dart:convert';

import 'package:http/http.dart' as http;

import '../support/app_defaults.dart';

class ApiClient {
  ApiClient({String? baseUrl})
      : _baseUrl = _normalizeBaseUrl(baseUrl ?? configuredMobileApiBaseUrl());

  String _baseUrl;
  String? _token;
  Map<String, dynamic> _user = <String, dynamic>{};

  String get baseUrl => _baseUrl;
  Map<String, dynamic> get user => _user;
  String get username => _user['username']?.toString() ?? '';
  String get role => _user['role']?.toString() ?? '';
  String get token => _token ?? '';

  void setBaseUrl(String value) {
    _baseUrl = _normalizeBaseUrl(value);
  }

  String _requireBaseUrl() {
    final current = _baseUrl.trim();
    if (current.isEmpty) {
      throw Exception('Indica o servidor API antes de continuar.');
    }
    return current;
  }

  Future<Map<String, dynamic>> login(String username, String password) async {
    final uri = Uri.parse('${_requireBaseUrl()}/auth/login');
    final res = await http.post(
      uri,
      headers: {'Content-Type': 'application/json'},
      body: jsonEncode({'username': username, 'password': password}),
    );
    if (res.statusCode != 200) {
      throw Exception(_errorMessage(res, 'Falha no login.'));
    }
    final data = _decodeMap(res.body);
    _token = data['token'] as String?;
    _user = data['user'] as Map<String, dynamic>? ?? <String, dynamic>{};
    return data;
  }

  Future<Map<String, dynamic>> getDashboard({
    String period = '7 dias',
    String? year,
    String encomenda = 'Todas',
    String visao = 'Todas',
    String origem = 'Ambos',
  }) async {
    final effectiveYear = year ?? currentOperationalYear();
    final uri = Uri.parse('${_requireBaseUrl()}/pulse/dashboard').replace(
      queryParameters: {
        'period': period,
        'year': effectiveYear,
        'encomenda': encomenda,
        'visao': visao,
        'origem': origem,
      },
    );
    final res = await http.get(uri, headers: _headers());
    if (res.statusCode != 200) {
      throw Exception(_errorMessage(res, 'Falha a carregar dashboard.'));
    }
    return _decodeMap(res.body);
  }

  Future<List<dynamic>> getEncomendas({String? year}) async {
    final effectiveYear = year ?? currentOperationalYear();
    final uri = Uri.parse(
      '${_requireBaseUrl()}/pulse/encomendas',
    ).replace(queryParameters: {'year': effectiveYear});
    final res = await http.get(uri, headers: _headers());
    if (res.statusCode != 200) {
      throw Exception(_errorMessage(res, 'Falha a carregar encomendas.'));
    }
    final data = _decodeMap(res.body);
    return (data['items'] as List<dynamic>? ?? <dynamic>[]);
  }

  Future<Map<String, dynamic>> getEncomenda(String numero) async {
    final uri = Uri.parse('${_requireBaseUrl()}/pulse/encomendas/$numero');
    final res = await http.get(uri, headers: _headers());
    if (res.statusCode != 200) {
      throw Exception(_errorMessage(res, 'Falha a carregar encomenda.'));
    }
    final data = _decodeMap(res.body);
    return data['item'] as Map<String, dynamic>? ?? <String, dynamic>{};
  }

  Future<Map<String, dynamic>> getOperatorBoard({String? year}) async {
    final effectiveYear = year ?? currentOperationalYear();
    final uri = Uri.parse(
      '${_requireBaseUrl()}/mobile/operator-board',
    ).replace(queryParameters: {'year': effectiveYear});
    final res = await http.get(uri, headers: _headers());
    if (res.statusCode != 200) {
      throw Exception(_errorMessage(res, 'Falha a carregar operador.'));
    }
    return _decodeMap(res.body);
  }

  Future<Map<String, dynamic>> getMobileAlerts() async {
    final uri = Uri.parse('${_requireBaseUrl()}/mobile/alerts');
    final res = await http.get(uri, headers: _headers());
    if (res.statusCode != 200) {
      throw Exception(_errorMessage(res, 'Falha a carregar alertas.'));
    }
    return _decodeMap(res.body);
  }

  Future<Map<String, dynamic>> getAvarias() async {
    final uri = Uri.parse('${_requireBaseUrl()}/mobile/avarias');
    final res = await http.get(uri, headers: _headers());
    if (res.statusCode != 200) {
      throw Exception(_errorMessage(res, 'Falha a carregar avarias.'));
    }
    return _decodeMap(res.body);
  }

  Future<Map<String, dynamic>> getPlanningOverview({
    String? year,
    String? weekStart,
  }) async {
    final effectiveYear = year ?? currentOperationalYear();
    final uri = Uri.parse('${_requireBaseUrl()}/mobile/planning').replace(
      queryParameters: {
        'year': effectiveYear,
        if (weekStart != null && weekStart.isNotEmpty) 'week_start': weekStart,
      },
    );
    final res = await http.get(uri, headers: _headers());
    if (res.statusCode != 200) {
      throw Exception(_errorMessage(res, 'Falha a carregar planeamento.'));
    }
    return _decodeMap(res.body);
  }

  Future<Map<String, dynamic>> getMaterialSeparation({
    int horizonDays = 4,
  }) async {
    final uri = Uri.parse('${_requireBaseUrl()}/mobile/material-separation')
        .replace(queryParameters: {'horizon_days': '$horizonDays'});
    final res = await http.get(uri, headers: _headers());
    if (res.statusCode != 200) {
      throw Exception(_errorMessage(res, 'Falha a carregar separacao MP.'));
    }
    return _decodeMap(res.body);
  }

  Future<void> setMaterialSeparationCheck({
    required String checkKey,
    required bool checked,
  }) async {
    final uri = Uri.parse('${_requireBaseUrl()}/mobile/material-separation/check');
    final res = await http.post(
      uri,
      headers: _headers(),
      body: jsonEncode({'check_key': checkKey, 'checked': checked}),
    );
    if (res.statusCode != 200) {
      throw Exception(_errorMessage(res, 'Falha a atualizar visto da separacao.'));
    }
  }

  String planningPdfUrl({String? year, String? weekStart}) {
    final token = _token;
    final effectiveYear = year ?? currentOperationalYear();
    final uri = Uri.parse('${_requireBaseUrl()}/mobile/planning/pdf').replace(
      queryParameters: {
        'year': effectiveYear,
        if (weekStart != null && weekStart.isNotEmpty) 'week_start': weekStart,
        if (token != null && token.isNotEmpty) 'access_token': token,
      },
    );
    return uri.toString();
  }

  String materialSeparationPdfUrl({int horizonDays = 4}) {
    final token = _token;
    final uri = Uri.parse('${_requireBaseUrl()}/mobile/material-separation/pdf').replace(
      queryParameters: {
        'horizon_days': '$horizonDays',
        if (token != null && token.isNotEmpty) 'access_token': token,
      },
    );
    return uri.toString();
  }

  Map<String, String> _headers() {
    final token = _token;
    return {
      'Content-Type': 'application/json',
      if (token != null && token.isNotEmpty) 'Authorization': 'Bearer $token',
    };
  }

  static Map<String, dynamic> _decodeMap(String body) {
    final decoded = jsonDecode(body);
    if (decoded is Map<String, dynamic>) {
      return decoded;
    }
    if (decoded is Map) {
      return decoded.map((key, value) => MapEntry(key.toString(), value));
    }
    return <String, dynamic>{};
  }

  static String _errorMessage(http.Response res, String fallback) {
    try {
      final data = _decodeMap(res.body);
      final detail = data['detail'] ?? data['message'] ?? data['error'];
      if (detail is List) {
        final parts = detail
            .map((item) => item?.toString().trim() ?? '')
            .where((item) => item.isNotEmpty)
            .toList();
        if (parts.isNotEmpty) {
          return parts.join(' | ');
        }
      }
      final text = detail?.toString().trim() ?? '';
      if (text.isNotEmpty) {
        return text;
      }
    } catch (_) {}
    final raw = res.body.trim();
    if (raw.isNotEmpty) {
      return raw;
    }
    return fallback;
  }

  static String _normalizeBaseUrl(String raw) {
    final base = raw.trim().isEmpty ? defaultMobileApiHost.trim() : raw.trim();
    if (base.isEmpty) {
      return '';
    }
    final uri = Uri.tryParse(base);
    final normalized =
        (uri == null || uri.scheme.isEmpty) ? 'http://$base' : base;
    if (normalized.endsWith('/api/v1')) {
      return normalized;
    }
    if (normalized.endsWith('/')) {
      return '${normalized}api/v1';
    }
    return '$normalized/api/v1';
  }
}
