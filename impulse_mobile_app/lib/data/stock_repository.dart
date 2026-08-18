import 'dart:convert';
import 'dart:io';

import 'package:crypto/crypto.dart';
import 'package:flutter_secure_storage/flutter_secure_storage.dart';
import 'package:http/http.dart' as http;
import 'package:shared_preferences/shared_preferences.dart';

import '../domain/stock_item.dart';
import '../domain/service_job.dart';

class StockApiSettings {
  const StockApiSettings({required this.baseUrl, required this.token});
  final String baseUrl;
  final String token;

  bool get isConfigured => baseUrl.trim().isNotEmpty && token.trim().isNotEmpty;
}

class JobSyncResult {
  const JobSyncResult({required this.syncedIds, required this.lockedIds});
  final Set<String> syncedIds;
  final Set<String> lockedIds;
}

class StockRepository {
  StockRepository({http.Client? client, FlutterSecureStorage? secureStorage})
      : _client = client ?? http.Client(),
        _secureStorage = secureStorage ?? const FlutterSecureStorage();

  static const _urlKey = 'lugest_mobile_api_url_v1';
  static const _tokenKey = 'lugest_mobile_api_token_v1';
  final http.Client _client;
  final FlutterSecureStorage _secureStorage;

  Future<StockApiSettings> loadSettings() async {
    final preferences = await SharedPreferences.getInstance();
    var token = await _secureStorage.read(key: _tokenKey) ?? '';
    // Migração transparente da primeira versão de testes.
    if (token.isEmpty) {
      token = preferences.getString(_tokenKey) ?? '';
      if (token.isNotEmpty) {
        await _secureStorage.write(key: _tokenKey, value: token);
        await preferences.remove(_tokenKey);
      }
    }
    return StockApiSettings(
      baseUrl: preferences.getString(_urlKey) ?? '',
      token: token,
    );
  }

  Future<void> saveSettings(StockApiSettings settings) async {
    final normalized = _normalizeUrl(settings.baseUrl);
    final uri = Uri.tryParse(normalized);
    if (uri == null || uri.scheme != 'https' || uri.host.isEmpty) {
      throw const StockConnectionException(
          'A ligação remota exige um endereço HTTPS válido.');
    }
    final preferences = await SharedPreferences.getInstance();
    await preferences.setString(_urlKey, normalized);
    await _secureStorage.write(key: _tokenKey, value: settings.token.trim());
  }

  Future<StockSummary> summary() async {
    final json = await _get('/v1/stock/summary');
    return StockSummary.fromJson(json);
  }

  Future<List<StockItem>> items({
    required String kind,
    String query = '',
    bool inStockOnly = true,
  }) async {
    final params = <String, String>{
      if (query.trim().isNotEmpty) 'q': query.trim(),
      'in_stock': inStockOnly ? '1' : '0',
    };
    final path =
        kind == 'material' ? '/v1/stock/materials' : '/v1/stock/products';
    final json = await _get(path, params);
    final raw = json['items'] as List<dynamic>? ?? const [];
    return raw
        .map((item) => StockItem.fromJson(item as Map<String, dynamic>))
        .toList();
  }

  Future<void> testConnection(StockApiSettings settings) async {
    await saveSettings(settings);
    await _get('/health');
  }

  Future<JobSyncResult> syncJobs(List<ServiceJob> jobs) async {
    final payload = {
      'jobs': jobs.map((job) => job.toJson()).toList(),
    };
    final json = await _post('/v1/service-jobs/sync', payload);
    final rows = json['jobs'] as List<dynamic>? ?? const [];
    final synced = <String>{};
    final locked = <String>{};
    for (final raw in rows) {
      final row = raw as Map<String, dynamic>;
      final id = row['id'] as String? ?? '';
      if (id.isEmpty) continue;
      synced.add(id);
      if (row['locked'] == true) locked.add(id);
    }
    return JobSyncResult(syncedIds: synced, lockedIds: locked);
  }

  Future<int> uploadJobAttachments(ServiceJob job) async {
    final attachments = <({String kind, String path})>[
      for (final path in job.photoPaths) (kind: 'photo', path: path),
      if (job.signaturePath case final path?) (kind: 'signature', path: path),
    ];
    var uploaded = 0;
    for (final attachment in attachments) {
      try {
        final file = File(attachment.path);
        if (!await file.exists()) {
          throw const StockConnectionException(
              'Uma fotografia ou assinatura já não existe neste telemóvel.');
        }
        final bytes = await file.readAsBytes();
        if (bytes.isEmpty || bytes.length > 10 * 1024 * 1024) {
          throw const StockConnectionException(
              'Uma fotografia ou assinatura está vazia ou excede 10 MB.');
        }
        final name = attachment.path.split(RegExp(r'[/\\]')).last;
        await _post(
          '/v1/service-jobs/attachments',
          {
            'jobId': job.id,
            'kind': attachment.kind,
            'fileName': name,
            'mimeType': _imageMime(name, bytes),
            'sha256': sha256.convert(bytes).toString(),
            'dataBase64': base64Encode(bytes),
          },
          timeout: const Duration(seconds: 60),
        );
        uploaded++;
      } on StockConnectionException {
        rethrow;
      } catch (_) {
        throw const StockConnectionException(
            'Não foi possível preparar uma fotografia ou assinatura para envio.');
      }
    }
    return uploaded;
  }

  Future<Map<String, dynamic>> _get(String path,
      [Map<String, String>? params]) async {
    final settings = await loadSettings();
    if (!settings.isConfigured) {
      throw const StockConnectionException(
          'Configure o endereço e a chave da ligação ao LuGEST.');
    }
    final uri = Uri.parse('${_normalizeUrl(settings.baseUrl)}$path')
        .replace(queryParameters: params);
    try {
      final response = await _client.get(uri, headers: {
        'Accept': 'application/json',
        'Authorization': 'Bearer ${settings.token}',
      }).timeout(const Duration(seconds: 10));
      final body = response.body.isEmpty
          ? <String, dynamic>{}
          : jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode < 200 || response.statusCode >= 300) {
        throw StockConnectionException(
          body['error'] as String? ??
              'Ligação recusada (${response.statusCode}).',
        );
      }
      return body;
    } on StockConnectionException {
      rethrow;
    } catch (_) {
      throw const StockConnectionException(
          'Não foi possível chegar ao LuGEST. Confirme a Internet e o endereço HTTPS.');
    }
  }

  Future<Map<String, dynamic>> _post(String path, Map<String, dynamic> payload,
      {Duration timeout = const Duration(seconds: 20)}) async {
    final settings = await loadSettings();
    if (!settings.isConfigured) {
      throw const StockConnectionException(
          'Configure primeiro a ligação ao LuGEST em Mais > Stock LuGEST.');
    }
    final uri = Uri.parse('${_normalizeUrl(settings.baseUrl)}$path');
    try {
      final response = await _client
          .post(uri,
              headers: {
                'Accept': 'application/json',
                'Content-Type': 'application/json; charset=utf-8',
                'Authorization': 'Bearer ${settings.token}',
              },
              body: jsonEncode(payload))
          .timeout(timeout);
      final body = response.body.isEmpty
          ? <String, dynamic>{}
          : jsonDecode(response.body) as Map<String, dynamic>;
      if (response.statusCode < 200 || response.statusCode >= 300) {
        throw StockConnectionException(
          body['error'] as String? ??
              'Sincronização recusada (${response.statusCode}).',
        );
      }
      return body;
    } on StockConnectionException {
      rethrow;
    } catch (_) {
      throw const StockConnectionException(
          'Não foi possível enviar os serviços para o LuGEST.');
    }
  }

  static String _normalizeUrl(String value) {
    var url = value.trim();
    while (url.endsWith('/')) {
      url = url.substring(0, url.length - 1);
    }
    if (url.isNotEmpty &&
        !url.startsWith('http://') &&
        !url.startsWith('https://')) {
      url = 'https://$url';
    }
    return url;
  }

  static String _imageMime(String name, List<int> bytes) {
    final lower = name.toLowerCase();
    if (bytes.length >= 8 &&
        bytes[0] == 0x89 &&
        bytes[1] == 0x50 &&
        bytes[2] == 0x4e &&
        bytes[3] == 0x47) {
      return 'image/png';
    }
    if (bytes.length >= 12 &&
        bytes[0] == 0x52 &&
        bytes[1] == 0x49 &&
        bytes[2] == 0x46 &&
        bytes[3] == 0x46 &&
        bytes[8] == 0x57 &&
        bytes[9] == 0x45 &&
        bytes[10] == 0x42 &&
        bytes[11] == 0x50) {
      return 'image/webp';
    }
    if (lower.endsWith('.heic') || lower.endsWith('.heif')) {
      return 'image/heic';
    }
    return 'image/jpeg';
  }
}

class StockConnectionException implements Exception {
  const StockConnectionException(this.message);
  final String message;

  @override
  String toString() => message;
}
