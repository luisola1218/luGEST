const String defaultMobileApiHost = String.fromEnvironment(
  'LUGEST_DEFAULT_API_HOST',
  defaultValue: '',
);
const String mobileApiHostExample = 'http://servidor:8050';

String configuredMobileApiBaseUrl() {
  final host = defaultMobileApiHost.trim();
  if (host.isEmpty) {
    return '';
  }
  if (host.endsWith('/api/v1')) {
    return host;
  }
  if (host.endsWith('/')) {
    return '${host}api/v1';
  }
  return '$host/api/v1';
}

String currentOperationalYear([DateTime? now]) {
  return (now ?? DateTime.now()).year.toString();
}

List<String> recentOperationalYears({int count = 3, DateTime? now}) {
  final currentYear = int.parse(currentOperationalYear(now));
  return List<String>.generate(count, (index) => '${currentYear - index}');
}
