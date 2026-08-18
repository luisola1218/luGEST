const _months = [
  'jan',
  'fev',
  'mar',
  'abr',
  'mai',
  'jun',
  'jul',
  'ago',
  'set',
  'out',
  'nov',
  'dez'
];
const _weekdays = ['seg', 'ter', 'qua', 'qui', 'sex', 'sáb', 'dom'];

String money(double value) {
  final parts = value.toStringAsFixed(2).split('.');
  final digits = parts.first;
  final grouped = StringBuffer();
  for (var index = 0; index < digits.length; index++) {
    if (index > 0 && (digits.length - index) % 3 == 0) grouped.write('.');
    grouped.write(digits[index]);
  }
  return '${grouped.toString()},${parts.last} €';
}

String shortDate(DateTime value) =>
    '${value.day.toString().padLeft(2, '0')} ${_months[value.month - 1]}';
String fullDate(DateTime value) =>
    '${_weekdays[value.weekday - 1]}, ${value.day} ${_months[value.month - 1]}';
String clock(DateTime value) =>
    '${value.hour.toString().padLeft(2, '0')}:${value.minute.toString().padLeft(2, '0')}';
