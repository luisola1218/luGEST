<?php

function app_config(string $key, $default = null) {
    $config = $GLOBALS['lugest_web_config'] ?? [];
    if (array_key_exists($key, $config)) {
        return $config[$key];
    }
    return $default;
}

function h($value): string {
    return htmlspecialchars((string) ($value ?? ''), ENT_QUOTES | ENT_SUBSTITUTE, 'UTF-8');
}

function redirect(string $path): void {
    header('Location: ' . $path);
    exit;
}

function query_param(string $name, string $default = ''): string {
    return trim((string) ($_GET[$name] ?? $default));
}

function fmt_date($value): string {
    $raw = trim((string) ($value ?? ''));
    if ($raw === '') {
        return '-';
    }
    try {
        return (new DateTime($raw))->format('d/m/Y');
    } catch (Throwable $e) {
        return $raw;
    }
}

function fmt_datetime($value): string {
    $raw = trim((string) ($value ?? ''));
    if ($raw === '') {
        return '-';
    }
    try {
        return (new DateTime($raw))->format('d/m/Y H:i');
    } catch (Throwable $e) {
        return $raw;
    }
}

function fmt_num($value, int $decimals = 0): string {
    $number = is_numeric($value) ? (float) $value : 0.0;
    return number_format($number, $decimals, ',', '.');
}

function fmt_money($value): string {
    $number = is_numeric($value) ? (float) $value : 0.0;
    return number_format($number, 2, ',', '.') . ' EUR';
}

function fmt_time_minutes($value): string {
    $number = is_numeric($value) ? (int) round((float) $value) : 0;
    if ($number <= 0) {
        return '0 min';
    }
    if ($number < 60) {
        return $number . ' min';
    }
    $hours = intdiv($number, 60);
    $minutes = $number % 60;
    if ($minutes === 0) {
        return $hours . ' h';
    }
    return $hours . ' h ' . str_pad((string) $minutes, 2, '0', STR_PAD_LEFT);
}

function fmt_compact_number($value): string {
    $number = is_numeric($value) ? (float) $value : 0.0;
    if ($number >= 1000000) {
        return number_format($number / 1000000, 1, ',', '.') . ' M';
    }
    if ($number >= 1000) {
        return number_format($number / 1000, 1, ',', '.') . ' K';
    }
    return fmt_num($number);
}

function non_empty_string($value, string $fallback = '-'): string {
    $text = trim((string) ($value ?? ''));
    return $text === '' ? $fallback : $text;
}

function value_or_dash($value): string {
    return non_empty_string($value, '-');
}

function sum_field(array $rows, string $key): float {
    $total = 0.0;
    foreach ($rows as $row) {
        $total += is_numeric($row[$key] ?? null) ? (float) $row[$key] : 0.0;
    }
    return $total;
}

function max_field(array $rows, string $key): ?string {
    $values = [];
    foreach ($rows as $row) {
        $value = trim((string) ($row[$key] ?? ''));
        if ($value !== '') {
            $values[] = $value;
        }
    }
    if ($values === []) {
        return null;
    }
    sort($values);
    return end($values) ?: null;
}

function min_field(array $rows, string $key): ?string {
    $values = [];
    foreach ($rows as $row) {
        $value = trim((string) ($row[$key] ?? ''));
        if ($value !== '') {
            $values[] = $value;
        }
    }
    if ($values === []) {
        return null;
    }
    sort($values);
    return $values[0] ?? null;
}

function max_numeric_field(array $rows, string $key): float {
    $max = 0.0;
    foreach ($rows as $row) {
        $value = is_numeric($row[$key] ?? null) ? (float) $row[$key] : 0.0;
        if ($value > $max) {
            $max = $value;
        }
    }
    return $max;
}

function percentage_of($value, $total): float {
    $base = is_numeric($total) ? (float) $total : 0.0;
    if ($base <= 0) {
        return 0.0;
    }
    $number = is_numeric($value) ? (float) $value : 0.0;
    return max(0.0, min(100.0, ($number / $base) * 100.0));
}

function hex_to_rgba(string $hex, float $alpha = 1.0): string {
    $clean = ltrim(trim($hex), '#');
    if (strlen($clean) === 3) {
        $clean = $clean[0] . $clean[0] . $clean[1] . $clean[1] . $clean[2] . $clean[2];
    }
    if (strlen($clean) !== 6 || !ctype_xdigit($clean)) {
        return 'rgba(58, 64, 72, ' . max(0, min(1, $alpha)) . ')';
    }
    $r = hexdec(substr($clean, 0, 2));
    $g = hexdec(substr($clean, 2, 2));
    $b = hexdec(substr($clean, 4, 2));
    $a = max(0, min(1, $alpha));
    return "rgba({$r}, {$g}, {$b}, {$a})";
}
