<?php

$envPath = dirname(__DIR__, 2) . '/lugest.env';
$env = [];

if (is_file($envPath)) {
    foreach (file($envPath, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) as $line) {
        $row = trim((string) $line);
        if ($row === '' || str_starts_with($row, '#') || !str_contains($row, '=')) {
            continue;
        }
        [$key, $value] = explode('=', $row, 2);
        $env[trim($key)] = trim($value);
    }
}

return [
    'app_name' => 'luGEST Consulta',
    'brand_eyebrow' => 'Portal industrial',
    'brand_tagline' => 'Ambiente local de teste para consulta web do luGEST.',
    'brand_accent' => '#3a4048',
    'brand_support_email' => 'geral@barcelbal.pt',
    'brand_support_phone' => '+351 253 606 590',
    'timezone' => 'Europe/Lisbon',
    'session_name' => 'lugest_web_ro_local',
    'allowed_roles' => ['admin', 'owner', 'manager', 'planeamento', 'consulta', 'Administrador', 'Admin'],
    'db' => [
        'host' => $env['LUGEST_DB_HOST'] ?? '127.0.0.1',
        'port' => (int) ($env['LUGEST_DB_PORT'] ?? 3306),
        'name' => $env['LUGEST_DB_NAME'] ?? 'lugest',
        'user' => $env['LUGEST_DB_USER'] ?? '',
        'pass' => $env['LUGEST_DB_PASS'] ?? '',
        'charset' => 'utf8mb4',
    ],
];
