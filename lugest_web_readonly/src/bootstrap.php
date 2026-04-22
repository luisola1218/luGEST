<?php

$configPath = dirname(__DIR__) . '/config/config.php';
if (!is_file($configPath)) {
    http_response_code(500);
    echo 'Falta o ficheiro config/config.php. Copia config.example.php e ajusta as credenciais.';
    exit;
}

$GLOBALS['lugest_web_config'] = require $configPath;

require_once __DIR__ . '/helpers.php';

date_default_timezone_set((string) app_config('timezone', 'Europe/Lisbon'));
session_name((string) app_config('session_name', 'lugest_web_ro'));
if (session_status() !== PHP_SESSION_ACTIVE) {
    session_start();
}

require_once __DIR__ . '/db.php';
require_once __DIR__ . '/auth.php';
require_once __DIR__ . '/repositories.php';
require_once __DIR__ . '/layout.php';
