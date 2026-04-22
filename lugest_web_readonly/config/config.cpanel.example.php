<?php

return [
    'app_name' => 'luGEST Consulta',
    'brand_eyebrow' => 'Portal industrial',
    'brand_tagline' => 'Consulta web de leitura para produção, planeamento e comercial.',
    'brand_accent' => '#3a4048',
    'brand_support_email' => 'geral@empresa.pt',
    'brand_support_phone' => '+351 000 000 000',
    'timezone' => 'Europe/Lisbon',
    'session_name' => 'lugest_web_ro',

    // Perfis autorizados a entrar no portal.
    'allowed_roles' => ['admin', 'owner', 'manager', 'planeamento', 'consulta'],

    // Ligação MySQL do alojamento/cPanel.
    'db' => [
        'host' => 'localhost',
        'port' => 3306,
        'name' => 'cpaneluser_lugest',
        'user' => 'cpaneluser_lugestro',
        'pass' => 'ALTERAR_PASSWORD_FORTE',
        'charset' => 'utf8mb4',
    ],
];
