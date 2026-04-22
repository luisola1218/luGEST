<?php

return [
    'app_name' => 'luGEST Consulta',
    'brand_eyebrow' => 'Portal industrial',
    'brand_tagline' => 'Consulta web de produção, encomendas e planeamento.',
    'brand_accent' => '#3a4048',
    'brand_support_email' => 'geral@empresa.pt',
    'brand_support_phone' => '+351 000 000 000',
    'timezone' => 'Europe/Lisbon',
    'session_name' => 'lugest_web_ro',
    'allowed_roles' => ['admin', 'owner', 'manager', 'planeamento', 'consulta'],
    'db' => [
        'host' => '127.0.0.1',
        'port' => 3306,
        'name' => 'lugest',
        'user' => 'lugest_web_reader',
        'pass' => 'ALTERAR_PASSWORD',
        'charset' => 'utf8mb4',
    ],
];
