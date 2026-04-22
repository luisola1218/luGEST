<?php

function app_nav_items(): array {
    return [
        'dashboard' => ['label' => 'Dashboard', 'href' => 'dashboard.php', 'icon' => 'dashboard'],
        'orders' => ['label' => 'Encomendas', 'href' => 'orders.php', 'icon' => 'orders'],
        'planning' => ['label' => 'Planeamento', 'href' => 'planning.php', 'icon' => 'planning'],
        'quotes' => ['label' => 'Orcamentos', 'href' => 'quotes.php', 'icon' => 'quotes'],
        'clients' => ['label' => 'Clientes', 'href' => 'clients.php', 'icon' => 'clients'],
        'products' => ['label' => 'Produtos', 'href' => 'products.php', 'icon' => 'products'],
        'materials' => ['label' => 'Materia-prima', 'href' => 'materials.php', 'icon' => 'materials'],
    ];
}

function app_icon(string $name): string {
    $icons = [
        'dashboard' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 13h7V4H4zm9 7h7v-9h-7zM4 20h7v-5H4zm9-9h7V4h-7z"/></svg>',
        'orders' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M6 4h9l5 5v11a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2zm8 1.5V10h4.5"/><path d="M8 13h8M8 17h8M8 9h3"/></svg>',
        'planning' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M7 3v3M17 3v3M4 8h16"/><rect x="4" y="5" width="16" height="16" rx="2"/><path d="M8 12h3v3H8zm5 0h3v3h-3z"/></svg>',
        'quotes' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 4h14a2 2 0 0 1 2 2v9a2 2 0 0 1-2 2H9l-4 3V6a2 2 0 0 1 2-2z"/><path d="M8 9h8M8 13h5"/></svg>',
        'clients' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M16 21v-2a4 4 0 0 0-4-4H8a4 4 0 0 0-4 4v2"/><circle cx="10" cy="7" r="4"/><path d="M20 21v-2a4 4 0 0 0-3-3.87"/><path d="M17 3.13a4 4 0 0 1 0 7.75"/></svg>',
        'products' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3 4 7l8 4 8-4-8-4z"/><path d="M4 7v10l8 4 8-4V7"/><path d="M12 11v10"/></svg>',
        'materials' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 7a2 2 0 0 1 2-2h12v12a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V7z"/><path d="M8 5v14M12 5v14M16 5v14"/></svg>',
        'arrow-left' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M19 12H5"/><path d="m11 6-6 6 6 6"/></svg>',
        'arrow-right' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M5 12h14"/><path d="m13 6 6 6-6 6"/></svg>',
        'filter' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 6h16M7 12h10M10 18h4"/></svg>',
        'clock' => '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="12" cy="12" r="9"/><path d="M12 7v5l3 2"/></svg>',
        'stack' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="m12 3 9 5-9 5-9-5 9-5z"/><path d="m3 12 9 5 9-5"/><path d="m3 16 9 5 9-5"/></svg>',
        'trend' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M4 16 10 10l4 4 6-7"/><path d="M14 7h6v6"/></svg>',
        'shield' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M12 3 5 6v5c0 5 3.5 8.74 7 10 3.5-1.26 7-5 7-10V6l-7-3z"/></svg>',
        'search' => '<svg viewBox="0 0 24 24" aria-hidden="true"><circle cx="11" cy="11" r="7"/><path d="m20 20-3.5-3.5"/></svg>',
        'user' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M20 21a8 8 0 0 0-16 0"/><circle cx="12" cy="8" r="4"/></svg>',
        'logout' => '<svg viewBox="0 0 24 24" aria-hidden="true"><path d="M9 21H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h4"/><path d="M16 17l5-5-5-5"/><path d="M21 12H9"/></svg>',
    ];

    return $icons[$name] ?? $icons['stack'];
}

function status_tone(string $value): string {
    $text = strtolower(trim($value));
    if ($text === '' || $text === '-') {
        return 'neutral';
    }
    if (str_contains($text, 'concl') || str_contains($text, 'fech') || str_contains($text, 'entreg')) {
        return 'success';
    }
    if (str_contains($text, 'curso') || str_contains($text, 'produc') || str_contains($text, 'planea') || str_contains($text, 'emit')) {
        return 'accent';
    }
    if (str_contains($text, 'pend') || str_contains($text, 'espera') || str_contains($text, 'analise')) {
        return 'warning';
    }
    if (str_contains($text, 'cancel') || str_contains($text, 'erro') || str_contains($text, 'atras')) {
        return 'danger';
    }
    return 'neutral';
}

function render_status_pill(string $value): string {
    $label = non_empty_string($value, '-');
    return '<span class="pill pill-' . h(status_tone($label)) . '">' . h($label) . '</span>';
}

function render_metric_card(string $label, string $value, string $note = '', string $icon = 'stack', string $tone = 'default'): string {
    $html = '<article class="metric-card tone-' . h($tone) . '">';
    $html .= '<div class="metric-icon">' . app_icon($icon) . '</div>';
    $html .= '<div class="metric-copy">';
    $html .= '<p class="metric-label">' . h($label) . '</p>';
    $html .= '<p class="metric-value">' . h($value) . '</p>';
    if ($note !== '') {
        $html .= '<p class="metric-note">' . h($note) . '</p>';
    }
    $html .= '</div></article>';
    return $html;
}

function render_overview_chip(string $label, string $value): string {
    return '<div class="overview-chip"><span>' . h($label) . '</span><strong>' . h($value) . '</strong></div>';
}

function render_page_intro(string $eyebrow, string $title, string $description, array $chips = [], string $actions = '', string $aside = ''): string {
    $html = '<section class="page-hero">';
    $html .= '<div class="page-hero-main">';
    $html .= '<p class="eyebrow">' . h($eyebrow) . '</p>';
    $html .= '<h1>' . h($title) . '</h1>';
    $html .= '<p class="hero-text">' . h($description) . '</p>';
    if ($chips !== []) {
        $html .= '<div class="hero-chip-row">';
        foreach ($chips as $chip) {
            $html .= render_overview_chip((string) ($chip['label'] ?? ''), (string) ($chip['value'] ?? ''));
        }
        $html .= '</div>';
    }
    if ($actions !== '') {
        $html .= '<div class="hero-actions">' . $actions . '</div>';
    }
    $html .= '</div>';
    if ($aside !== '') {
        $html .= '<aside class="page-hero-side">' . $aside . '</aside>';
    }
    $html .= '</section>';
    return $html;
}

function render_guest_page(string $title, string $content): void {
    $appName = app_config('app_name', 'luGEST Consulta');
    $accent = app_config('brand_accent', '#3a4048');
    echo '<!doctype html><html lang="pt"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">';
    echo '<title>' . h($title) . ' | ' . h($appName) . '</title>';
    echo '<link rel="stylesheet" href="assets/app.css">';
    echo '<style>:root{--brand-accent:' . h($accent) . ';}</style></head><body class="guest-body">';
    echo '<div class="ambient ambient-a"></div><div class="ambient ambient-b"></div>';
    echo $content;
    echo '</body></html>';
}

function render_app_page(string $title, string $content, string $active): void {
    $appName = app_config('app_name', 'luGEST Consulta');
    $accent = app_config('brand_accent', '#3a4048');
    $user = current_user();
    $eyebrow = app_config('brand_eyebrow', 'Portal industrial');
    $tagline = app_config('brand_tagline', 'Consulta web de leitura');
    $supportEmail = app_config('brand_support_email', '');
    $supportPhone = app_config('brand_support_phone', '');
    $nav = app_nav_items();

    echo '<!doctype html><html lang="pt"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1">';
    echo '<title>' . h($title) . ' | ' . h($appName) . '</title>';
    echo '<link rel="stylesheet" href="assets/app.css">';
    echo '<style>:root{--brand-accent:' . h($accent) . ';}</style></head><body>';
    echo '<div class="ambient ambient-a"></div><div class="ambient ambient-b"></div>';
    echo '<div class="app-shell">';
    echo '<aside class="sidebar">';
    echo '<div class="sidebar-card brand-card">';
    echo '<div class="brand-logo-shell"><img class="brand-logo" src="assets/desktop-brand-logo.jpg" alt="Barcelbal"></div>';
    echo '<div class="brand-copy"><small>' . h($eyebrow) . '</small><h1>' . h($appName) . '</h1><p>' . h($tagline) . '</p></div>';
    echo '</div>';
    echo '<nav class="nav">';
    foreach ($nav as $key => $item) {
        $class = $key === $active ? 'active' : '';
        echo '<a class="' . h($class) . '" href="' . h($item['href']) . '"><span class="nav-icon">' . app_icon((string) $item['icon']) . '</span><span class="nav-copy">' . h($item['label']) . '</span></a>';
    }
    echo '</nav>';
    echo '<div class="sidebar-card user-card">';
    echo '<div class="user-card-head"><span class="user-icon">' . app_icon('user') . '</span><div><strong>' . h((string) ($user['username'] ?? '-')) . '</strong><span>' . h((string) ($user['role'] ?? '-')) . '</span></div></div>';
    echo '<div class="user-meta"><div><span>Ambiente</span><strong>Leitura local</strong></div><div><span>Atualizado</span><strong>' . h(date('d/m/Y H:i')) . '</strong></div></div>';
    if ($supportEmail !== '' || $supportPhone !== '') {
        echo '<div class="support-block">';
        echo '<div class="support-title">Suporte</div>';
        if ($supportPhone !== '') {
            echo '<div>' . h($supportPhone) . '</div>';
        }
        if ($supportEmail !== '') {
            echo '<div>' . h($supportEmail) . '</div>';
        }
        echo '</div>';
    }
    echo '<a class="logout-link" href="logout.php"><span class="nav-icon">' . app_icon('logout') . '</span><span>Terminar sessao</span></a>';
    echo '</div>';
    echo '</aside>';
    echo '<main class="main-shell">';
    echo '<header class="topbar">';
    echo '<div><p class="topbar-label">Software premium</p><h2>' . h($title) . '</h2></div>';
    echo '<div class="topbar-meta"><span class="topbar-pill">Read-only</span><span class="topbar-pill">' . h(date('d/m/Y')) . '</span></div>';
    echo '</header>';
    echo '<div class="page-frame">' . $content . '</div>';
    echo '</main>';
    echo '</div></body></html>';
}
