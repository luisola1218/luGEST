<?php

require_once dirname(__DIR__) . '/src/bootstrap.php';
require_login();

$year = query_param('year', (string) date('Y'));
$stats = repo_dashboard_stats($year);
$recentOrders = repo_recent_orders($year, 8);
$statusBreakdown = repo_dashboard_status_breakdown($year, 5);
$planningDays = repo_dashboard_planning_by_day($year, 7);
$topClients = repo_dashboard_top_clients($year, 5);

$statIcons = [
    'Encomendas' => ['icon' => 'orders', 'tone' => 'default'],
    'Planeamento' => ['icon' => 'planning', 'tone' => 'warning'],
    'Produtos' => ['icon' => 'products', 'tone' => 'success'],
    'Materia-prima' => ['icon' => 'materials', 'tone' => 'accent'],
];

$chips = [
    ['label' => 'Ano', 'value' => $year],
    ['label' => 'Modo', 'value' => 'Consulta'],
    ['label' => 'Base', 'value' => 'luGEST'],
];

$actions = '
    <form method="get" class="toolbar">
        <label class="field compact">
            <span>Ano</span>
            <input type="number" name="year" value="' . h($year) . '" min="2020" max="2100">
        </label>
        <button type="submit" class="btn btn-primary">
            <span class="btn-icon">' . app_icon('trend') . '</span>
            <span>Atualizar painel</span>
        </button>
    </form>
';

$aside = '
    <div class="hero-side-card">
        <div>
            <p class="eyebrow">Resumo executivo</p>
            <p>Vista rapida do ambiente industrial para consulta comercial, producao e acompanhamento remoto.</p>
        </div>
        <div class="hero-side-grid">
            <div><span>Utilizador</span><strong>' . h((string) (current_user()['username'] ?? '-')) . '</strong></div>
            <div><span>Ultima leitura</span><strong>' . h(date('d/m/Y H:i')) . '</strong></div>
            <div><span>Registos listados</span><strong>' . h((string) count($recentOrders)) . '</strong></div>
        </div>
    </div>
';

$statusTotal = max(1, (int) sum_field($statusBreakdown, 'total'));
$statusSlices = [];
$statusPalette = ['var(--brand-highlight)', 'var(--brand-accent)', 'var(--success)', 'var(--warning)', '#9aa0a6'];
$statusOffset = 0.0;
foreach ($statusBreakdown as $index => $row) {
    $percent = percentage_of($row['total'] ?? 0, $statusTotal);
    $start = $statusOffset;
    $end = $statusOffset + $percent;
    $color = $statusPalette[$index] ?? '#b2b7bd';
    $statusSlices[] = $color . ' ' . number_format($start, 2, '.', '') . '% ' . number_format($end, 2, '.', '') . '%';
    $statusOffset = $end;
}
$statusChartStyle = $statusSlices !== [] ? implode(', ', $statusSlices) : 'var(--brand-accent-soft) 0% 100%';

$maxPlanningLoad = max_numeric_field($planningDays, 'duracao_total');
$maxClientTotal = max_numeric_field($topClients, 'total');

ob_start();
echo render_page_intro(
    'Consulta operacional',
    'Dashboard industrial',
    'Uma leitura clara do estado global do luGEST, agora com graficos executivos para perceber distribuicao, carga e prioridade sem sair do portal.',
    $chips,
    $actions,
    $aside
);
?>
<section class="metric-grid">
    <?php foreach ($stats as $item): ?>
        <?php
        $label = (string) ($item['label'] ?? '');
        $iconMeta = $statIcons[$label] ?? ['icon' => 'stack', 'tone' => 'default'];
        echo render_metric_card(
            $label,
            (string) ($item['value'] ?? '0'),
            (string) ($item['note'] ?? ''),
            (string) $iconMeta['icon'],
            (string) $iconMeta['tone']
        );
        ?>
    <?php endforeach; ?>
</section>

<section class="chart-grid">
    <article class="surface chart-surface">
        <div class="surface-head">
            <div>
                <p class="surface-kicker">Distribuicao de estados</p>
                <h2>Encomendas por estado</h2>
                <p class="muted">Vista imediata do peso de cada estado no ano selecionado.</p>
            </div>
        </div>
        <div class="surface-body donut-layout">
            <div class="donut-wrap">
                <div class="donut-chart" style="--chart-slices: <?php echo h($statusChartStyle); ?>;">
                    <div class="donut-center">
                        <strong><?php echo h(fmt_num($statusTotal)); ?></strong>
                        <span>encomendas</span>
                    </div>
                </div>
            </div>
            <div class="legend-list">
                <?php foreach ($statusBreakdown as $index => $row): ?>
                    <?php
                    $color = $statusPalette[$index] ?? '#b2b7bd';
                    $percent = percentage_of($row['total'] ?? 0, $statusTotal);
                    ?>
                    <div class="legend-item">
                        <span class="legend-swatch" style="--legend-color: <?php echo h($color); ?>;"></span>
                        <div class="legend-copy">
                            <strong><?php echo h((string) $row['estado']); ?></strong>
                            <span><?php echo h(fmt_num($row['total'])); ?> registos</span>
                        </div>
                        <em><?php echo h(fmt_num($percent, 1)); ?>%</em>
                    </div>
                <?php endforeach; ?>
                <?php if (!$statusBreakdown): ?>
                    <div class="empty-row">Sem estados para mostrar.</div>
                <?php endif; ?>
            </div>
        </div>
    </article>

    <article class="surface chart-surface">
        <div class="surface-head">
            <div>
                <p class="surface-kicker">Carga semanal</p>
                <h2>Planeamento por dia</h2>
                <p class="muted">Duracao total dos blocos por dia planeado.</p>
            </div>
        </div>
        <div class="surface-body chart-bars">
            <?php foreach ($planningDays as $row): ?>
                <?php $height = max(14, percentage_of($row['duracao_total'] ?? 0, $maxPlanningLoad)); ?>
                <div class="chart-bar-card">
                    <div class="chart-bar-meta">
                        <strong><?php echo h(fmt_time_minutes($row['duracao_total'])); ?></strong>
                        <span><?php echo h(fmt_num($row['blocos'])); ?> blocos</span>
                    </div>
                    <div class="chart-bar-track">
                        <div class="chart-bar-fill" style="height: <?php echo h(fmt_num($height, 2)); ?>%;"></div>
                    </div>
                    <div class="chart-bar-label"><?php echo h(fmt_date($row['data_planeada'])); ?></div>
                </div>
            <?php endforeach; ?>
            <?php if (!$planningDays): ?>
                <div class="empty-row">Sem carga planeada para mostrar.</div>
            <?php endif; ?>
        </div>
    </article>
</section>

<section class="split-grid">
    <div class="surface">
        <div class="surface-head">
            <div>
                <p class="surface-kicker">Fluxo comercial</p>
                <h2>Encomendas recentes</h2>
                <p class="muted">Ultimos registos visiveis no portal para o ano selecionado.</p>
            </div>
            <a class="btn btn-secondary btn-linkish" href="orders.php?year=<?php echo urlencode($year); ?>">
                <span class="btn-icon"><?php echo app_icon('arrow-right'); ?></span>
                <span>Ver lista completa</span>
            </a>
        </div>
        <div class="surface-body">
            <div class="data-table-wrap">
                <table class="data-table">
                    <thead>
                        <tr>
                            <th>Encomenda</th>
                            <th>Cliente</th>
                            <th>Estado</th>
                            <th>Entrega</th>
                            <th>Pecas</th>
                            <th>Blocos</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php foreach ($recentOrders as $row): ?>
                            <tr>
                                <td>
                                    <a class="table-link" href="order.php?numero=<?php echo urlencode((string) $row['numero']); ?>">
                                        <?php echo h($row['numero']); ?>
                                    </a>
                                </td>
                                <td><?php echo h($row['cliente_nome']); ?></td>
                                <td><?php echo render_status_pill((string) $row['estado']); ?></td>
                                <td><?php echo h(fmt_date($row['data_entrega'])); ?></td>
                                <td><?php echo h(fmt_num($row['pecas_total'])); ?></td>
                                <td><?php echo h(fmt_num($row['blocos_total'])); ?></td>
                            </tr>
                        <?php endforeach; ?>
                        <?php if (!$recentOrders): ?>
                            <tr><td colspan="6" class="empty-row">Sem registos para mostrar.</td></tr>
                        <?php endif; ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <div class="surface-stack">
        <section class="surface">
            <div class="surface-head">
                <div>
                    <p class="surface-kicker">Clientes em foco</p>
                    <h2>Top clientes do ano</h2>
                </div>
            </div>
            <div class="surface-body ranking-list">
                <?php foreach ($topClients as $row): ?>
                    <?php $barWidth = max(12, percentage_of($row['total'] ?? 0, $maxClientTotal)); ?>
                    <div class="ranking-item">
                        <div class="ranking-copy">
                            <strong><?php echo h((string) $row['cliente_nome']); ?></strong>
                            <span><?php echo h(fmt_time_minutes($row['tempo_total'])); ?> de carga estimada</span>
                        </div>
                        <div class="ranking-bar">
                            <div class="ranking-bar-fill" style="width: <?php echo h(fmt_num($barWidth, 2)); ?>%;"></div>
                        </div>
                        <em><?php echo h(fmt_num($row['total'])); ?></em>
                    </div>
                <?php endforeach; ?>
                <?php if (!$topClients): ?>
                    <div class="empty-row">Sem clientes para analisar.</div>
                <?php endif; ?>
            </div>
        </section>

        <section class="surface">
            <div class="surface-head">
                <div>
                    <p class="surface-kicker">Acesso rapido</p>
                    <h2>Areas principais</h2>
                </div>
            </div>
            <div class="surface-body mini-grid">
                <a class="sub-card" href="planning.php">
                    <h3>Planeamento</h3>
                    <p>Consulta blocos, datas e duracoes sem abrir a app desktop.</p>
                </a>
                <a class="sub-card" href="quotes.php">
                    <h3>Orcamentos</h3>
                    <p>Segue a atividade comercial e o estado dos documentos emitidos.</p>
                </a>
                <a class="sub-card" href="materials.php">
                    <h3>Materia-prima</h3>
                    <p>Verifica stock, localizacao e movimentos mais recentes.</p>
                </a>
            </div>
        </section>
    </div>
</section>
<?php
$content = ob_get_clean();
render_app_page('Dashboard', $content, 'dashboard');
