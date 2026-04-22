<?php

require_once dirname(__DIR__) . '/src/bootstrap.php';
require_login();

$filters = [
    'q' => query_param('q'),
    'year' => query_param('year', (string) date('Y')),
    'estado' => query_param('estado'),
];
$rows = repo_orders($filters, 250);

$countOrders = count($rows);
$totalBlocks = sum_field($rows, 'blocos_total');
$totalTime = sum_field($rows, 'tempo_estimado');
$latestDelivery = max_field($rows, 'data_entrega');

$actions = '
    <form method="get" class="toolbar">
        <label class="field">
            <span>Pesquisa</span>
            <input type="text" name="q" placeholder="Numero, cliente ou nota..." value="' . h($filters['q']) . '">
        </label>
        <label class="field compact">
            <span>Ano</span>
            <input type="number" name="year" value="' . h($filters['year']) . '" min="2020" max="2100">
        </label>
        <label class="field compact">
            <span>Estado</span>
            <input type="text" name="estado" placeholder="Estado" value="' . h($filters['estado']) . '">
        </label>
        <button type="submit" class="btn btn-primary">
            <span class="btn-icon">' . app_icon('filter') . '</span>
            <span>Filtrar</span>
        </button>
    </form>
';

$aside = '
    <div class="hero-side-card">
        <div>
            <p class="eyebrow">Leitura da carteira</p>
            <p>Vista global das encomendas registadas, com foco em entrega, estado e volume planeado.</p>
        </div>
        <div class="hero-side-grid">
            <div><span>Total visivel</span><strong>' . h((string) $countOrders) . '</strong></div>
            <div><span>Ultima entrega</span><strong>' . h(fmt_date($latestDelivery)) . '</strong></div>
            <div><span>Tempo acumulado</span><strong>' . h(fmt_time_minutes($totalTime)) . '</strong></div>
        </div>
    </div>
';

ob_start();
echo render_page_intro(
    'Consulta comercial',
    'Encomendas',
    'Pesquisa e acompanha a carteira de encomendas com uma apresentacao mais limpa, contextual e pronta para consulta rapida.',
    [
        ['label' => 'Ano', 'value' => $filters['year']],
        ['label' => 'Resultados', 'value' => (string) $countOrders],
        ['label' => 'Blocos', 'value' => fmt_num($totalBlocks)],
    ],
    $actions,
    $aside
);
?>
<section class="metric-grid compact">
    <?php echo render_metric_card('Encomendas visiveis', fmt_compact_number($countOrders), 'Registos dentro dos filtros ativos.', 'orders', 'default'); ?>
    <?php echo render_metric_card('Tempo estimado', fmt_time_minutes($totalTime), 'Carga agregada das encomendas listadas.', 'clock', 'warning'); ?>
    <?php echo render_metric_card('Blocos ligados', fmt_num($totalBlocks), 'Quantidade total de blocos associados.', 'planning', 'success'); ?>
</section>

<section class="surface">
    <div class="surface-head">
        <div>
            <p class="surface-kicker">Lista operacional</p>
            <h2>Carteira de encomendas</h2>
            <p class="muted">Abertura rapida do detalhe de cada encomenda e respetivo estado.</p>
        </div>
    </div>
    <div class="surface-body">
        <div class="data-table-wrap">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Numero</th>
                        <th>Cliente</th>
                        <th>Estado</th>
                        <th>Entrega</th>
                        <th>Tempo</th>
                        <th>Pecas</th>
                        <th>Blocos</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($rows as $row): ?>
                        <tr>
                            <td>
                                <a class="table-link" href="order.php?numero=<?php echo urlencode((string) $row['numero']); ?>">
                                    <?php echo h($row['numero']); ?>
                                </a>
                                <span class="table-secondary">Detalhe da encomenda</span>
                            </td>
                            <td><?php echo h($row['cliente_nome']); ?></td>
                            <td><?php echo render_status_pill((string) $row['estado']); ?></td>
                            <td><?php echo h(fmt_date($row['data_entrega'])); ?></td>
                            <td><?php echo h(fmt_time_minutes($row['tempo_estimado'])); ?></td>
                            <td><?php echo h(fmt_num($row['pecas_total'])); ?></td>
                            <td><?php echo h(fmt_num($row['blocos_total'])); ?></td>
                        </tr>
                    <?php endforeach; ?>
                    <?php if (!$rows): ?>
                        <tr><td colspan="7" class="empty-row">Sem resultados para os filtros atuais.</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</section>
<?php
$content = ob_get_clean();
render_app_page('Encomendas', $content, 'orders');
