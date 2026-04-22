<?php

require_once dirname(__DIR__) . '/src/bootstrap.php';
require_login();

$filters = [
    'q' => query_param('q'),
    'year' => query_param('year', (string) date('Y')),
];
$rows = repo_quotes($filters, 250);

$totalSubtotal = sum_field($rows, 'subtotal');
$totalAmount = sum_field($rows, 'total');

$actions = '
    <form method="get" class="toolbar">
        <label class="field">
            <span>Pesquisa</span>
            <input type="text" name="q" placeholder="Numero, cliente ou estado..." value="' . h($filters['q']) . '">
        </label>
        <label class="field compact">
            <span>Ano</span>
            <input type="number" name="year" value="' . h($filters['year']) . '" min="2020" max="2100">
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
            <p class="eyebrow">Contexto comercial</p>
            <p>Consulta de orcamentos emitidos com visao imediata de cliente, estado e montante associado.</p>
        </div>
        <div class="hero-side-grid">
            <div><span>Registos</span><strong>' . h((string) count($rows)) . '</strong></div>
            <div><span>Subtotal</span><strong>' . h(fmt_money($totalSubtotal)) . '</strong></div>
            <div><span>Total</span><strong>' . h(fmt_money($totalAmount)) . '</strong></div>
        </div>
    </div>
';

ob_start();
echo render_page_intro(
    'Painel comercial',
    'Orcamentos',
    'Acompanha orcamentos emitidos com uma apresentacao mais clara, elegante e orientada a decisao comercial.',
    [
        ['label' => 'Ano', 'value' => $filters['year']],
        ['label' => 'Registos', 'value' => (string) count($rows)],
        ['label' => 'Total', 'value' => fmt_money($totalAmount)],
    ],
    $actions,
    $aside
);
?>
<section class="metric-grid compact">
    <?php echo render_metric_card('Orcamentos', fmt_compact_number(count($rows)), 'Numero de documentos visiveis.', 'quotes', 'default'); ?>
    <?php echo render_metric_card('Subtotal agregado', fmt_money($totalSubtotal), 'Valor antes de impostos.', 'trend', 'warning'); ?>
    <?php echo render_metric_card('Total agregado', fmt_money($totalAmount), 'Soma global dos montantes listados.', 'shield', 'success'); ?>
</section>

<section class="surface">
    <div class="surface-head">
        <div>
            <p class="surface-kicker">Documentos comerciais</p>
            <h2>Orcamentos emitidos</h2>
            <p class="muted">Consulta de cliente, estado, montante e ligacao a encomenda.</p>
        </div>
    </div>
    <div class="surface-body">
        <div class="data-table-wrap">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Numero</th>
                        <th>Data</th>
                        <th>Cliente</th>
                        <th>Estado</th>
                        <th>Subtotal</th>
                        <th>Total</th>
                        <th>Encomenda</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($rows as $row): ?>
                        <tr>
                            <td><?php echo h(value_or_dash($row['numero'])); ?></td>
                            <td><?php echo h(fmt_datetime($row['data'])); ?></td>
                            <td><?php echo h(value_or_dash($row['cliente_nome'])); ?></td>
                            <td><?php echo render_status_pill((string) $row['estado']); ?></td>
                            <td><?php echo h(fmt_money($row['subtotal'])); ?></td>
                            <td><?php echo h(fmt_money($row['total'])); ?></td>
                            <td><?php echo h(value_or_dash($row['numero_encomenda'])); ?></td>
                        </tr>
                    <?php endforeach; ?>
                    <?php if (!$rows): ?>
                        <tr><td colspan="7" class="empty-row">Sem orcamentos para mostrar.</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</section>
<?php
$content = ob_get_clean();
render_app_page('Orcamentos', $content, 'quotes');
