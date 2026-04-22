<?php

require_once dirname(__DIR__) . '/src/bootstrap.php';
require_login();

$q = query_param('q');
$rows = repo_products($q, 250);

$totalQty = sum_field($rows, 'qty');
$totalStockValue = 0.0;
foreach ($rows as $row) {
    $qty = is_numeric($row['qty'] ?? null) ? (float) $row['qty'] : 0.0;
    $price = is_numeric($row['p_compra'] ?? null) ? (float) $row['p_compra'] : 0.0;
    $totalStockValue += $qty * $price;
}

$actions = '
    <form method="get" class="toolbar">
        <label class="field">
            <span>Pesquisa</span>
            <input type="text" name="q" placeholder="Codigo, descricao ou categoria..." value="' . h($q) . '">
        </label>
        <button type="submit" class="btn btn-primary">
            <span class="btn-icon">' . app_icon('search') . '</span>
            <span>Pesquisar</span>
        </button>
    </form>
';

$aside = '
    <div class="hero-side-card">
        <div>
            <p class="eyebrow">Produto acabado</p>
            <p>Consulta stock atual, tipologia e atualizacao recente das referencias de produto.</p>
        </div>
        <div class="hero-side-grid">
            <div><span>Registos</span><strong>' . h((string) count($rows)) . '</strong></div>
            <div><span>Qtd. total</span><strong>' . h(fmt_num($totalQty)) . '</strong></div>
            <div><span>Valor estimado</span><strong>' . h(fmt_money($totalStockValue)) . '</strong></div>
        </div>
    </div>
';

ob_start();
echo render_page_intro(
    'Stock final',
    'Produtos',
    'Vista elegante do stock de produto acabado, com foco em quantidade, categoria e valor de compra.',
    [
        ['label' => 'Registos', 'value' => (string) count($rows)],
        ['label' => 'Qtd. total', 'value' => fmt_num($totalQty)],
        ['label' => 'Pesquisa', 'value' => $q !== '' ? $q : 'Livre'],
    ],
    $actions,
    $aside
);
?>
<section class="metric-grid compact">
    <?php echo render_metric_card('Produtos visiveis', fmt_compact_number(count($rows)), 'Registos dentro da pesquisa atual.', 'products', 'success'); ?>
    <?php echo render_metric_card('Quantidade total', fmt_num($totalQty), 'Somatorio das quantidades listadas.', 'stack', 'default'); ?>
    <?php echo render_metric_card('Valor estimado', fmt_money($totalStockValue), 'Quantidade x preco de compra.', 'trend', 'warning'); ?>
</section>

<section class="surface">
    <div class="surface-head">
        <div>
            <p class="surface-kicker">Produto acabado</p>
            <h2>Referencias em stock</h2>
            <p class="muted">Consulta de codigo, categoria, quantidade e valor de compra.</p>
        </div>
    </div>
    <div class="surface-body">
        <div class="data-table-wrap">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Codigo</th>
                        <th>Descricao</th>
                        <th>Categoria</th>
                        <th>Tipo</th>
                        <th>Qtd.</th>
                        <th>Alerta</th>
                        <th>Compra</th>
                        <th>Atualizado</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($rows as $row): ?>
                        <tr>
                            <td><?php echo h(value_or_dash($row['codigo'])); ?></td>
                            <td><?php echo h(value_or_dash($row['descricao'])); ?></td>
                            <td><?php echo h(value_or_dash($row['categoria'])); ?></td>
                            <td><?php echo h(value_or_dash($row['tipo'])); ?></td>
                            <td><?php echo h(fmt_num($row['qty'])); ?></td>
                            <td><?php echo h(fmt_num($row['alerta'])); ?></td>
                            <td><?php echo h(fmt_money($row['p_compra'])); ?></td>
                            <td><?php echo h(fmt_datetime($row['atualizado_em'])); ?></td>
                        </tr>
                    <?php endforeach; ?>
                    <?php if (!$rows): ?>
                        <tr><td colspan="8" class="empty-row">Sem produtos para mostrar.</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</section>
<?php
$content = ob_get_clean();
render_app_page('Produtos', $content, 'products');
