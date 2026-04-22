<?php

require_once dirname(__DIR__) . '/src/bootstrap.php';
require_login();

$q = query_param('q');
$rows = repo_materials($q, 250);

$totalQty = sum_field($rows, 'quantidade');
$reservedQty = sum_field($rows, 'reservado');
$latestUpdate = max_field($rows, 'atualizado_em');

$actions = '
    <form method="get" class="toolbar">
        <label class="field">
            <span>Pesquisa</span>
            <input type="text" name="q" placeholder="Material, espessura, lote ou localizacao..." value="' . h($q) . '">
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
            <p class="eyebrow">Stock base</p>
            <p>Consulta materia-prima e retalhos com localizacao, reserva e atualizacao mais recente.</p>
        </div>
        <div class="hero-side-grid">
            <div><span>Registos</span><strong>' . h((string) count($rows)) . '</strong></div>
            <div><span>Reservado</span><strong>' . h(fmt_num($reservedQty)) . '</strong></div>
            <div><span>Ultima leitura</span><strong>' . h(fmt_datetime($latestUpdate)) . '</strong></div>
        </div>
    </div>
';

ob_start();
echo render_page_intro(
    'Stock metalico',
    'Materia-prima',
    'Consulta profissional do stock de materiais e retalhos, com visibilidade de quantidades, reservas e localizacao.',
    [
        ['label' => 'Registos', 'value' => (string) count($rows)],
        ['label' => 'Qtd. total', 'value' => fmt_num($totalQty)],
        ['label' => 'Reservado', 'value' => fmt_num($reservedQty)],
    ],
    $actions,
    $aside
);
?>
<section class="metric-grid compact">
    <?php echo render_metric_card('Materiais visiveis', fmt_compact_number(count($rows)), 'Registos listados para a pesquisa atual.', 'materials', 'default'); ?>
    <?php echo render_metric_card('Quantidade total', fmt_num($totalQty), 'Somatorio das quantidades disponiveis.', 'stack', 'success'); ?>
    <?php echo render_metric_card('Reservas', fmt_num($reservedQty), 'Volume atualmente reservado.', 'shield', 'warning'); ?>
</section>

<section class="surface">
    <div class="surface-head">
        <div>
            <p class="surface-kicker">Inventario base</p>
            <h2>Stock de materia-prima</h2>
            <p class="muted">Leitura de material, formato, reservas e localizacao.</p>
        </div>
    </div>
    <div class="surface-body">
        <div class="data-table-wrap">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Material</th>
                        <th>Esp.</th>
                        <th>Formato</th>
                        <th>Qtd.</th>
                        <th>Reservado</th>
                        <th>Localizacao</th>
                        <th>Atualizado</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($rows as $row): ?>
                        <tr>
                            <td><?php echo h($row['id']); ?></td>
                            <td><?php echo h(value_or_dash($row['material'])); ?></td>
                            <td><?php echo h(value_or_dash($row['espessura'])); ?></td>
                            <td><?php echo h(value_or_dash($row['formato'])); ?></td>
                            <td><?php echo h(fmt_num($row['quantidade'])); ?></td>
                            <td><?php echo h(fmt_num($row['reservado'])); ?></td>
                            <td><?php echo h(value_or_dash($row['localizacao'])); ?></td>
                            <td><?php echo h(fmt_datetime($row['atualizado_em'])); ?></td>
                        </tr>
                    <?php endforeach; ?>
                    <?php if (!$rows): ?>
                        <tr><td colspan="8" class="empty-row">Sem materiais para mostrar.</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</section>
<?php
$content = ob_get_clean();
render_app_page('Materia-prima', $content, 'materials');
