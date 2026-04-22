<?php

require_once dirname(__DIR__) . '/src/bootstrap.php';
require_login();

$numero = query_param('numero');
if ($numero === '') {
    redirect('orders.php');
}

$order = repo_order($numero);
if (!$order) {
    http_response_code(404);
    $content = '<section class="surface"><div class="surface-head"><div><p class="surface-kicker">Detalhe indisponivel</p><h1>Encomenda nao encontrada</h1><p class="muted">A encomenda pedida nao existe na base visivel do portal.</p></div></div></section>';
    render_app_page('Encomenda', $content, 'orders');
    exit;
}

$pieces = repo_order_pieces($numero);
$planning = repo_order_planning($numero);

$actions = '
    <div class="hero-actions">
        <a class="btn btn-secondary" href="orders.php">
            <span class="btn-icon">' . app_icon('arrow-right') . '</span>
            <span>Voltar a encomendas</span>
        </a>
    </div>
';

$aside = '
    <div class="hero-side-card">
        <div>
            <p class="eyebrow">Estado da ordem</p>
            <p>Consulta detalhada da encomenda, com visao de pecas, observacoes e planeamento associado.</p>
        </div>
        <div class="hero-side-grid">
            <div><span>Cliente</span><strong>' . h(non_empty_string($order['cliente_nome'] ?? '-')) . '</strong></div>
            <div><span>Estado</span><strong>' . h(non_empty_string($order['estado'] ?? '-')) . '</strong></div>
            <div><span>Entrega</span><strong>' . h(fmt_date($order['data_entrega'] ?? null)) . '</strong></div>
        </div>
    </div>
';

ob_start();
echo render_page_intro(
    'Detalhe de encomenda',
    (string) $order['numero'],
    'Painel de leitura para acompanhamento de cliente, pecas registadas e blocos de producao ligados a esta encomenda.',
    [
        ['label' => 'Cliente', 'value' => non_empty_string($order['cliente_nome'] ?? '-')],
        ['label' => 'Estado', 'value' => non_empty_string($order['estado'] ?? '-')],
        ['label' => 'Entrega', 'value' => fmt_date($order['data_entrega'] ?? null)],
    ],
    $actions,
    $aside
);
?>
<section class="metric-grid">
    <?php echo render_metric_card('Entrega', fmt_date($order['data_entrega'] ?? null), 'Data prevista para entrega da encomenda.', 'clock', 'default'); ?>
    <?php echo render_metric_card('Pecas', fmt_num($order['pecas_total'] ?? 0), 'Total de pecas ligadas a esta encomenda.', 'products', 'success'); ?>
    <?php echo render_metric_card('Blocos', fmt_num($order['blocos_total'] ?? 0), 'Numero de blocos ativos associados.', 'planning', 'warning'); ?>
    <?php echo render_metric_card('Tempo estimado', fmt_time_minutes($order['tempo_estimado'] ?? 0), 'Tempo total previsto no registo da encomenda.', 'trend', 'default'); ?>
</section>

<section class="surface">
    <div class="surface-head">
        <div>
            <p class="surface-kicker">Contexto principal</p>
            <h2>Resumo da encomenda</h2>
            <p class="muted">Leitura limpa dos dados essenciais e notas associadas.</p>
        </div>
    </div>
    <div class="surface-body detail-grid">
        <div class="detail-card">
            <strong>Cliente</strong>
            <span><?php echo h(non_empty_string($order['cliente_nome'] ?? '-')); ?></span>
        </div>
        <div class="detail-card">
            <strong>Numero orcamento</strong>
            <span><?php echo h(non_empty_string($order['numero_orcamento'] ?? '-')); ?></span>
        </div>
        <div class="detail-card">
            <strong>Entrega</strong>
            <span><?php echo h(fmt_date($order['data_entrega'] ?? null)); ?></span>
        </div>
        <div class="detail-card">
            <strong>Estado</strong>
            <span><?php echo render_status_pill((string) ($order['estado'] ?? '-')); ?></span>
        </div>
        <div class="detail-card wide">
            <strong>Nota cliente</strong>
            <span><?php echo h(non_empty_string($order['nota_cliente'] ?? '-', '-')); ?></span>
        </div>
        <div class="detail-card wide">
            <strong>Observacoes</strong>
            <span><?php echo h(non_empty_string($order['observacoes'] ?? '-', '-')); ?></span>
        </div>
    </div>
</section>

<section class="surface">
    <div class="surface-head">
        <div>
            <p class="surface-kicker">Estrutura produtiva</p>
            <h2>Pecas registadas</h2>
            <p class="muted">Referencias, materiais e operacoes associadas a esta encomenda.</p>
        </div>
    </div>
    <div class="surface-body">
        <div class="data-table-wrap">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>ID</th>
                        <th>Ref. interna</th>
                        <th>Ref. externa</th>
                        <th>Material</th>
                        <th>Esp.</th>
                        <th>Qtd.</th>
                        <th>Estado</th>
                        <th>Operacoes</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($pieces as $row): ?>
                        <tr>
                            <td><?php echo h($row['id']); ?></td>
                            <td><?php echo h(value_or_dash($row['ref_interna'])); ?></td>
                            <td><?php echo h(value_or_dash($row['ref_externa'])); ?></td>
                            <td><?php echo h(value_or_dash($row['material'])); ?></td>
                            <td><?php echo h(value_or_dash($row['espessura'])); ?></td>
                            <td><?php echo h(fmt_num($row['quantidade_pedida'])); ?></td>
                            <td><?php echo render_status_pill((string) $row['estado']); ?></td>
                            <td><?php echo h(value_or_dash($row['operacoes'])); ?></td>
                        </tr>
                    <?php endforeach; ?>
                    <?php if (!$pieces): ?>
                        <tr><td colspan="8" class="empty-row">Sem pecas registadas.</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</section>

<section class="surface">
    <div class="surface-head">
        <div>
            <p class="surface-kicker">Ligacao ao quadro</p>
            <h2>Planeamento associado</h2>
            <p class="muted">Blocos ativos ligados diretamente a esta encomenda.</p>
        </div>
    </div>
    <div class="surface-body">
        <div class="data-table-wrap">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Bloco</th>
                        <th>Material</th>
                        <th>Esp.</th>
                        <th>Data</th>
                        <th>Inicio</th>
                        <th>Duracao</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($planning as $row): ?>
                        <tr>
                            <td><?php echo render_status_pill((string) $row['bloco_id']); ?></td>
                            <td><?php echo h(value_or_dash($row['material'])); ?></td>
                            <td><?php echo h(value_or_dash($row['espessura'])); ?></td>
                            <td><?php echo h(fmt_date($row['data_planeada'])); ?></td>
                            <td><?php echo h(value_or_dash($row['inicio'])); ?></td>
                            <td><?php echo h(fmt_time_minutes($row['duracao_min'])); ?></td>
                        </tr>
                    <?php endforeach; ?>
                    <?php if (!$planning): ?>
                        <tr><td colspan="6" class="empty-row">Sem blocos ativos ligados a esta encomenda.</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</section>
<?php
$content = ob_get_clean();
render_app_page('Detalhe da encomenda', $content, 'orders');
