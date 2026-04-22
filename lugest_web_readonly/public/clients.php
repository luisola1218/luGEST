<?php

require_once dirname(__DIR__) . '/src/bootstrap.php';
require_login();

$q = query_param('q');
$rows = repo_clients($q, 250);

$actions = '
    <form method="get" class="toolbar">
        <label class="field">
            <span>Pesquisa</span>
            <input type="text" name="q" placeholder="Codigo, nome ou NIF..." value="' . h($q) . '">
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
            <p class="eyebrow">Relacionamento</p>
            <p>Consulta da carteira de clientes com acesso rapido a identificacao fiscal e contactos.</p>
        </div>
        <div class="hero-side-grid">
            <div><span>Clientes visiveis</span><strong>' . h((string) count($rows)) . '</strong></div>
            <div><span>Filtro atual</span><strong>' . h($q !== '' ? $q : 'Todos') . '</strong></div>
        </div>
    </div>
';

ob_start();
echo render_page_intro(
    'Base comercial',
    'Clientes',
    'Leitura organizada da carteira de clientes para apoio a comercial, producao e acompanhamento administrativo.',
    [
        ['label' => 'Registos', 'value' => (string) count($rows)],
        ['label' => 'Pesquisa', 'value' => $q !== '' ? $q : 'Livre'],
    ],
    $actions,
    $aside
);
?>
<section class="surface">
    <div class="surface-head">
        <div>
            <p class="surface-kicker">Base de clientes</p>
            <h2>Carteira visivel</h2>
            <p class="muted">Lista pronta para consulta rapida de codigo, nome, NIF e contactos.</p>
        </div>
    </div>
    <div class="surface-body">
        <div class="data-table-wrap">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Codigo</th>
                        <th>Nome</th>
                        <th>NIF</th>
                        <th>Contacto</th>
                        <th>Email</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($rows as $row): ?>
                        <tr>
                            <td><?php echo h(value_or_dash($row['codigo'])); ?></td>
                            <td><?php echo h(value_or_dash($row['nome'])); ?></td>
                            <td><?php echo h(value_or_dash($row['nif'])); ?></td>
                            <td><?php echo h(value_or_dash($row['contacto'])); ?></td>
                            <td><?php echo h(value_or_dash($row['email'])); ?></td>
                        </tr>
                    <?php endforeach; ?>
                    <?php if (!$rows): ?>
                        <tr><td colspan="5" class="empty-row">Sem clientes para mostrar.</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</section>
<?php
$content = ob_get_clean();
render_app_page('Clientes', $content, 'clients');
