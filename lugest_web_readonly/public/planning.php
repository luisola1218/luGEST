<?php

require_once dirname(__DIR__) . '/src/bootstrap.php';
require_login();

$filters = [
    'q' => query_param('q'),
    'year' => query_param('year', (string) date('Y')),
    'date' => query_param('date'),
    'posto' => query_param('posto'),
    'operacao' => query_param('operacao'),
    'resource' => query_param('resource'),
];

$planningOptions = repo_planning_options();
$operationOptions = repo_workcenter_operation_options();
if ($filters['operacao'] === '' && $operationOptions !== []) {
    $preferredOperation = in_array('Corte Laser', $operationOptions, true) ? 'Corte Laser' : (string) $operationOptions[0];
    $filters['operacao'] = $preferredOperation;
}

$anchorDate = $filters['date'] !== ''
    ? $filters['date']
    : (repo_planning_anchor_date([
        'year' => $filters['year'],
        'q' => $filters['q'],
        'posto' => $filters['posto'],
        'operacao' => $filters['operacao'],
    ]) ?? date('Y-m-d'));

$anchor = new DateTimeImmutable($anchorDate);
$weekStart = $anchor->modify('monday this week');
$weekEnd = $weekStart->modify('+6 days');
$dayNames = ['Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sab', 'Dom'];

$queryFilters = [
    'q' => $filters['q'],
    'year' => $filters['year'],
    'posto' => $filters['posto'],
    'operacao' => $filters['operacao'],
    'date_from' => $weekStart->format('Y-m-d'),
    'date_to' => $weekEnd->format('Y-m-d'),
];

$rowsAll = repo_planning_enrich_rows(repo_planning($queryFilters, 600));
usort($rowsAll, static function (array $left, array $right): int {
    $leftKey = ($left['data_planeada'] ?? '') . ' ' . ($left['inicio'] ?? '') . ' ' . ($left['resource'] ?? '');
    $rightKey = ($right['data_planeada'] ?? '') . ' ' . ($right['inicio'] ?? '') . ' ' . ($right['resource'] ?? '');
    return strcmp($leftKey, $rightKey);
});

$availableResources = repo_workcenter_resource_options($filters['operacao']);
$hasUndefinedResources = false;
foreach ($rowsAll as $row) {
    $resource = trim((string) ($row['resource'] ?? ''));
    if ($resource !== '' && !in_array($resource, $availableResources, true) && strcasecmp($resource, 'Sem recurso definido') !== 0) {
        $availableResources[] = $resource;
    }
    if (strcasecmp($resource, 'Sem recurso definido') === 0) {
        $hasUndefinedResources = true;
    }
}
usort($availableResources, static fn(string $left, string $right): int => strcasecmp($left, $right));
if ($hasUndefinedResources && !in_array('Sem recurso definido', $availableResources, true)) {
    $availableResources[] = 'Sem recurso definido';
}

$rows = $rowsAll;
if ($filters['resource'] !== '') {
    $selectedResource = strtolower($filters['resource']);
    $rows = array_values(array_filter(
        $rowsAll,
        static fn(array $row): bool => strtolower(trim((string) ($row['resource'] ?? ''))) === $selectedResource
    ));
}

$countBlocks = count($rows);
$totalDuration = sum_field($rows, 'duracao_min');
$uniqueOrders = [];
foreach ($rows as $row) {
    $order = trim((string) ($row['encomenda_numero'] ?? ''));
    if ($order !== '') {
        $uniqueOrders[$order] = true;
    }
}
$countOrders = count($uniqueOrders);

$resourceCountMap = [];
$operationCountMap = [];
foreach ($rowsAll as $row) {
    $resourceKey = strtolower(trim((string) ($row['resource'] ?? '')));
    if ($resourceKey !== '') {
        $resourceCountMap[$resourceKey] = ($resourceCountMap[$resourceKey] ?? 0) + 1;
    }
    $operationKey = strtolower(trim((string) ($row['operacao'] ?? '')));
    if ($operationKey !== '') {
        $operationCountMap[$operationKey] = ($operationCountMap[$operationKey] ?? 0) + 1;
    }
}

$visibleResources = $filters['resource'] !== '' ? [$filters['resource']] : $availableResources;
if ($visibleResources === []) {
    $visibleResources = ['Sem recurso definido'];
}

$parseTimeMinutes = static function (string $value): ?int {
    $text = trim($value);
    if ($text === '' || !preg_match('/^(\d{1,2}):(\d{2})$/', $text, $matches)) {
        return null;
    }
    return ((int) $matches[1] * 60) + (int) $matches[2];
};

$formatTimeMinutes = static function (int $minutes): string {
    $hours = intdiv($minutes, 60);
    $mins = $minutes % 60;
    return str_pad((string) $hours, 2, '0', STR_PAD_LEFT) . ':' . str_pad((string) $mins, 2, '0', STR_PAD_LEFT);
};

$timelineSource = $rows !== [] ? $rows : $rowsAll;
$minMinutes = 8 * 60;
$maxMinutes = 18 * 60;
foreach ($timelineSource as $row) {
    $startMinutes = $parseTimeMinutes((string) ($row['inicio'] ?? ''));
    if ($startMinutes === null) {
        continue;
    }
    $duration = is_numeric($row['duracao_min'] ?? null) ? (int) round((float) $row['duracao_min']) : 0;
    $endMinutes = $startMinutes + max(30, $duration);
    $minMinutes = min($minMinutes, $startMinutes);
    $maxMinutes = max($maxMinutes, $endMinutes);
}
$slotMinutes = 30;
$slotHeight = 28;
$timelineStart = max(0, intdiv($minMinutes, $slotMinutes) * $slotMinutes);
$timelineEnd = min(24 * 60, (int) ceil($maxMinutes / $slotMinutes) * $slotMinutes);
if ($timelineEnd <= $timelineStart) {
    $timelineEnd = $timelineStart + (10 * 60);
}
$timelineHeight = max(560, (int) ((($timelineEnd - $timelineStart) / $slotMinutes) * $slotHeight));
$timeSlots = [];
for ($minutes = $timelineStart; $minutes <= $timelineEnd; $minutes += $slotMinutes) {
    $timeSlots[] = [
        'minutes' => $minutes,
        'label' => $formatTimeMinutes($minutes),
        'offset' => (int) ((($minutes - $timelineStart) / $slotMinutes) * $slotHeight),
    ];
}

$days = [];
for ($index = 0; $index < 7; $index++) {
    $dayDate = $weekStart->modify('+' . $index . ' days');
    $dateKey = $dayDate->format('Y-m-d');
    $resourceMap = [];
    foreach ($visibleResources as $resourceName) {
        $resourceMap[$resourceName] = [];
    }
    foreach ($rows as $row) {
        if ((string) ($row['data_planeada'] ?? '') !== $dateKey) {
            continue;
        }
        $rowResource = trim((string) ($row['resource'] ?? 'Sem recurso definido'));
        if (!array_key_exists($rowResource, $resourceMap)) {
            $resourceMap[$rowResource] = [];
        }
        $resourceMap[$rowResource][] = $row;
    }
    $days[] = [
        'label' => $dayNames[$index],
        'date' => $dayDate,
        'resources' => $resourceMap,
    ];
}

$buildUrl = static function (array $overrides = []) use ($filters): string {
    $params = [
        'q' => $filters['q'],
        'year' => $filters['year'],
        'date' => $filters['date'],
        'posto' => $filters['posto'],
        'operacao' => $filters['operacao'],
        'resource' => $filters['resource'],
    ];
    foreach ($overrides as $key => $value) {
        if ($value === null) {
            unset($params[$key]);
            continue;
        }
        $params[$key] = $value;
    }
    $params = array_filter($params, static fn($value): bool => trim((string) $value) !== '');
    return 'planning.php' . ($params ? '?' . http_build_query($params) : '');
};

$prevWeekUrl = $buildUrl(['date' => $weekStart->modify('-7 days')->format('Y-m-d')]);
$nextWeekUrl = $buildUrl(['date' => $weekStart->modify('+7 days')->format('Y-m-d')]);
$currentWeekUrl = $buildUrl(['date' => date('Y-m-d')]);

$renderSwitchPill = static function (string $label, string $url, bool $active, string $countText = ''): string {
    return '<a class="switch-pill' . ($active ? ' active' : '') . '" href="' . h($url) . '"><span>' . h($label) . '</span>' . ($countText !== '' ? '<em>' . h($countText) . '</em>' : '') . '</a>';
};

$operationPills = [];
foreach ($operationOptions as $operationName) {
    $count = $operationCountMap[strtolower($operationName)] ?? 0;
    $operationPills[] = $renderSwitchPill(
        $operationName,
        $buildUrl(['operacao' => $operationName, 'resource' => null]),
        strcasecmp($filters['operacao'], $operationName) === 0,
        (string) $count
    );
}

$resourcePills = [];
$resourcePills[] = $renderSwitchPill(
    'Todas as maquinas',
    $buildUrl(['resource' => null]),
    $filters['resource'] === '',
    (string) count($rowsAll)
);
foreach ($availableResources as $resourceName) {
    $count = $resourceCountMap[strtolower($resourceName)] ?? 0;
    $resourcePills[] = $renderSwitchPill(
        $resourceName,
        $buildUrl(['resource' => $resourceName]),
        strcasecmp($filters['resource'], $resourceName) === 0,
        (string) $count
    );
}

ob_start();
?>
<section class="surface planning-control-surface">
    <div class="surface-head">
        <div>
            <p class="surface-kicker">Quadro de producao</p>
            <h2>Planeamento semanal</h2>
            <p class="muted">Vista em grelha horaria, com dias e maquinas em colunas, para ler o planeamento como um quadro real de fabrica.</p>
        </div>
        <div class="board-nav">
            <button type="button" class="btn btn-secondary btn-linkish" id="planning-focus-toggle">
                <span>Maximizar quadro</span>
            </button>
            <a class="btn btn-secondary btn-linkish" href="<?php echo h($prevWeekUrl); ?>">
                <span class="btn-icon"><?php echo app_icon('arrow-left'); ?></span>
                <span>Semana anterior</span>
            </a>
            <a class="btn btn-secondary btn-linkish" href="<?php echo h($nextWeekUrl); ?>">
                <span class="btn-icon"><?php echo app_icon('arrow-right'); ?></span>
                <span>Semana seguinte</span>
            </a>
            <a class="btn btn-secondary btn-linkish" href="<?php echo h($currentWeekUrl); ?>">
                <span>Atual</span>
            </a>
        </div>
    </div>
    <div class="surface-body planning-tool-body">
        <div class="switch-stack">
            <div class="switch-row">
                <div class="switch-label">Operacao</div>
                <div class="switch-pills"><?php echo implode('', $operationPills); ?></div>
            </div>
            <div class="switch-row">
                <div class="switch-label">Maquinas</div>
                <div class="switch-pills"><?php echo implode('', $resourcePills); ?></div>
            </div>
        </div>

        <form method="get" class="toolbar">
            <input type="hidden" name="operacao" value="<?php echo h($filters['operacao']); ?>">
            <label class="field">
                <span>Pesquisa</span>
                <input type="text" name="q" placeholder="Encomenda, material, bloco ou chapa..." value="<?php echo h($filters['q']); ?>">
            </label>
            <label class="field compact">
                <span>Ano</span>
                <input type="number" name="year" value="<?php echo h($filters['year']); ?>" min="2020" max="2100">
            </label>
            <label class="field compact">
                <span>Semana de</span>
                <input type="date" name="date" value="<?php echo h($anchor->format('Y-m-d')); ?>">
            </label>
            <label class="field compact">
                <span>Posto / grupo</span>
                <select name="posto">
                    <option value="">Todos</option>
                    <?php foreach ($planningOptions['postos'] as $row): ?>
                        <option value="<?php echo h((string) $row['valor']); ?>"<?php echo ((string) $row['valor'] === $filters['posto']) ? ' selected' : ''; ?>><?php echo h((string) $row['valor']); ?></option>
                    <?php endforeach; ?>
                </select>
            </label>
            <label class="field compact">
                <span>Maquina</span>
                <select name="resource">
                    <option value="">Todas</option>
                    <?php foreach ($availableResources as $resourceName): ?>
                        <option value="<?php echo h($resourceName); ?>"<?php echo strcasecmp($resourceName, $filters['resource']) === 0 ? ' selected' : ''; ?>><?php echo h($resourceName); ?></option>
                    <?php endforeach; ?>
                </select>
            </label>
            <button type="submit" class="btn btn-primary">
                <span class="btn-icon"><?php echo app_icon('filter'); ?></span>
                <span>Atualizar quadro</span>
            </button>
        </form>

        <div class="planning-summary-strip">
            <div><span>Semana</span><strong><?php echo h($weekStart->format('d/m')) . ' - ' . h($weekEnd->format('d/m')); ?></strong></div>
            <div><span>Operacao</span><strong><?php echo h($filters['operacao']); ?></strong></div>
            <div><span>Blocos</span><strong><?php echo h((string) $countBlocks); ?></strong></div>
            <div><span>Encomendas</span><strong><?php echo h((string) $countOrders); ?></strong></div>
            <div><span>Carga</span><strong><?php echo h(fmt_time_minutes($totalDuration)); ?></strong></div>
        </div>
    </div>
</section>

<section class="surface planning-timeline-surface">
    <div class="surface-head">
        <div>
            <p class="surface-kicker">Grelha horaria</p>
            <h2><?php echo h($filters['resource'] !== '' ? $filters['resource'] : 'Todas as maquinas'); ?></h2>
            <p class="muted">Cada dia mostra as maquinas em colunas, com o eixo das horas na vertical e os blocos encaixados pela duracao real.</p>
        </div>
    </div>
    <div class="surface-body">
        <div class="timeline-scroll">
            <div class="timeline-week">
                <?php foreach ($days as $day): ?>
                    <section class="timeline-day-panel" style="--machine-count: <?php echo h((string) count($day['resources'])); ?>;">
                        <header class="timeline-day-header">
                            <p><?php echo h((string) $day['label']); ?></p>
                            <h3><?php echo h($day['date']->format('d/m/Y')); ?></h3>
                        </header>
                        <div
                            class="timeline-day-grid"
                            style="--machine-count: <?php echo h((string) count($day['resources'])); ?>; --timeline-height: <?php echo h((string) $timelineHeight); ?>px; --slot-height: <?php echo h((string) $slotHeight); ?>px;"
                        >
                            <div class="timeline-axis-head">Horas</div>
                            <?php foreach (array_keys($day['resources']) as $resourceName): ?>
                                <div class="timeline-machine-head"><?php echo h($resourceName); ?></div>
                            <?php endforeach; ?>

                            <div class="timeline-axis-column">
                                <?php foreach ($timeSlots as $slot): ?>
                                    <div class="timeline-axis-label" style="top: <?php echo h((string) $slot['offset']); ?>px;">
                                        <?php echo h((string) $slot['label']); ?>
                                    </div>
                                <?php endforeach; ?>
                            </div>

                            <?php foreach ($day['resources'] as $resourceName => $resourceRows): ?>
                                <div class="timeline-track">
                                    <?php foreach ($resourceRows as $row): ?>
                                        <?php
                                        $color = trim((string) ($row['color'] ?? ''));
                                        $solidColor = $color !== '' ? $color : '#d4b11a';
                                        $softColor = hex_to_rgba($solidColor, 0.9);
                                        $lineColor = hex_to_rgba($solidColor, 0.28);
                                        $startMinutes = $parseTimeMinutes((string) ($row['inicio'] ?? '')) ?? $timelineStart;
                                        $durationMinutes = max(30, is_numeric($row['duracao_min'] ?? null) ? (int) round((float) $row['duracao_min']) : 30);
                                        $topOffset = max(0, (($startMinutes - $timelineStart) / $slotMinutes) * $slotHeight);
                                        $blockHeight = max(24, ($durationMinutes / $slotMinutes) * $slotHeight);
                                        ?>
                                        <a
                                            class="timeline-block"
                                            href="order.php?numero=<?php echo urlencode((string) $row['encomenda_numero']); ?>"
                                            style="top: <?php echo h((string) $topOffset); ?>px; height: <?php echo h((string) $blockHeight); ?>px; --block-solid: <?php echo h($solidColor); ?>; --block-soft: <?php echo h($softColor); ?>; --block-line: <?php echo h($lineColor); ?>;"
                                        >
                                            <strong><?php echo h((string) $row['encomenda_numero']); ?></strong>
                                            <span><?php echo h(value_or_dash($row['material'])); ?> <?php echo h(value_or_dash($row['espessura'])); ?>mm</span>
                                            <span><?php echo h(value_or_dash($row['inicio'])); ?> | <?php echo h(fmt_time_minutes($row['duracao_min'])); ?></span>
                                            <?php if ((int) ($row['resource_defined'] ?? 0) === 0 && trim((string) ($row['resource_group'] ?? '')) !== ''): ?>
                                                <em>Posto: <?php echo h((string) $row['resource_group']); ?></em>
                                            <?php endif; ?>
                                        </a>
                                    <?php endforeach; ?>
                                </div>
                            <?php endforeach; ?>
                        </div>
                    </section>
                <?php endforeach; ?>
            </div>
        </div>
        <?php if (!$rows): ?>
            <div class="planning-empty" style="margin-top: 16px;">Sem blocos para mostrar nesta semana.</div>
        <?php endif; ?>
    </div>
</section>

<section class="surface">
    <div class="surface-head">
        <div>
            <p class="surface-kicker">Tabela de apoio</p>
            <h2>Blocos visiveis</h2>
            <p class="muted">Lista detalhada da vista atual, para pesquisa rapida e abertura da encomenda.</p>
        </div>
    </div>
    <div class="surface-body">
        <div class="data-table-wrap">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Data</th>
                        <th>Inicio</th>
                        <th>Encomenda</th>
                        <th>Maquina</th>
                        <th>Material</th>
                        <th>Esp.</th>
                        <th>Duracao</th>
                        <th>Bloco</th>
                    </tr>
                </thead>
                <tbody>
                    <?php foreach ($rows as $row): ?>
                        <tr>
                            <td><?php echo h(fmt_date($row['data_planeada'])); ?></td>
                            <td><?php echo h(value_or_dash($row['inicio'])); ?></td>
                            <td>
                                <a class="table-link" href="order.php?numero=<?php echo urlencode((string) $row['encomenda_numero']); ?>">
                                    <?php echo h((string) $row['encomenda_numero']); ?>
                                </a>
                            </td>
                            <td><?php echo h(value_or_dash($row['resource'])); ?></td>
                            <td><?php echo h(value_or_dash($row['material'])); ?></td>
                            <td><?php echo h(value_or_dash($row['espessura'])); ?></td>
                            <td><?php echo h(fmt_time_minutes($row['duracao_min'])); ?></td>
                            <td><?php echo render_status_pill((string) $row['bloco_id']); ?></td>
                        </tr>
                    <?php endforeach; ?>
                    <?php if (!$rows): ?>
                        <tr><td colspan="8" class="empty-row">Sem blocos para mostrar nesta vista.</td></tr>
                    <?php endif; ?>
                </tbody>
            </table>
        </div>
    </div>
</section>
<?php
$content = ob_get_clean();
$content .= <<<HTML
<script>
(function () {
    var storageKey = 'lugestPlanningFocus';
    var button = document.getElementById('planning-focus-toggle');
    if (!button) {
        return;
    }

    function applyState(enabled) {
        document.body.classList.toggle('planning-focus', enabled);
        button.querySelector('span').textContent = enabled ? 'Sair do maximo' : 'Maximizar quadro';
    }

    var enabled = window.localStorage.getItem(storageKey) === '1';
    applyState(enabled);

    button.addEventListener('click', function () {
        enabled = !enabled;
        window.localStorage.setItem(storageKey, enabled ? '1' : '0');
        applyState(enabled);
    });
})();
</script>
HTML;
render_app_page('Planeamento', $content, 'planning');
