<?php

function repo_dashboard_stats(string $year): array {
    return [
        [
            'label' => 'Encomendas',
            'value' => (string) db_value(
                'SELECT COUNT(*) FROM encomendas WHERE CAST(ano AS CHAR) = ? OR YEAR(COALESCE(data_entrega, data_criacao)) = ?',
                [$year, $year]
            ),
            'note' => 'Registos no ano selecionado',
        ],
        [
            'label' => 'Planeamento',
            'value' => (string) db_value(
                'SELECT COUNT(*) FROM plano WHERE CAST(ano AS CHAR) = ? OR YEAR(data_planeada) = ?',
                [$year, $year]
            ),
            'note' => 'Blocos ativos',
        ],
        [
            'label' => 'Produtos',
            'value' => (string) db_value('SELECT COUNT(*) FROM produtos'),
            'note' => 'Referencias em stock',
        ],
        [
            'label' => 'Materia-prima',
            'value' => (string) db_value('SELECT COUNT(*) FROM materiais'),
            'note' => 'Registos de stock',
        ],
    ];
}

function repo_recent_orders(string $year, int $limit = 8): array {
    $limit = max(1, min($limit, 50));
    $sql = "
        SELECT
            e.numero,
            COALESCE(c.nome, e.cliente_codigo, '-') AS cliente_nome,
            COALESCE(e.estado, '-') AS estado,
            e.data_entrega,
            COALESCE(p.pecas_total, 0) AS pecas_total,
            COALESCE(pl.blocos_total, 0) AS blocos_total
        FROM encomendas e
        LEFT JOIN clientes c ON c.codigo = e.cliente_codigo
        LEFT JOIN (
            SELECT encomenda_numero, COUNT(*) AS pecas_total
            FROM pecas
            GROUP BY encomenda_numero
        ) p ON p.encomenda_numero = e.numero
        LEFT JOIN (
            SELECT encomenda_numero, COUNT(*) AS blocos_total
            FROM plano
            GROUP BY encomenda_numero
        ) pl ON pl.encomenda_numero = e.numero
        WHERE CAST(e.ano AS CHAR) = ? OR YEAR(COALESCE(e.data_entrega, e.data_criacao)) = ?
        ORDER BY COALESCE(e.data_entrega, DATE(e.data_criacao)) DESC, e.numero DESC
        LIMIT {$limit}
    ";
    return db_all($sql, [$year, $year]);
}

function repo_dashboard_status_breakdown(string $year, int $limit = 6): array {
    $limit = max(1, min($limit, 10));
    $sql = "
        SELECT
            COALESCE(NULLIF(TRIM(estado), ''), 'Sem estado') AS estado,
            COUNT(*) AS total
        FROM encomendas
        WHERE CAST(ano AS CHAR) = ? OR YEAR(COALESCE(data_entrega, data_criacao)) = ?
        GROUP BY COALESCE(NULLIF(TRIM(estado), ''), 'Sem estado')
        ORDER BY total DESC, estado ASC
        LIMIT {$limit}
    ";
    return db_all($sql, [$year, $year]);
}

function repo_dashboard_planning_by_day(string $year, int $limit = 7): array {
    $limit = max(1, min($limit, 14));
    $sql = "
        SELECT
            data_planeada,
            COUNT(*) AS blocos,
            COALESCE(SUM(COALESCE(duracao_min, 0)), 0) AS duracao_total
        FROM plano
        WHERE CAST(ano AS CHAR) = ? OR YEAR(data_planeada) = ?
        GROUP BY data_planeada
        ORDER BY data_planeada DESC
        LIMIT {$limit}
    ";
    $rows = db_all($sql, [$year, $year]);
    return array_reverse($rows);
}

function repo_dashboard_top_clients(string $year, int $limit = 5): array {
    $limit = max(1, min($limit, 10));
    $sql = "
        SELECT
            COALESCE(c.nome, e.cliente_codigo, '-') AS cliente_nome,
            COUNT(*) AS total,
            COALESCE(SUM(COALESCE(e.tempo_estimado, 0)), 0) AS tempo_total
        FROM encomendas e
        LEFT JOIN clientes c ON c.codigo = e.cliente_codigo
        WHERE CAST(e.ano AS CHAR) = ? OR YEAR(COALESCE(e.data_entrega, e.data_criacao)) = ?
        GROUP BY COALESCE(c.nome, e.cliente_codigo, '-')
        ORDER BY total DESC, cliente_nome ASC
        LIMIT {$limit}
    ";
    return db_all($sql, [$year, $year]);
}

function repo_orders(array $filters, int $limit = 250): array {
    $where = [];
    $params = [];

    if (($filters['year'] ?? '') !== '') {
        $where[] = '(CAST(e.ano AS CHAR) = ? OR YEAR(COALESCE(e.data_entrega, e.data_criacao)) = ?)';
        $params[] = $filters['year'];
        $params[] = $filters['year'];
    }
    if (($filters['estado'] ?? '') !== '') {
        $where[] = 'e.estado LIKE ?';
        $params[] = '%' . $filters['estado'] . '%';
    }
    if (($filters['q'] ?? '') !== '') {
        $where[] = '(e.numero LIKE ? OR COALESCE(c.nome, "") LIKE ? OR COALESCE(e.nota_cliente, "") LIKE ?)';
        $like = '%' . $filters['q'] . '%';
        $params[] = $like;
        $params[] = $like;
        $params[] = $like;
    }

    $sql = "
        SELECT
            e.numero,
            COALESCE(c.nome, e.cliente_codigo, '-') AS cliente_nome,
            COALESCE(e.estado, '-') AS estado,
            e.data_entrega,
            COALESCE(e.tempo_estimado, 0) AS tempo_estimado,
            COALESCE(p.pecas_total, 0) AS pecas_total,
            COALESCE(pl.blocos_total, 0) AS blocos_total
        FROM encomendas e
        LEFT JOIN clientes c ON c.codigo = e.cliente_codigo
        LEFT JOIN (
            SELECT encomenda_numero, COUNT(*) AS pecas_total
            FROM pecas
            GROUP BY encomenda_numero
        ) p ON p.encomenda_numero = e.numero
        LEFT JOIN (
            SELECT encomenda_numero, COUNT(*) AS blocos_total
            FROM plano
            GROUP BY encomenda_numero
        ) pl ON pl.encomenda_numero = e.numero
    ";

    if ($where) {
        $sql .= ' WHERE ' . implode(' AND ', $where);
    }

    $limit = max(1, min($limit, 500));
    $sql .= " ORDER BY COALESCE(e.data_entrega, DATE(e.data_criacao)) DESC, e.numero DESC LIMIT {$limit}";
    return db_all($sql, $params);
}

function repo_order(string $numero): ?array {
    $sql = "
        SELECT
            e.*,
            COALESCE(c.nome, e.cliente_codigo, '-') AS cliente_nome,
            COALESCE(p.pecas_total, 0) AS pecas_total,
            COALESCE(pl.blocos_total, 0) AS blocos_total
        FROM encomendas e
        LEFT JOIN clientes c ON c.codigo = e.cliente_codigo
        LEFT JOIN (
            SELECT encomenda_numero, COUNT(*) AS pecas_total
            FROM pecas
            GROUP BY encomenda_numero
        ) p ON p.encomenda_numero = e.numero
        LEFT JOIN (
            SELECT encomenda_numero, COUNT(*) AS blocos_total
            FROM plano
            GROUP BY encomenda_numero
        ) pl ON pl.encomenda_numero = e.numero
        WHERE e.numero = ?
        LIMIT 1
    ";
    return db_one($sql, [$numero]);
}

function repo_order_pieces(string $numero): array {
    return db_all(
        'SELECT id, ref_interna, ref_externa, material, espessura, quantidade_pedida, estado, operacoes FROM pecas WHERE encomenda_numero = ? ORDER BY ref_interna ASC, id ASC',
        [$numero]
    );
}

function repo_order_planning(string $numero): array {
    return db_all(
        'SELECT bloco_id, material, espessura, data_planeada, inicio, duracao_min FROM plano WHERE encomenda_numero = ? ORDER BY data_planeada ASC, inicio ASC',
        [$numero]
    );
}

function repo_planning(array $filters, int $limit = 300): array {
    $where = [];
    $params = [];

    if (($filters['year'] ?? '') !== '') {
        $where[] = '(CAST(ano AS CHAR) = ? OR YEAR(data_planeada) = ?)';
        $params[] = $filters['year'];
        $params[] = $filters['year'];
    }
    if (($filters['date'] ?? '') !== '') {
        $where[] = 'data_planeada = ?';
        $params[] = $filters['date'];
    }
    if (($filters['date_from'] ?? '') !== '') {
        $where[] = 'data_planeada >= ?';
        $params[] = $filters['date_from'];
    }
    if (($filters['date_to'] ?? '') !== '') {
        $where[] = 'data_planeada <= ?';
        $params[] = $filters['date_to'];
    }
    if (($filters['posto'] ?? '') !== '') {
        $where[] = 'posto = ?';
        $params[] = $filters['posto'];
    }
    if (($filters['operacao'] ?? '') !== '') {
        $where[] = 'operacao = ?';
        $params[] = $filters['operacao'];
    }
    if (($filters['q'] ?? '') !== '') {
        $where[] = '(encomenda_numero LIKE ? OR material LIKE ? OR bloco_id LIKE ? OR COALESCE(chapa, "") LIKE ?)';
        $like = '%' . $filters['q'] . '%';
        $params[] = $like;
        $params[] = $like;
        $params[] = $like;
        $params[] = $like;
    }

    $sql = 'SELECT bloco_id, encomenda_numero, material, espessura, data_planeada, inicio, duracao_min, color, chapa, posto, operacao FROM plano';
    if ($where) {
        $sql .= ' WHERE ' . implode(' AND ', $where);
    }
    $limit = max(1, min($limit, 500));
    $sql .= " ORDER BY data_planeada DESC, inicio DESC LIMIT {$limit}";
    return db_all($sql, $params);
}

function repo_planning_anchor_date(array $filters): ?string {
    $where = [];
    $params = [];

    if (($filters['year'] ?? '') !== '') {
        $where[] = '(CAST(ano AS CHAR) = ? OR YEAR(data_planeada) = ?)';
        $params[] = $filters['year'];
        $params[] = $filters['year'];
    }
    if (($filters['posto'] ?? '') !== '') {
        $where[] = 'posto = ?';
        $params[] = $filters['posto'];
    }
    if (($filters['operacao'] ?? '') !== '') {
        $where[] = 'operacao = ?';
        $params[] = $filters['operacao'];
    }
    if (($filters['q'] ?? '') !== '') {
        $where[] = '(encomenda_numero LIKE ? OR material LIKE ? OR bloco_id LIKE ? OR COALESCE(chapa, "") LIKE ?)';
        $like = '%' . $filters['q'] . '%';
        $params[] = $like;
        $params[] = $like;
        $params[] = $like;
        $params[] = $like;
    }

    $sql = 'SELECT MAX(data_planeada) FROM plano';
    if ($where) {
        $sql .= ' WHERE ' . implode(' AND ', $where);
    }

    $value = db_value($sql, $params);
    $text = trim((string) ($value ?? ''));
    return $text === '' ? null : $text;
}

function repo_planning_options(): array {
    return [
        'postos' => db_all("SELECT DISTINCT posto AS valor FROM plano WHERE COALESCE(posto, '') <> '' ORDER BY posto ASC"),
        'operacoes' => db_all("SELECT DISTINCT operacao AS valor FROM plano WHERE COALESCE(operacao, '') <> '' ORDER BY operacao ASC"),
    ];
}

function repo_workcenter_catalog(): array {
    static $catalog = null;
    if (is_array($catalog)) {
        return $catalog;
    }

    $catalog = [];
    $runtimePath = dirname(__DIR__, 2) . DIRECTORY_SEPARATOR . 'lugest_runtime_state.json';
    if (!is_file($runtimePath)) {
        return $catalog;
    }

    $decoded = json_decode((string) file_get_contents($runtimePath), true);
    if (!is_array($decoded)) {
        return $catalog;
    }

    foreach ((array) ($decoded['workcenter_catalog'] ?? []) as $row) {
        $group = trim((string) ($row['name'] ?? ''));
        $operation = trim((string) ($row['operation'] ?? $group));
        if ($group === '' || $operation === '') {
            continue;
        }
        $machines = [];
        foreach ((array) ($row['machines'] ?? []) as $machine) {
            $machineText = trim((string) $machine);
            if ($machineText !== '') {
                $machines[] = $machineText;
            }
        }
        $catalog[] = [
            'name' => $group,
            'operation' => $operation,
            'machines' => array_values(array_unique($machines)),
        ];
    }

    return $catalog;
}

function repo_workcenter_catalog_for_operation(string $operation = ''): array {
    $targetOperation = strtolower(trim($operation));
    $rows = [];
    foreach (repo_workcenter_catalog() as $row) {
        $rowOperation = strtolower(trim((string) ($row['operation'] ?? '')));
        if ($targetOperation !== '' && $rowOperation !== $targetOperation) {
            continue;
        }
        $rows[] = $row;
    }
    return $rows;
}

function repo_workcenter_operation_options(): array {
    $values = [];

    foreach (repo_workcenter_catalog() as $row) {
        $operation = trim((string) ($row['operation'] ?? ''));
        if ($operation !== '') {
            $values[strtolower($operation)] = $operation;
        }
    }

    foreach (repo_planning_options()['operacoes'] as $row) {
        $operation = trim((string) ($row['valor'] ?? ''));
        if ($operation !== '') {
            $values[strtolower($operation)] = $operation;
        }
    }

    uasort($values, static fn(string $left, string $right): int => strcasecmp($left, $right));
    return array_values($values);
}

function repo_workcenter_resource_options(string $operation = ''): array {
    $values = [];

    foreach (repo_workcenter_catalog_for_operation($operation) as $row) {
        $group = trim((string) ($row['name'] ?? ''));
        $machines = array_values(array_filter(
            array_map(static fn($value): string => trim((string) $value), (array) ($row['machines'] ?? [])),
            static fn(string $value): bool => $value !== ''
        ));

        if ($machines !== []) {
            foreach ($machines as $machine) {
                $values[strtolower($machine)] = $machine;
            }
            continue;
        }

        if ($group !== '') {
            $values[strtolower($group)] = $group;
        }
    }

    uasort($values, static fn(string $left, string $right): int => strcasecmp($left, $right));
    return array_values($values);
}

function repo_order_workcenter(string $orderNumber): string {
    static $cache = [];

    $orderKey = trim($orderNumber);
    if ($orderKey === '') {
        return '';
    }
    if (array_key_exists($orderKey, $cache)) {
        return $cache[$orderKey];
    }

    $row = db_one(
        "SELECT posto_trabalho FROM encomendas WHERE numero = ? LIMIT 1",
        [$orderKey]
    );
    $resource = trim((string) ($row['posto_trabalho'] ?? ''));
    $cache[$orderKey] = $resource;
    return $resource;
}

function repo_order_operation_resource(string $orderNumber, string $material, string $espessura, string $operation): string {
    static $cache = [];

    $orderKey = trim($orderNumber);
    $materialKey = trim($material);
    $espessuraKey = trim($espessura);
    $operationKey = trim($operation);
    $cacheKey = strtolower($orderKey . '|' . $materialKey . '|' . $espessuraKey . '|' . $operationKey);

    if ($cacheKey === '|||') {
        return '';
    }
    if (array_key_exists($cacheKey, $cache)) {
        return $cache[$cacheKey];
    }

    if ($orderKey === '' || $materialKey === '' || $espessuraKey === '' || $operationKey === '') {
        $cache[$cacheKey] = '';
        return '';
    }

    $row = db_one(
        "SELECT maquinas_operacao_json
         FROM encomenda_espessuras
         WHERE encomenda_numero = ? AND material = ? AND espessura = ?
         LIMIT 1",
        [$orderKey, $materialKey, $espessuraKey]
    );

    $resource = '';
    $rawJson = trim((string) ($row['maquinas_operacao_json'] ?? ''));
    if ($rawJson !== '') {
        $decoded = json_decode($rawJson, true);
        if (is_array($decoded)) {
            foreach ($decoded as $opName => $resourceName) {
                if (strcasecmp(trim((string) $opName), $operationKey) === 0) {
                    $resource = trim((string) $resourceName);
                    break;
                }
            }
        }
    }

    $cache[$cacheKey] = $resource;
    return $resource;
}

function repo_planning_infer_resource(array $row): array {
    $operation = trim((string) ($row['operacao'] ?? ''));
    $stored = trim((string) ($row['posto'] ?? ''));
    $orderNumber = trim((string) ($row['encomenda_numero'] ?? ''));
    $material = trim((string) ($row['material'] ?? ''));
    $espessura = trim((string) ($row['espessura'] ?? ''));
    $operationResource = repo_order_operation_resource($orderNumber, $material, $espessura, $operation);
    $orderWorkcenter = repo_order_workcenter(trim((string) ($row['encomenda_numero'] ?? '')));
    $catalogRows = repo_workcenter_catalog_for_operation($operation);
    $groupIndex = [];
    $machineIndex = [];

    foreach ($catalogRows as $catalogRow) {
        $group = trim((string) ($catalogRow['name'] ?? ''));
        if ($group !== '') {
            $groupIndex[strtolower($group)] = $catalogRow;
        }
        foreach ((array) ($catalogRow['machines'] ?? []) as $machine) {
            $machineText = trim((string) $machine);
            if ($machineText !== '') {
                $machineIndex[strtolower($machineText)] = $catalogRow;
            }
        }
    }

    if ($stored !== '' && isset($machineIndex[strtolower($stored)])) {
        $catalogRow = $machineIndex[strtolower($stored)];
        return [
            'resource' => $stored,
            'group' => trim((string) ($catalogRow['name'] ?? $stored)),
            'defined' => true,
        ];
    }

    if ($stored !== '' && isset($groupIndex[strtolower($stored)])) {
        $catalogRow = $groupIndex[strtolower($stored)];
        $machines = array_values(array_filter(
            array_map(static fn($value): string => trim((string) $value), (array) ($catalogRow['machines'] ?? [])),
            static fn(string $value): bool => $value !== ''
        ));
        if ($operationResource !== '' && isset($machineIndex[strtolower($operationResource)])) {
            return [
                'resource' => $operationResource,
                'group' => $stored,
                'defined' => true,
            ];
        }
        if ($orderWorkcenter !== '' && isset($machineIndex[strtolower($orderWorkcenter)])) {
            return [
                'resource' => $orderWorkcenter,
                'group' => $stored,
                'defined' => true,
            ];
        }
        if (count($machines) === 1) {
            return [
                'resource' => $machines[0],
                'group' => $stored,
                'defined' => true,
            ];
        }
        if (count($machines) > 1) {
            return [
                'resource' => 'Sem recurso definido',
                'group' => $stored,
                'defined' => false,
            ];
        }
        return [
            'resource' => $stored,
            'group' => $stored,
            'defined' => true,
        ];
    }

    if ($stored !== '') {
        return [
            'resource' => $stored,
            'group' => $stored,
            'defined' => true,
        ];
    }

    if ($operationResource !== '' && isset($machineIndex[strtolower($operationResource)])) {
        return [
            'resource' => $operationResource,
            'group' => trim((string) (($machineIndex[strtolower($operationResource)]['name'] ?? ''))),
            'defined' => true,
        ];
    }

    if ($orderWorkcenter !== '' && isset($machineIndex[strtolower($orderWorkcenter)])) {
        return [
            'resource' => $orderWorkcenter,
            'group' => trim((string) (($machineIndex[strtolower($orderWorkcenter)]['name'] ?? ''))),
            'defined' => true,
        ];
    }

    return [
        'resource' => 'Sem recurso definido',
        'group' => '',
        'defined' => false,
    ];
}

function repo_planning_enrich_rows(array $rows): array {
    $items = [];
    foreach ($rows as $row) {
        $resourceInfo = repo_planning_infer_resource((array) $row);
        $row['resource'] = $resourceInfo['resource'];
        $row['resource_group'] = $resourceInfo['group'];
        $row['resource_defined'] = $resourceInfo['defined'] ? 1 : 0;
        $items[] = $row;
    }
    return $items;
}

function repo_clients(string $q, int $limit = 250): array {
    $sql = 'SELECT codigo, nome, nif, contacto, email FROM clientes';
    $params = [];
    if ($q !== '') {
        $sql .= ' WHERE codigo LIKE ? OR nome LIKE ? OR nif LIKE ?';
        $like = '%' . $q . '%';
        $params = [$like, $like, $like];
    }
    $limit = max(1, min($limit, 500));
    $sql .= " ORDER BY nome ASC LIMIT {$limit}";
    return db_all($sql, $params);
}

function repo_products(string $q, int $limit = 250): array {
    $sql = 'SELECT codigo, descricao, categoria, tipo, qty, alerta, p_compra, atualizado_em FROM produtos';
    $params = [];
    if ($q !== '') {
        $sql .= ' WHERE codigo LIKE ? OR descricao LIKE ? OR categoria LIKE ? OR tipo LIKE ?';
        $like = '%' . $q . '%';
        $params = [$like, $like, $like, $like];
    }
    $limit = max(1, min($limit, 500));
    $sql .= " ORDER BY descricao ASC LIMIT {$limit}";
    return db_all($sql, $params);
}

function repo_materials(string $q, int $limit = 250): array {
    $sql = 'SELECT id, material, espessura, formato, quantidade, reservado, localizacao, atualizado_em FROM materiais';
    $params = [];
    if ($q !== '') {
        $sql .= ' WHERE id LIKE ? OR material LIKE ? OR espessura LIKE ? OR lote_fornecedor LIKE ? OR localizacao LIKE ?';
        $like = '%' . $q . '%';
        $params = [$like, $like, $like, $like, $like];
    }
    $limit = max(1, min($limit, 500));
    $sql .= " ORDER BY atualizado_em DESC, material ASC LIMIT {$limit}";
    return db_all($sql, $params);
}

function repo_quotes(array $filters, int $limit = 250): array {
    $where = [];
    $params = [];

    if (($filters['year'] ?? '') !== '') {
        $where[] = '(CAST(o.ano AS CHAR) = ? OR YEAR(o.data) = ?)';
        $params[] = $filters['year'];
        $params[] = $filters['year'];
    }
    if (($filters['q'] ?? '') !== '') {
        $where[] = '(o.numero LIKE ? OR COALESCE(c.nome, "") LIKE ? OR COALESCE(o.estado, "") LIKE ?)';
        $like = '%' . $filters['q'] . '%';
        $params[] = $like;
        $params[] = $like;
        $params[] = $like;
    }

    $sql = '
        SELECT
            o.numero,
            o.data,
            COALESCE(c.nome, o.cliente_codigo, "-") AS cliente_nome,
            COALESCE(o.estado, "-") AS estado,
            COALESCE(o.subtotal, 0) AS subtotal,
            COALESCE(o.total, 0) AS total,
            COALESCE(o.numero_encomenda, "-") AS numero_encomenda
        FROM orcamentos o
        LEFT JOIN clientes c ON c.codigo = o.cliente_codigo
    ';
    if ($where) {
        $sql .= ' WHERE ' . implode(' AND ', $where);
    }
    $limit = max(1, min($limit, 500));
    $sql .= " ORDER BY o.data DESC, o.numero DESC LIMIT {$limit}";
    return db_all($sql, $params);
}
