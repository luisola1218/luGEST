<?php

function is_password_hash_php(string $value): bool {
    $parts = explode('$', trim($value));
    return count($parts) === 4 && $parts[0] === 'pbkdf2_sha256' && ctype_digit($parts[1]);
}

function verify_password_php(string $candidate, string $stored): bool {
    $stored = trim($stored);
    if ($stored === '') {
        return false;
    }
    if (!is_password_hash_php($stored)) {
        return hash_equals($stored, $candidate);
    }

    [$algo, $iterationsTxt, $saltTxt, $digestTxt] = explode('$', $stored, 4);
    if ($algo !== 'pbkdf2_sha256') {
        return false;
    }

    $iterations = max(1, (int) $iterationsTxt);
    $salt = base64_decode(strtr($saltTxt, '-_', '+/') . str_repeat('=', (4 - strlen($saltTxt) % 4) % 4), true);
    $expected = base64_decode(strtr($digestTxt, '-_', '+/') . str_repeat('=', (4 - strlen($digestTxt) % 4) % 4), true);
    if ($salt === false || $expected === false) {
        return false;
    }

    $calc = hash_pbkdf2('sha256', $candidate, $salt, $iterations, strlen($expected), true);
    return hash_equals($expected, $calc);
}

function current_user(): ?array {
    $user = $_SESSION['lugest_web_user'] ?? null;
    return is_array($user) ? $user : null;
}

function is_logged_in(): bool {
    return current_user() !== null;
}

function require_login(): void {
    if (!is_logged_in()) {
        redirect('login.php');
    }
}

function logout_user(): void {
    unset($_SESSION['lugest_web_user']);
}

function attempt_login(string $username, string $password): array {
    if ($username === '' || $password === '') {
        return ['ok' => false, 'message' => 'Preenche utilizador e password.'];
    }

    $row = db_one('SELECT username, password, role FROM users WHERE LOWER(username) = LOWER(?) LIMIT 1', [$username]);
    if (!$row) {
        return ['ok' => false, 'message' => 'Utilizador nao encontrado.'];
    }

    if (!verify_password_php($password, (string) ($row['password'] ?? ''))) {
        return ['ok' => false, 'message' => 'Password incorreta.'];
    }

    $role = trim((string) ($row['role'] ?? ''));
    $allowedRoles = app_config('allowed_roles', []);
    if (is_array($allowedRoles) && $allowedRoles !== [] && !in_array($role, $allowedRoles, true)) {
        return ['ok' => false, 'message' => 'Sem permissao para o portal web.'];
    }

    $_SESSION['lugest_web_user'] = [
        'username' => (string) $row['username'],
        'role' => $role,
    ];

    return ['ok' => true, 'message' => 'OK'];
}
