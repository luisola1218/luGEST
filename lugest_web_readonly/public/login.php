<?php

require_once dirname(__DIR__) . '/src/bootstrap.php';

if (is_logged_in()) {
    redirect('dashboard.php');
}

$error = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $username = trim((string) ($_POST['username'] ?? ''));
    $password = (string) ($_POST['password'] ?? '');
    $result = attempt_login($username, $password);
    if ($result['ok']) {
        redirect('dashboard.php');
    }
    $error = $result['message'];
}

ob_start();
?>
<section class="auth-shell">
    <aside class="auth-spotlight">
        <div class="auth-logo-shell">
            <img class="auth-logo" src="assets/desktop-brand-logo.jpg" alt="Barcelbal">
        </div>
        <p class="eyebrow">Portal premium</p>
        <h1>Consulta empresarial com uma presenca mais elegante.</h1>
        <p>Uma experiencia mais neutra, profissional e executiva para consultar encomendas, planeamento, clientes, produtos e stock com a base real do luGEST.</p>
        <div class="auth-grid">
            <div class="mini-panel">
                <strong>Visual mais premium</strong>
                <span>Paleta neutra, superfícies mais limpas e uma apresentação mais madura para uso profissional.</span>
            </div>
            <div class="mini-panel">
                <strong>Consulta focada</strong>
                <span>Leitura rápida do estado industrial sem ruído visual nem excesso de cor.</span>
            </div>
            <div class="mini-panel">
                <strong>Base segura</strong>
                <span>Autenticação do luGEST e acesso pensado exclusivamente para leitura.</span>
            </div>
        </div>
    </aside>

    <div class="auth-card">
        <div class="auth-copy">
            <p class="eyebrow">Acesso reservado</p>
            <h1><?php echo h(app_config('app_name')); ?></h1>
            <p class="muted">Entra com um utilizador valido do luGEST para abrir o portal de consulta.</p>
        </div>

        <?php if ($error !== ''): ?>
            <div class="alert alert-danger"><?php echo h($error); ?></div>
        <?php endif; ?>

        <form method="post" class="form-grid">
            <label class="field">
                <span>Utilizador</span>
                <input type="text" name="username" autocomplete="username" required>
            </label>
            <label class="field">
                <span>Password</span>
                <input type="password" name="password" autocomplete="current-password" required>
            </label>
            <button type="submit" class="btn btn-primary">
                <span class="btn-icon"><?php echo app_icon('shield'); ?></span>
                <span>Entrar no portal</span>
            </button>
        </form>
    </div>
</section>
<?php
$content = ob_get_clean();
render_guest_page('Login', $content);
