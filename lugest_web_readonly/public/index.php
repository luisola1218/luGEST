<?php

require_once dirname(__DIR__) . '/src/bootstrap.php';

if (is_logged_in()) {
    redirect('dashboard.php');
}

redirect('login.php');
