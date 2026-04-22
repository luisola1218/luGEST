<?php

require_once dirname(__DIR__) . '/src/bootstrap.php';

logout_user();
redirect('login.php');
