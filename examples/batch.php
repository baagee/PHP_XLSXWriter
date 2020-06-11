<?php
/**
 * Desc:
 * User: baagee
 * Date: 2020/6/11
 * Time: 下午5:56
 */
$dir = __DIR__;
if (!is_dir($dir . '/excel')) {
    mkdir($dir . '/excel', 0755, true);
} else {
    foreach (scandir($dir . '/excel') as $item) {
        if (!in_array($item, ['.', '..'])) {
            unlink($dir . '/excel/' . $item);
        }
    }
}

$arr = scandir($dir);
foreach ($arr as $item) {
    preg_match('/^ex\d+.*?\.php$/', $item, $m);
    if (!empty($m)) {
        $res = popen(PHP_BINARY . ' ' . realpath($dir . '/' . $item), 'w');
        if (is_resource($res)) {
            pclose($res);
            echo $item . ' success' . PHP_EOL;
        } else {
            echo sprintf('popen error file:%s' . PHP_EOL, $item);
        }
    }
}
