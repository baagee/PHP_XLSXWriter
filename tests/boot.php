<?php
/**
 * Desc:
 * User: baagee
 * Date: 2020/7/3
 * Time: 下午8:00
 */

$dir = __DIR__;
include $dir . '/../vendor/autoload.php';

if (!is_dir($dir . '/excel')) {
    mkdir($dir . '/excel', 0755, true);
} else {
    foreach (scandir($dir . '/excel') as $item) {
        if (!in_array($item, ['.', '..'])) {
            unlink($dir . '/excel/' . $item);
        }
    }
}