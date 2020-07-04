<?php
include __DIR__ . '/../vendor/autoload.php';
$fileName = __DIR__ . '/excel/' . basename(__FILE__, '.php') . '.xlsx';

$chars = "abcdefghijklmnopqrstuvwxyz0123456789 ";
$s = '';
for ($j = 0; $j < 16192; $j++)
    $s .= $chars[rand() % 36];

$t1 = microtime(true);
$writer = new \BaAGee\Excel\XLSXWriter();
$writer->writeSheetHeader('Sheet1', array('c1' => 'string', 'c2' => 'string', 'c3' => 'string', 'c4' => 'string'));//optional
for ($i = 0; $i < 250000; $i++) {
    $s1 = substr($s, rand() % 4000, rand() % 5 + 5);
    $s2 = substr($s, rand() % 8000, rand() % 5 + 5);
    $s3 = substr($s, rand() % 12000, rand() % 5 + 5);
    $s4 = substr($s, rand() % 16000, rand() % 5 + 5);
    $writer->writeSheetRow('Sheet1', array($s1, $s2, $s3, $s4));
}
$writer->writeToFile($fileName);
$t2 = microtime(true);
echo '#' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . PHP_EOL;
echo '#' . (($t2 - $t1)) . "s" . PHP_EOL;

