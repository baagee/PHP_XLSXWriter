<?php
include __DIR__ . '/../vendor/autoload.php';
$fileName = __DIR__ . '/excel/' . basename(__FILE__ , '.php') . '.xlsx';

$header = array(
    'c1-text' => 'string',//text
    'c2-text' => '@',//text
);
$rows = array(
    array('abcdefg', 'hijklmnop'),
);
$writer = new \BaAGee\Excel\XLSXWriter();
$writer->setRightToLeft(true);

$writer->writeSheetHeader('Sheet1', $header);
foreach ($rows as $row)
    $writer->writeSheetRow('Sheet1', $row);
$writer->writeToFile($fileName);

