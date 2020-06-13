<?php
include __DIR__ . '/../vendor/autoload.php';
$fileName = __DIR__ . '/excel/' . basename(__FILE__ , '.php') . '.xlsx';

$writer = new \BaAGee\Excel\XLSXWriter();
$writer->writeSheetHeader('Sheet1', array(300, 234, 456, 789), ['widths' => [10, 20, 30, 40]]);
$writer->writeSheetRow('Sheet1', array(300, 234, 456, 789), ['height' => 20]);
$writer->writeSheetRow('Sheet1', array(300, 234, 456, 789), ['height' => 30]);
$writer->writeSheetRow('Sheet1', array(300, 234, 456, 789), ['height' => 40]);
$writer->writeToFile($fileName);


