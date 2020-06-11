<?php
include __DIR__ . '/../vendor/autoload.php';
$fileName = __DIR__ . '/excel/' . basename(__FILE__ . '.php') . '.xlsx';

$chars = 'abcdefgh';

$writer = new \BaAGee\Excel\XLSXWriter();
$writer->writeSheetHeader('Sheet1', array('c1'=>'string','c2'=>'integer','c3'=>'integer','c4'=>'integer','c5'=>'integer'), ['freeze_rows'=>1, 'freeze_columns'=>1] );
for($i=0; $i<250; $i++)
{
    $writer->writeSheetRow('Sheet1', array(
        str_shuffle($chars),
        rand()%10000,
        rand()%10000,
        rand()%10000,
        rand()%10000
    ));
}
$writer->writeToFile($fileName);
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";
