<?php
include __DIR__ . '/../vendor/autoload.php';
$fileName = __DIR__ . '/excel/' . basename(__FILE__ , '.php') . '.xlsx';

$chars = 'abcdefgh';

$writer = new \BaAGee\Excel\XLSXWriter();
$writer->writeSheetHeader('Sheet1', array('col-string'=>'string','col-numbers'=>'integer','col-timestamps'=>'datetime'), ['auto_filter'=>true, 'widths'=>[15,15,30]] );
for($i=0; $i<1000; $i++)
{
    $writer->writeSheetRow('Sheet1', array(
        str_shuffle($chars),
        rand()%10000,
        date('Y-m-d H:i:s',time()-(rand()%31536000))
    ));
}
$writer->writeToFile($fileName);
echo '#'.floor((memory_get_peak_usage())/1024/1024)."MB"."\n";
