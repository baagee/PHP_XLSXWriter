<?php
/**
 * Desc:
 * User: baagee
 * Date: 2019/7/27
 * Time: 20:39
 */

include __DIR__ . '/../vendor/autoload.php';


class Numbers250kRowsTest extends \PHPUnit\Framework\TestCase
{
    public function test()
    {
        $fileName = __DIR__ . '/excel/' . basename(__FILE__, '.php') . '.xlsx';

        $writer = new \BaAGee\Excel\XLSXWriter();
        $writer->writeSheetHeader('Sheet1', array('c1' => 'integer', 'c2' => 'integer', 'c3' => 'integer', 'c4' => 'integer'));//optional
        for ($i = 0; $i < 250000; $i++) {
            $writer->writeSheetRow('Sheet1', array(rand() % 10000, rand() % 10000, rand() % 10000, rand() % 10000));
        }
        $writer->writeToFile($fileName);
        echo '#' . floor((memory_get_peak_usage()) / 1024 / 1024) . "MB" . "\n";

        $this->assertNotEmpty('ok');
    }

}

