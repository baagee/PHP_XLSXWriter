<?php
/**
 * Desc:
 * User: baagee
 * Date: 2019/7/27
 * Time: 20:39
 */

include __DIR__ . '/../vendor/autoload.php';


class RightToLeftTest extends \PHPUnit\Framework\TestCase
{
    public function test()
    {
        $fileName = __DIR__ . '/excel/' . basename(__FILE__, '.php') . '.xlsx';

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


        $this->assertNotEmpty('ok');
    }

}

