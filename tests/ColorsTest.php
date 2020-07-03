<?php
/**
 * Desc:
 * User: baagee
 * Date: 2019/7/27
 * Time: 20:39
 */

include __DIR__ . '/../vendor/autoload.php';


class ColorsTest extends \PHPUnit\Framework\TestCase
{
    public function test()
    {
        $fileName = __DIR__ . '/excel/' . basename(__FILE__, '.php') . '.xlsx';

        $writer = new \BaAGee\Excel\XLSXWriter();
        $colors = array('ff', 'cc', '99', '66', '33', '00');
        foreach ($colors as $b) {
            foreach ($colors as $g) {
                $rowdata = array();
                $rowstyle = array();
                foreach ($colors as $r) {
                    $rowdata[] = "#$r$g$b";
                    $rowstyle[] = array('fill' => "#$r$g$b");
                }
                $writer->writeSheetRow('Sheet1', $rowdata, $rowstyle);
            }
        }
        $writer->writeToFile($fileName);
        $this->assertNotEmpty('ok');
    }

}

