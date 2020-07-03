<?php
/**
 * Desc:
 * User: baagee
 * Date: 2019/7/27
 * Time: 20:39
 */

include __DIR__ . '/../vendor/autoload.php';


class StylesTest extends \PHPUnit\Framework\TestCase
{
    public function test()
    {
        $fileName = __DIR__ . '/excel/' . basename(__FILE__, '.php') . '.xlsx';
        $writer = new \BaAGee\Excel\XLSXWriter();
        $styles1 = array('font' => 'Arial', 'font-size' => 10, 'font-style' => 'bold', 'fill' => '#eee', 'halign' => 'center', 'border' => 'left,right,top,bottom');
        $styles2 = array(['font-size' => 6], ['font-size' => 8], ['font-size' => 10], ['font-size' => 16]);
        $styles3 = array(['font' => 'Arial'], ['font' => 'Courier New'], ['font' => 'Times New Roman'], ['font' => 'Comic Sans MS']);
        $styles4 = array(['font-style' => 'bold'], ['font-style' => 'italic'], ['font-style' => 'underline'], ['font-style' => 'strikethrough']);
        $styles5 = array(['color' => '#f00'], ['color' => '#0f0'], ['color' => '#00f'], ['color' => '#666']);
        $styles6 = array(['fill' => '#ffc'], ['fill' => '#fcf'], ['fill' => '#ccf'], ['fill' => '#cff']);
        $styles7 = array('border' => 'left,right,top,bottom');
        $styles8 = array(['halign' => 'left'], ['halign' => 'right'], ['halign' => 'center'], ['halign' => 'none']);
        $styles9 = array(array(), ['border' => 'left,top,bottom'], ['border' => 'top,bottom'], ['border' => 'top,bottom,right']);

        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles1);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles2);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles3);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles4);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles5);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles6);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles7);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles8);
        $writer->writeSheetRow('Sheet1', $rowdata = array(300, 234, 456, 789), $styles9);
        $writer->writeToFile($fileName);
        $this->assertNotEmpty('ok');
    }

}

