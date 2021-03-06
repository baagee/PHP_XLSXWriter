<?php
/**
 * Desc: 生成xlsx表格
 * User: baagee
 * Date: 2020/6/11
 * Time: 下午2:32
 */

//http://www.ecma-international.org/publications/standards/Ecma-376.htm
//http://officeopenxml.com/SSstyles.php

//http://office.microsoft.com/en-us/excel-help/excel-specifications-and-limits-HP010073849.aspx
namespace BaAGee\Excel;

/**
 * Class XLSXWriter
 * @package BaAGee\Excel
 */
class XLSXWriter
{
    /**
     *
     */
    const EXCEL_2007_MAX_ROW = 1048576;
    /**
     *
     */
    const EXCEL_2007_MAX_COL = 16384;

    /**
     * @var string
     */
    protected $title;
    /**
     * @var string
     */
    protected $subject;
    /**
     * @var string
     */
    protected $author;
    /**
     * @var bool
     */
    protected $isRightToLeft;
    /**
     * @var string
     */
    protected $company;
    /**
     * @var string
     */
    protected $description;
    /**
     * @var array
     */
    protected $keywords = array();

    /**
     * @var string
     */
    protected $currentSheet;
    /**
     * @var array
     */
    protected $sheets = array();
    /**
     * @var array
     */
    protected $tempFiles = array();
    /**
     * @var string
     */
    protected $tempDir = '';
    /**
     * @var array
     */
    protected $cellStyles = array();
    /**
     * @var array
     */
    protected $numberFormats = array();

    /**
     * XLSXWriter constructor.
     */
    public function __construct()
    {
        defined('ENT_XML1') or define('ENT_XML1', 16);//for php 5.3, avoid fatal error
        date_default_timezone_get() or date_default_timezone_set('UTC');//php.ini missing tz, avoid warning
        is_writeable($this->tempFilename()) or self::log("Warning: tempdir " . sys_get_temp_dir() . " not writeable, use ->setTempDir()");
        class_exists('\ZipArchive') or self::log("Error: ZipArchive class does not exist");
        $this->addCellStyle('GENERAL', null);
    }

    /**
     * @param string $title
     */
    public function setTitle($title = '')
    {
        $this->title = $title;
    }

    /**
     * @param string $subject
     */
    public function setSubject($subject = '')
    {
        $this->subject = $subject;
    }

    /**
     * @param string $author
     */
    public function setAuthor($author = '')
    {
        $this->author = $author;
    }

    /**
     * @param string $company
     */
    public function setCompany($company = '')
    {
        $this->company = $company;
    }

    /**
     * @param string $keywords
     */
    public function setKeywords($keywords = '')
    {
        $this->keywords = $keywords;
    }

    /**
     * @param string $description
     */
    public function setDescription($description = '')
    {
        $this->description = $description;
    }

    /**
     * @param string $tempDir
     */
    public function setTempDir($tempDir = '')
    {
        $this->tempDir = $tempDir;
    }

    /**
     * @param bool $isRightToLeft
     */
    public function setRightToLeft($isRightToLeft = false)
    {
        $this->isRightToLeft = $isRightToLeft;
    }

    /**
     *
     */
    public function __destruct()
    {
        if (!empty($this->tempFiles)) {
            foreach ($this->tempFiles as $tempFile) {
                @unlink($tempFile);
            }
        }
    }

    /**
     * @return bool|string
     */
    protected function tempFilename()
    {
        $tempdir = !empty($this->tempDir) ? $this->tempDir : sys_get_temp_dir();
        $filename = tempnam($tempdir, "xlsx_writer_");
        $this->tempFiles[] = $filename;
        return $filename;
    }

    /**
     * 直接输出到std out
     */
    public function writeToStdOut()
    {
        $tempFile = $this->tempFilename();
        self::writeToFile($tempFile);
        readfile($tempFile);
    }

    /**
     * 输出字符串
     * @return bool|string
     */
    public function writeToString()
    {
        $tempFile = $this->tempFilename();
        self::writeToFile($tempFile);
        return file_get_contents($tempFile);
    }

    /**
     * 保存到文件
     * @param $filename
     */
    public function writeToFile($filename)
    {
        foreach ($this->sheets as $sheetName => $sheet) {
            self::finalizeSheet($sheetName);//making sure all footers have been written
        }

        if (file_exists($filename)) {
            if (is_writable($filename)) {
                @unlink($filename); //if the zip already exists, remove it
            } else {
                self::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", file is not writeable.");
                return;
            }
        }
        $zip = new \ZipArchive();
        if (empty($this->sheets)) {
            self::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", no worksheets defined.");
            return;
        }
        if (!$zip->open($filename, \ZipArchive::CREATE)) {
            self::log("Error in " . __CLASS__ . "::" . __FUNCTION__ . ", unable to create zip.");
            return;
        }

        $zip->addEmptyDir("docProps/");
        $zip->addFromString("docProps/app.xml", self::buildAppXML());
        $zip->addFromString("docProps/core.xml", self::buildCoreXML());

        $zip->addEmptyDir("_rels/");
        $zip->addFromString("_rels/.rels", self::buildRelationshipsXML());

        $zip->addEmptyDir("xl/worksheets/");
        foreach ($this->sheets as $sheet) {
            $zip->addFile($sheet->fileName, "xl/worksheets/" . $sheet->xmlName);
        }
        $zip->addFromString("xl/workbook.xml", self::buildWorkbookXML());
        $zip->addFile($this->writeStylesXML(), "xl/styles.xml");  //$zip->addFromString("xl/styles.xml"           , self::buildStylesXML() );
        $zip->addFromString("[Content_Types].xml", self::buildContentTypesXML());

        $zip->addEmptyDir("xl/_rels/");
        $zip->addFromString("xl/_rels/workbook.xml.rels", self::buildWorkbookRelsXML());
        $zip->close();
    }

    /**
     * @param       $sheetName
     * @param array $colWidths
     * @param bool  $autoFilter
     * @param bool  $freezeRows
     * @param bool  $freezeColumns
     */
    protected function initializeSheet($sheetName, $colWidths = array(), $autoFilter = false, $freezeRows = false, $freezeColumns = false)
    {
        //if already initialized
        if ($this->currentSheet == $sheetName || isset($this->sheets[$sheetName]))
            return;

        $sheetFileName = $this->tempFilename();
        $sheetXmlName = 'sheet' . (count($this->sheets) + 1) . ".xml";
        $sheet = new ExcelSheet($sheetFileName, $sheetName, $sheetXmlName, $autoFilter, $freezeRows, $freezeColumns);
        $this->sheets[$sheetName] = $sheet;
        $rightToLeftValue = $this->isRightToLeft ? 'true' : 'false';
        $tabSelected = count($this->sheets) == 1 ? 'true' : 'false';//only first sheet is selected
        $maxCell = XLSXWriter::xlsCell(self::EXCEL_2007_MAX_ROW, self::EXCEL_2007_MAX_COL);//XFE1048577
        $sheet->fileWriter->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $sheet->fileWriter->write('<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">');
        $sheet->fileWriter->write('<sheetPr filterMode="false">');
        $sheet->fileWriter->write('<pageSetUpPr fitToPage="false"/>');
        $sheet->fileWriter->write('</sheetPr>');
        $sheet->maxCellTagStart = $sheet->fileWriter->fTell();
        $sheet->fileWriter->write('<dimension ref="A1:' . $maxCell . '"/>');
        $sheet->maxCellTagEnd = $sheet->fileWriter->fTell();
        $sheet->fileWriter->write('<sheetViews>');
        $sheet->fileWriter->write('<sheetView colorId="64" defaultGridColor="true" rightToLeft="' . $rightToLeftValue . '" showFormulas="false" showGridLines="true" showOutlineSymbols="true" showRowColHeaders="true" showZeros="true" tabSelected="' . $tabSelected . '" topLeftCell="A1" view="normal" windowProtection="false" workbookViewId="0" zoomScale="100" zoomScaleNormal="100" zoomScalePageLayoutView="100">');
        if ($sheet->freezeRows && $sheet->freezeColumns) {
            $sheet->fileWriter->write('<pane ySplit="' . $sheet->freezeRows . '" xSplit="' . $sheet->freezeColumns . '" topLeftCell="' . self::xlsCell($sheet->freezeRows, $sheet->freezeColumns) . '" activePane="bottomRight" state="frozen"/>');
            $sheet->fileWriter->write('<selection activeCell="' . self::xlsCell($sheet->freezeRows, 0) . '" activeCellId="0" pane="topRight" sqref="' . self::xlsCell($sheet->freezeRows, 0) . '"/>');
            $sheet->fileWriter->write('<selection activeCell="' . self::xlsCell(0, $sheet->freezeColumns) . '" activeCellId="0" pane="bottomLeft" sqref="' . self::xlsCell(0, $sheet->freezeColumns) . '"/>');
            $sheet->fileWriter->write('<selection activeCell="' . self::xlsCell($sheet->freezeRows, $sheet->freezeColumns) . '" activeCellId="0" pane="bottomRight" sqref="' . self::xlsCell($sheet->freezeRows, $sheet->freezeColumns) . '"/>');
        } elseif ($sheet->freezeRows) {
            $sheet->fileWriter->write('<pane ySplit="' . $sheet->freezeRows . '" topLeftCell="' . self::xlsCell($sheet->freezeRows, 0) . '" activePane="bottomLeft" state="frozen"/>');
            $sheet->fileWriter->write('<selection activeCell="' . self::xlsCell($sheet->freezeRows, 0) . '" activeCellId="0" pane="bottomLeft" sqref="' . self::xlsCell($sheet->freezeRows, 0) . '"/>');
        } elseif ($sheet->freezeColumns) {
            $sheet->fileWriter->write('<pane xSplit="' . $sheet->freezeColumns . '" topLeftCell="' . self::xlsCell(0, $sheet->freezeColumns) . '" activePane="topRight" state="frozen"/>');
            $sheet->fileWriter->write('<selection activeCell="' . self::xlsCell(0, $sheet->freezeColumns) . '" activeCellId="0" pane="topRight" sqref="' . self::xlsCell(0, $sheet->freezeColumns) . '"/>');
        } else { // not frozen
            $sheet->fileWriter->write('<selection activeCell="A1" activeCellId="0" pane="topLeft" sqref="A1"/>');
        }
        $sheet->fileWriter->write('</sheetView>');
        $sheet->fileWriter->write('</sheetViews>');
        $sheet->fileWriter->write('<cols>');
        $i = 0;
        if (!empty($colWidths)) {
            foreach ($colWidths as $columnWidth) {
                $sheet->fileWriter->write('<col collapsed="false" hidden="false" max="' . ($i + 1) . '" min="' . ($i + 1) . '" style="0" customWidth="true" width="' . floatval($columnWidth) . '"/>');
                $i++;
            }
        }
        $sheet->fileWriter->write('<col collapsed="false" hidden="false" max="1024" min="' . ($i + 1) . '" style="0" customWidth="false" width="11.5"/>');
        $sheet->fileWriter->write('</cols>');
        $sheet->fileWriter->write('<sheetData>');
    }

    /**
     * @param $numberFormat
     * @param $cellStyleString
     * @return false|int|string
     */
    private function addCellStyle($numberFormat, $cellStyleString)
    {
        $numberFormatIdx = self::addToListGetIndex($this->numberFormats, $numberFormat);
        $lookupString = $numberFormatIdx . ";" . $cellStyleString;
        return self::addToListGetIndex($this->cellStyles, $lookupString);
    }

    /**
     * @param $headerTypes
     * @return array
     */
    private function initializeColumnTypes($headerTypes)
    {
        $columnTypes = array();
        foreach ($headerTypes as $v) {
            $numberFormat = self::numberFormatStandardized($v);
            $numberFormatType = self::determineNumberFormatType($numberFormat);
            $cellStyleIdx = $this->addCellStyle($numberFormat, null);
            $columnTypes[] = array('number_format' => $numberFormat,//contains excel format like 'YYYY-MM-DD HH:MM:SS'
                                   'number_format_type' => $numberFormatType, //contains friendly format like 'datetime'
                                   'default_cell_style' => $cellStyleIdx,
            );
        }
        return $columnTypes;
    }

    /**
     * 写入表头
     * @param string $sheetName
     * @param array  $headerTypes
     * @param null   $colOptions
     */
    public function writeSheetHeader($sheetName, array $headerTypes, $colOptions = null)
    {
        if (empty($sheetName) || empty($headerTypes) || !empty($this->sheets[$sheetName]))
            return;

        $suppressRow = isset($colOptions['suppress_row']) ? intval($colOptions['suppress_row']) : false;
        if (is_bool($colOptions)) {
            self::log("Warning! passing $suppressRow=false|true to writeSheetHeader() is deprecated, this will be removed in a future version.");
            $suppressRow = intval($colOptions);
        }
        $style = &$colOptions;

        $colWidths = isset($colOptions['widths']) ? (array)$colOptions['widths'] : array();
        $autoFilter = isset($colOptions['auto_filter']) ? intval($colOptions['auto_filter']) : false;
        $freezeRows = isset($colOptions['freeze_rows']) ? intval($colOptions['freeze_rows']) : false;
        $freezeColumns = isset($colOptions['freeze_columns']) ? intval($colOptions['freeze_columns']) : false;
        self::initializeSheet($sheetName, $colWidths, $autoFilter, $freezeRows, $freezeColumns);
        $sheet = $this->sheets[$sheetName];
        $sheet->columns = $this->initializeColumnTypes($headerTypes);
        if (!$suppressRow) {
            $headerRow = array_keys($headerTypes);

            $sheet->fileWriter->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . (1) . '">');
            foreach ($headerRow as $c => $v) {
                $cellStyleIdx = empty($style) ? $sheet->columns[$c]['default_cell_style'] : $this->addCellStyle('GENERAL', json_encode(isset($style[0]) ? $style[$c] : $style));
                $this->writeCell($sheet->fileWriter, 0, $c, $v, $numberFormatType = 'n_string', $cellStyleIdx);
            }
            $sheet->fileWriter->write('</row>');
            $sheet->rowCount++;
        }
        $this->currentSheet = $sheetName;
    }

    /**
     * 写入一行
     * @param string $sheetName
     * @param array  $row
     * @param null   $rowOptions
     */
    public function writeSheetRow($sheetName, array $row, $rowOptions = null)
    {
        if (empty($sheetName))
            return;

        $this->initializeSheet($sheetName);
        $sheet = $this->sheets[$sheetName];
        if (count($sheet->columns) < count($row)) {
            $defaultColumnTypes = $this->initializeColumnTypes(array_fill($from = 0, $until = count($row), 'GENERAL'));//will map to n_auto
            $sheet->columns = array_merge((array)$sheet->columns, $defaultColumnTypes);
        }

        if (!empty($rowOptions)) {
            $ht = isset($rowOptions['height']) ? floatval($rowOptions['height']) : 12.1;
            $customHt = isset($rowOptions['height']);
            $hidden = isset($rowOptions['hidden']) ? (bool)($rowOptions['hidden']) : false;
            $collapsed = isset($rowOptions['collapsed']) ? (bool)($rowOptions['collapsed']) : false;
            $sheet->fileWriter->write('<row collapsed="' . ($collapsed) . '" customFormat="false" customHeight="' . ($customHt) . '" hidden="' . ($hidden) . '" ht="' . ($ht) . '" outlineLevel="0" r="' . ($sheet->rowCount + 1) . '">');
        } else {
            $sheet->fileWriter->write('<row collapsed="false" customFormat="false" customHeight="false" hidden="false" ht="12.1" outlineLevel="0" r="' . ($sheet->rowCount + 1) . '">');
        }

        $style = &$rowOptions;
        $c = 0;
        foreach ($row as $v) {
            $numberFormat = $sheet->columns[$c]['number_format'];
            $numberFormatType = $sheet->columns[$c]['number_format_type'];
            $cellStyleIdx = empty($style) ? $sheet->columns[$c]['default_cell_style'] : $this->addCellStyle($numberFormat, json_encode(isset($style[0]) ? $style[$c] : $style));
            $this->writeCell($sheet->fileWriter, $sheet->rowCount, $c, $v, $numberFormatType, $cellStyleIdx);
            $c++;
        }
        $sheet->fileWriter->write('</row>');
        $sheet->rowCount++;
        $this->currentSheet = $sheetName;
    }

    /**
     * 获取行数
     * @param string $sheetName
     * @return int
     */
    public function countSheetRows($sheetName = '')
    {
        $sheetName = $sheetName ? $sheetName : $this->currentSheet;
        return array_key_exists($sheetName, $this->sheets) ? $this->sheets[$sheetName]->rowCount : 0;
    }

    /**
     * @param $sheetName
     */
    protected function finalizeSheet($sheetName)
    {
        if (empty($sheetName) || $this->sheets[$sheetName]->finalized)
            return;

        /**
         * @var $sheet ExcelSheet
         */
        $sheet = $this->sheets[$sheetName];

        $sheet->fileWriter->write('</sheetData>');

        if (!empty($sheet->mergeCells)) {
            $sheet->fileWriter->write('<mergeCells>');
            foreach ($sheet->mergeCells as $range) {
                $sheet->fileWriter->write('<mergeCell ref="' . $range . '"/>');
            }
            $sheet->fileWriter->write('</mergeCells>');
        }

        $maxCell = self::xlsCell($sheet->rowCount - 1, count($sheet->columns) - 1);

        if ($sheet->autoFilter) {
            $sheet->fileWriter->write('<autoFilter ref="A1:' . $maxCell . '"/>');
        }

        $sheet->fileWriter->write('<printOptions headings="false" gridLines="false" gridLinesSet="true" horizontalCentered="false" verticalCentered="false"/>');
        $sheet->fileWriter->write('<pageMargins left="0.5" right="0.5" top="1.0" bottom="1.0" header="0.5" footer="0.5"/>');
        $sheet->fileWriter->write('<pageSetup blackAndWhite="false" cellComments="none" copies="1" draft="false" firstPageNumber="1" fitToHeight="1" fitToWidth="1" horizontalDpi="300" orientation="portrait" pageOrder="downThenOver" paperSize="1" scale="100" useFirstPageNumber="true" usePrinterDefaults="false" verticalDpi="300"/>');
        $sheet->fileWriter->write('<headerFooter differentFirst="false" differentOddEven="false">');
        $sheet->fileWriter->write('<oddHeader>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12&amp;A</oddHeader>');
        $sheet->fileWriter->write('<oddFooter>&amp;C&amp;&quot;Times New Roman,Regular&quot;&amp;12Page &amp;P</oddFooter>');
        $sheet->fileWriter->write('</headerFooter>');
        $sheet->fileWriter->write('</worksheet>');

        $maxCellTag = '<dimension ref="A1:' . $maxCell . '"/>';
        $paddingLength = $sheet->maxCellTagEnd - $sheet->maxCellTagStart - strlen($maxCellTag);
        $sheet->fileWriter->fSeek($sheet->maxCellTagStart);
        $sheet->fileWriter->write($maxCellTag . str_repeat(" ", $paddingLength));
        $sheet->fileWriter->close();
        $sheet->finalized = true;
    }

    /**
     * 合并单元格
     * @param string $sheetName
     * @param int    $startCellRow    开始行 从1开始
     * @param int    $startCellColumn 开始列 从1开始
     * @param int    $endCellRow      结束行 从1开始
     * @param int    $endCellColumn   结束列 从1开始
     */
    public function markMergedCell($sheetName, $startCellRow, $startCellColumn, $endCellRow, $endCellColumn)
    {
        if (empty($sheetName) || $this->sheets[$sheetName]->finalized)
            return;

        self::initializeSheet($sheetName);
        $sheet = $this->sheets[$sheetName];

        $startCell = self::xlsCell($startCellRow - 1, $startCellColumn - 1);
        $endCell = self::xlsCell($endCellRow - 1, $endCellColumn - 1);
        $sheet->mergeCells[] = $startCell . ":" . $endCell;
    }

    /**
     * 批量写入数据
     * @param array  $data
     * @param string $sheetName
     * @param array  $headerTypes
     */
    public function writeSheet(array $data, $sheetName = '', array $headerTypes = array())
    {
        $sheetName = empty($sheetName) ? 'Sheet1' : $sheetName;
        $data = empty($data) ? array(array('')) : $data;
        if (!empty($headerTypes)) {
            $this->writeSheetHeader($sheetName, $headerTypes);
        }
        foreach ($data as $i => $row) {
            $this->writeSheetRow($sheetName, $row);
        }
        $this->finalizeSheet($sheetName);
    }

    /**
     * @param WriterBuffer $file
     * @param              $rowNumber
     * @param              $columnNumber
     * @param              $value
     * @param              $numFormatType
     * @param              $cellStyleIdx
     */
    protected function writeCell(WriterBuffer $file, $rowNumber, $columnNumber, $value, $numFormatType, $cellStyleIdx)
    {
        $cellName = self::xlsCell($rowNumber, $columnNumber);

        if (!is_scalar($value) || $value === '') { //objects, array, empty
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '"/>');
        } elseif (is_string($value) && $value[0] == '=') {
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="s"><f>' . self::xmlSpecialChars($value) . '</f></c>');
        } elseif ($numFormatType == 'n_date') {
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . intval(self::convertDateTime($value)) . '</v></c>');
        } elseif ($numFormatType == 'n_datetime') {
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . self::convertDateTime($value) . '</v></c>');
        } elseif ($numFormatType == 'n_numeric') {
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . self::xmlSpecialChars($value) . '</v></c>');//int,float,currency
        } elseif ($numFormatType == 'n_string') {
            $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="inlineStr"><is><t>' . self::xmlSpecialChars($value) . '</t></is></c>');
        } elseif ($numFormatType == 'n_auto' || 1) { //auto-detect unknown column types
            if (!is_string($value) || $value == '0' || ($value[0] != '0' && ctype_digit($value)) || preg_match("/^-?(0|[1-9][0-9]*)(\.[0-9]+)?$/", $value)) {
                $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="n"><v>' . self::xmlSpecialChars($value) . '</v></c>');//int,float,currency
            } else { //implied: ($cellFormat=='string')
                $file->write('<c r="' . $cellName . '" s="' . $cellStyleIdx . '" t="inlineStr"><is><t>' . self::xmlSpecialChars($value) . '</t></is></c>');
            }
        }
    }

    /**
     * @return array
     */
    protected function styleFontIndexes()
    {
        static $borderAllowed = array('left', 'right', 'top', 'bottom');
        static $borderStyleAllowed = array('thin', 'medium', 'thick', 'dashDot', 'dashDotDot', 'dashed', 'dotted', 'double', 'hair', 'mediumDashDot', 'mediumDashDotDot', 'mediumDashed', 'slantDashDot');
        static $horizontalAllowed = array('general', 'left', 'right', 'justify', 'center');
        static $verticalAllowed = array('bottom', 'center', 'distributed', 'top');
        $defaultFont = array('size' => '10', 'name' => 'Arial', 'family' => '2');
        $fills = array('', '');//2 placeholders for static xml later
        $fonts = array('', '', '', '');//4 placeholders for static xml later
        $borders = array('');//1 placeholder for static xml later
        $styleIndexes = array();
        foreach ($this->cellStyles as $i => $cellStyleString) {
            $semiColonPos = strpos($cellStyleString, ";");
            $numberFormatIdx = substr($cellStyleString, 0, $semiColonPos);
            $styleJsonString = substr($cellStyleString, $semiColonPos + 1);
            $style = json_decode($styleJsonString, true);

            $styleIndexes[$i] = array('num_fmt_idx' => $numberFormatIdx);//initialize entry
            if (isset($style['border']) && is_string($style['border']))//border is a comma delimited str
            {
                $borderValue['side'] = array_intersect(explode(",", $style['border']), $borderAllowed);
                if (isset($style['border-style']) && in_array($style['border-style'], $borderStyleAllowed)) {
                    $borderValue['style'] = $style['border-style'];
                }
                if (isset($style['border-color']) && is_string($style['border-color']) && $style['border-color'][0] == '#') {
                    $v = substr($style['border-color'], 1, 6);
                    $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v;// expand cf0 => ccff00
                    $borderValue['color'] = "FF" . strtoupper($v);
                }
                $styleIndexes[$i]['border_idx'] = self::addToListGetIndex($borders, json_encode($borderValue));
            }
            if (isset($style['fill']) && is_string($style['fill']) && $style['fill'][0] == '#') {
                $v = substr($style['fill'], 1, 6);
                $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v;// expand cf0 => ccff00
                $styleIndexes[$i]['fill_idx'] = self::addToListGetIndex($fills, "FF" . strtoupper($v));
            }
            if (isset($style['halign']) && in_array($style['halign'], $horizontalAllowed)) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['halign'] = $style['halign'];
            }
            if (isset($style['valign']) && in_array($style['valign'], $verticalAllowed)) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['valign'] = $style['valign'];
            }
            if (isset($style['wrap_text'])) {
                $styleIndexes[$i]['alignment'] = true;
                $styleIndexes[$i]['wrap_text'] = (bool)$style['wrap_text'];
            }

            $font = $defaultFont;
            if (isset($style['font-size'])) {
                $font['size'] = floatval($style['font-size']);//floatval to allow "10.5" etc
            }
            if (isset($style['font']) && is_string($style['font'])) {
                if ($style['font'] == 'Comic Sans MS') {
                    $font['family'] = 4;
                }
                if ($style['font'] == 'Times New Roman') {
                    $font['family'] = 1;
                }
                if ($style['font'] == 'Courier New') {
                    $font['family'] = 3;
                }
                $font['name'] = strval($style['font']);
            }
            if (isset($style['font-style']) && is_string($style['font-style'])) {
                if (strpos($style['font-style'], 'bold') !== false) {
                    $font['bold'] = true;
                }
                if (strpos($style['font-style'], 'italic') !== false) {
                    $font['italic'] = true;
                }
                if (strpos($style['font-style'], 'strike') !== false) {
                    $font['strike'] = true;
                }
                if (strpos($style['font-style'], 'underline') !== false) {
                    $font['underline'] = true;
                }
            }
            if (isset($style['color']) && is_string($style['color']) && $style['color'][0] == '#') {
                $v = substr($style['color'], 1, 6);
                $v = strlen($v) == 3 ? $v[0] . $v[0] . $v[1] . $v[1] . $v[2] . $v[2] : $v;// expand cf0 => ccff00
                $font['color'] = "FF" . strtoupper($v);
            }
            if ($font != $defaultFont) {
                $styleIndexes[$i]['font_idx'] = self::addToListGetIndex($fonts, json_encode($font));
            }
        }
        return array('fills' => $fills, 'fonts' => $fonts, 'borders' => $borders, 'styles' => $styleIndexes);
    }

    /**
     * @return bool|string
     */
    protected function writeStylesXML()
    {
        $r = self::styleFontIndexes();
        $fills = $r['fills'];
        $fonts = $r['fonts'];
        $borders = $r['borders'];
        $styleIndexes = $r['styles'];

        $temporaryFilename = $this->tempFilename();
        $file = new WriterBuffer($temporaryFilename);
        $file->write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n");
        $file->write('<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">');
        $file->write('<numFmts count="' . count($this->numberFormats) . '">');
        foreach ($this->numberFormats as $i => $v) {
            $file->write('<numFmt numFmtId="' . (164 + $i) . '" formatCode="' . self::xmlSpecialChars($v) . '" />');
        }
        //$file->write(		'<numFmt formatCode="GENERAL" numFmtId="164"/>');
        //$file->write(		'<numFmt formatCode="[$$-1009]#,##0.00;[RED]\-[$$-1009]#,##0.00" numFmtId="165"/>');
        //$file->write(		'<numFmt formatCode="YYYY-MM-DD\ HH:MM:SS" numFmtId="166"/>');
        //$file->write(		'<numFmt formatCode="YYYY-MM-DD" numFmtId="167"/>');
        $file->write('</numFmts>');

        $file->write('<fonts count="' . (count($fonts)) . '">');
        $file->write('<font><name val="Arial"/><charset val="1"/><family val="2"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');
        $file->write('<font><name val="Arial"/><family val="0"/><sz val="10"/></font>');

        foreach ($fonts as $font) {
            if (!empty($font)) { //fonts have 4 empty placeholders in array to offset the 4 static xml entries above
                $f = json_decode($font, true);
                $file->write('<font>');
                $file->write('<name val="' . htmlspecialchars($f['name']) . '"/><charset val="1"/><family val="' . intval($f['family']) . '"/>');
                $file->write('<sz val="' . intval($f['size']) . '"/>');
                if (!empty($f['color'])) {
                    $file->write('<color rgb="' . strval($f['color']) . '"/>');
                }
                if (!empty($f['bold'])) {
                    $file->write('<b val="true"/>');
                }
                if (!empty($f['italic'])) {
                    $file->write('<i val="true"/>');
                }
                if (!empty($f['underline'])) {
                    $file->write('<u val="single"/>');
                }
                if (!empty($f['strike'])) {
                    $file->write('<strike val="true"/>');
                }
                $file->write('</font>');
            }
        }
        $file->write('</fonts>');

        $file->write('<fills count="' . (count($fills)) . '">');
        $file->write('<fill><patternFill patternType="none"/></fill>');
        $file->write('<fill><patternFill patternType="gray125"/></fill>');
        foreach ($fills as $fill) {
            if (!empty($fill)) { //fills have 2 empty placeholders in array to offset the 2 static xml entries above
                $file->write('<fill><patternFill patternType="solid"><fgColor rgb="' . strval($fill) . '"/><bgColor indexed="64"/></patternFill></fill>');
            }
        }
        $file->write('</fills>');

        $file->write('<borders count="' . (count($borders)) . '">');
        $file->write('<border diagonalDown="false" diagonalUp="false"><left/><right/><top/><bottom/><diagonal/></border>');
        foreach ($borders as $border) {
            if (!empty($border)) { //fonts have an empty placeholder in the array to offset the static xml entry above
                $pieces = json_decode($border, true);
                $borderStyle = !empty($pieces['style']) ? $pieces['style'] : 'hair';
                $borderColor = !empty($pieces['color']) ? '<color rgb="' . strval($pieces['color']) . '"/>' : '';
                $file->write('<border diagonalDown="false" diagonalUp="false">');
                foreach (array('left', 'right', 'top', 'bottom') as $side) {
                    $showSide = in_array($side, $pieces['side']) ? true : false;
                    $file->write($showSide ? "<$side style=\"$borderStyle\">$borderColor</$side>" : "<$side/>");
                }
                $file->write('<diagonal/>');
                $file->write('</border>');
            }
        }
        $file->write('</borders>');

        $file->write('<cellStyleXfs count="20">');
        $file->write('<xf applyAlignment="true" applyBorder="true" applyFont="true" applyProtection="true" borderId="0" fillId="0" fontId="0" numFmtId="164">');
        $file->write('<alignment horizontal="general" indent="0" shrinkToFit="false" textRotation="0" vertical="bottom" wrapText="false"/>');
        $file->write('<protection hidden="false" locked="true"/>');
        $file->write('</xf>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="2" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="0"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="43"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="41"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="44"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="42"/>');
        $file->write('<xf applyAlignment="false" applyBorder="false" applyFont="true" applyProtection="false" borderId="0" fillId="0" fontId="1" numFmtId="9"/>');
        $file->write('</cellStyleXfs>');

        $file->write('<cellXfs count="' . (count($styleIndexes)) . '">');
        //$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="164" xfId="0"/>');
        //$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="165" xfId="0"/>');
        //$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="166" xfId="0"/>');
        //$file->write(		'<xf applyAlignment="false" applyBorder="false" applyFont="false" applyProtection="false" borderId="0" fillId="0" fontId="0" numFmtId="167" xfId="0"/>');
        foreach ($styleIndexes as $v) {
            $applyAlignment = isset($v['alignment']) ? 'true' : 'false';
            $wrapText = !empty($v['wrap_text']) ? 'true' : 'false';
            $horizAlignment = isset($v['halign']) ? $v['halign'] : 'general';
            $vertAlignment = isset($v['valign']) ? $v['valign'] : 'bottom';
            $applyBorder = isset($v['border_idx']) ? 'true' : 'false';
            $applyFont = 'true';
            $borderIdx = isset($v['border_idx']) ? intval($v['border_idx']) : 0;
            $fillIdx = isset($v['fill_idx']) ? intval($v['fill_idx']) : 0;
            $fontIdx = isset($v['font_idx']) ? intval($v['font_idx']) : 0;
            //$file->write('<xf applyAlignment="'.$applyAlignment.'" applyBorder="'.$applyBorder.'" applyFont="'.$applyFont.'" applyProtection="false" borderId="'.($borderIdx).'" fillId="'.($fillIdx).'" fontId="'.($fontIdx).'" numFmtId="'.(164+$v['num_fmt_idx']).'" xfId="0"/>');
            $file->write('<xf applyAlignment="' . $applyAlignment . '" applyBorder="' . $applyBorder . '" applyFont="' . $applyFont . '" applyProtection="false" borderId="' . ($borderIdx) . '" fillId="' . ($fillIdx) . '" fontId="' . ($fontIdx) . '" numFmtId="' . (164 + $v['num_fmt_idx']) . '" xfId="0">');
            $file->write('	<alignment horizontal="' . $horizAlignment . '" vertical="' . $vertAlignment . '" textRotation="0" wrapText="' . $wrapText . '" indent="0" shrinkToFit="false"/>');
            $file->write('	<protection locked="true" hidden="false"/>');
            $file->write('</xf>');
        }
        $file->write('</cellXfs>');
        $file->write('<cellStyles count="6">');
        $file->write('<cellStyle builtinId="0" customBuiltin="false" name="Normal" xfId="0"/>');
        $file->write('<cellStyle builtinId="3" customBuiltin="false" name="Comma" xfId="15"/>');
        $file->write('<cellStyle builtinId="6" customBuiltin="false" name="Comma [0]" xfId="16"/>');
        $file->write('<cellStyle builtinId="4" customBuiltin="false" name="Currency" xfId="17"/>');
        $file->write('<cellStyle builtinId="7" customBuiltin="false" name="Currency [0]" xfId="18"/>');
        $file->write('<cellStyle builtinId="5" customBuiltin="false" name="Percent" xfId="19"/>');
        $file->write('</cellStyles>');
        $file->write('</styleSheet>');
        $file->close();
        return $temporaryFilename;
    }

    /**
     * @return string
     */
    protected function buildAppXML()
    {
        $appXml = "";
        $appXml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $appXml .= '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">';
        $appXml .= '<TotalTime>0</TotalTime>';
        $appXml .= '<Company>' . self::xmlSpecialChars($this->company) . '</Company>';
        $appXml .= '</Properties>';
        return $appXml;
    }

    /**
     * @return string
     */
    protected function buildCoreXML()
    {
        $coreXml = "";
        $coreXml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $coreXml .= '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">';
        $coreXml .= '<dcterms:created xsi:type="dcterms:W3CDTF">' . date("Y-m-d\TH:i:s.00\Z") . '</dcterms:created>';//$dateTime = '2014-10-25T15:54:37.00Z';
        $coreXml .= '<dc:title>' . self::xmlSpecialChars($this->title) . '</dc:title>';
        $coreXml .= '<dc:subject>' . self::xmlSpecialChars($this->subject) . '</dc:subject>';
        $coreXml .= '<dc:creator>' . self::xmlSpecialChars($this->author) . '</dc:creator>';
        if (!empty($this->keywords)) {
            $coreXml .= '<cp:keywords>' . self::xmlSpecialChars(implode(", ", (array)$this->keywords)) . '</cp:keywords>';
        }
        $coreXml .= '<dc:description>' . self::xmlSpecialChars($this->description) . '</dc:description>';
        $coreXml .= '<cp:revision>0</cp:revision>';
        $coreXml .= '</cp:coreProperties>';
        return $coreXml;
    }

    /**
     * @return string
     */
    protected function buildRelationshipsXML()
    {
        $relsXml = "";
        $relsXml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $relsXml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $relsXml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>';
        $relsXml .= '<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>';
        $relsXml .= '<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>';
        $relsXml .= "\n";
        $relsXml .= '</Relationships>';
        return $relsXml;
    }

    /**
     * @return string
     */
    protected function buildWorkbookXML()
    {
        $i = 0;
        $workbookXml = "";
        $workbookXml .= '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' . "\n";
        $workbookXml .= '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">';
        $workbookXml .= '<fileVersion appName="Calc"/><workbookPr backupFile="false" showObjects="all" date1904="false"/><workbookProtection/>';
        $workbookXml .= '<bookViews><workbookView activeTab="0" firstSheet="0" showHorizontalScroll="true" showSheetTabs="true" showVerticalScroll="true" tabRatio="212" windowHeight="8192" windowWidth="16384" xWindow="0" yWindow="0"/></bookViews>';
        $workbookXml .= '<sheets>';
        foreach ($this->sheets as $sheetName => $sheet) {
            $sheetName = self::sanitizeSheetName($sheet->sheetName);
            $workbookXml .= '<sheet name="' . self::xmlSpecialChars($sheetName) . '" sheetId="' . ($i + 1) . '" state="visible" r:id="rId' . ($i + 2) . '"/>';
            $i++;
        }
        $workbookXml .= '</sheets>';
        $workbookXml .= '<definedNames>';
        foreach ($this->sheets as $sheetName => $sheet) {
            if ($sheet->autoFilter) {
                $sheetName = self::sanitizeSheetName($sheet->sheetName);
                $workbookXml .= '<definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\'' . self::xmlSpecialChars($sheetName) . '\'!$A$1:' . self::xlsCell($sheet->rowCount - 1, count($sheet->columns) - 1, true) . '</definedName>';
                $i++;
            }
        }
        $workbookXml .= '</definedNames>';
        $workbookXml .= '<calcPr iterateCount="100" refMode="A1" iterate="false" iterateDelta="0.001"/></workbook>';
        return $workbookXml;
    }

    /**
     * @return string
     */
    protected function buildWorkbookRelsXML()
    {
        $i = 0;
        $wkbkrelsXml = "";
        $wkbkrelsXml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $wkbkrelsXml .= '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">';
        $wkbkrelsXml .= '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>';
        foreach ($this->sheets as $sheetName => $sheet) {
            $wkbkrelsXml .= '<Relationship Id="rId' . ($i + 2) . '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/' . ($sheet->xmlName) . '"/>';
            $i++;
        }
        $wkbkrelsXml .= "\n";
        $wkbkrelsXml .= '</Relationships>';
        return $wkbkrelsXml;
    }

    /**
     * @return string
     */
    protected function buildContentTypesXML()
    {
        $contentTypesXml = "";
        $contentTypesXml .= '<?xml version="1.0" encoding="UTF-8"?>' . "\n";
        $contentTypesXml .= '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">';
        $contentTypesXml .= '<Override PartName="/_rels/.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        $contentTypesXml .= '<Override PartName="/xl/_rels/workbook.xml.rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>';
        foreach ($this->sheets as $sheetName => $sheet) {
            $contentTypesXml .= '<Override PartName="/xl/worksheets/' . ($sheet->xmlName) . '" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
        }
        $contentTypesXml .= '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>';
        $contentTypesXml .= '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>';
        $contentTypesXml .= '<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>';
        $contentTypesXml .= '<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>';
        $contentTypesXml .= "\n";
        $contentTypesXml .= '</Types>';
        return $contentTypesXml;
    }

    /**
     * @param $rowNumber     int, zero based
     * @param $columnNumber  int, zero based
     * @param $absolute      bool
     * @return string Cell label/coordinates, ex: A1, C3, AA42 (or if $absolute==true: $A$1, $C$3, $AA$42)
     */
    public static function xlsCell($rowNumber, $columnNumber, $absolute = false)
    {
        $n = $columnNumber;
        for ($r = ""; $n >= 0; $n = intval($n / 26) - 1) {
            $r = chr($n % 26 + 0x41) . $r;
        }
        if ($absolute) {
            return '$' . $r . '$' . ($rowNumber + 1);
        }
        return $r . ($rowNumber + 1);
    }

    /**
     * @param $string
     */
    public static function log($string)
    {
        //file_put_contents("php://stderr", date("Y-m-d H:i:s:").rtrim(is_array($string) ? json_encode($string) : $string)."\n");
        error_log(date("Y-m-d H:i:s:") . rtrim(is_array($string) ? json_encode($string) : $string) . "\n");
    }

    /**
     * @param $filename
     * @return mixed
     */
    public static function sanitizeFilename($filename) //http://msdn.microsoft.com/en-us/library/aa365247%28VS.85%29.aspx
    {
        $nonprinting = array_map('chr', range(0, 31));
        $invalidChars = array('<', '>', '?', '"', ':', '|', '\\', '/', '*', '&');
        $allInvalids = array_merge($nonprinting, $invalidChars);
        return str_replace($allInvalids, "", $filename);
    }

    /**
     * @param $sheetName
     * @return string
     */
    public static function sanitizeSheetName($sheetName)
    {
        static $badChars = '\\/?*:[]';
        static $goodChars = '        ';
        $sheetName = strtr($sheetName, $badChars, $goodChars);
        $sheetName = function_exists('mb_substr') ? mb_substr($sheetName, 0, 31) : substr($sheetName, 0, 31);
        $sheetName = trim(trim(trim($sheetName), "'"));//trim before and after trimming single quotes
        return !empty($sheetName) ? $sheetName : 'Sheet' . ((rand() % 900) + 100);
    }

    /**
     * @param $val
     * @return string
     */
    public static function xmlSpecialChars($val)
    {
        //note, badchars does not include \t\n\r (\x09\x0a\x0d)
        static $badChars = "\x00\x01\x02\x03\x04\x05\x06\x07\x08\x0b\x0c\x0e\x0f\x10\x11\x12\x13\x14\x15\x16\x17\x18\x19\x1a\x1b\x1c\x1d\x1e\x1f\x7f";
        static $goodChars = "                              ";
        return strtr(htmlspecialchars($val, ENT_QUOTES | ENT_XML1), $badChars, $goodChars);//strtr appears to be faster than str_replace
    }

    /**
     * @param $numFormat
     * @return string
     */
    private static function determineNumberFormatType($numFormat)
    {
        $numFormat = preg_replace("/\[(Black|Blue|Cyan|Green|Magenta|Red|White|Yellow)\]/i", "", $numFormat);
        if ($numFormat == 'GENERAL')
            return 'n_auto';
        if ($numFormat == '@')
            return 'n_string';
        if ($numFormat == '0')
            return 'n_numeric';
        if (preg_match('/[H]{1,2}:[M]{1,2}(?![^"]*+")/i', $numFormat))
            return 'n_datetime';
        if (preg_match('/[M]{1,2}:[S]{1,2}(?![^"]*+")/i', $numFormat))
            return 'n_datetime';
        if (preg_match('/[Y]{2,4}(?![^"]*+")/i', $numFormat))
            return 'n_date';
        if (preg_match('/[D]{1,2}(?![^"]*+")/i', $numFormat))
            return 'n_date';
        if (preg_match('/[M]{1,2}(?![^"]*+")/i', $numFormat))
            return 'n_date';
        if (preg_match('/$(?![^"]*+")/', $numFormat))
            return 'n_numeric';
        if (preg_match('/%(?![^"]*+")/', $numFormat))
            return 'n_numeric';
        if (preg_match('/0(?![^"]*+")/', $numFormat))
            return 'n_numeric';
        return 'n_auto';
    }

    /**
     * @param $numFormat
     * @return string
     */
    private static function numberFormatStandardized($numFormat)
    {
        if ($numFormat == 'money') {
            $numFormat = 'dollar';
        }
        if ($numFormat == 'number') {
            $numFormat = 'integer';
        }

        if ($numFormat == 'string')
            $numFormat = '@';
        else if ($numFormat == 'integer')
            $numFormat = '0';
        else if ($numFormat == 'date')
            $numFormat = 'YYYY-MM-DD';
        else if ($numFormat == 'datetime')
            $numFormat = 'YYYY-MM-DD HH:MM:SS';
        else if ($numFormat == 'time')
            $numFormat = 'HH:MM:SS';
        else if ($numFormat == 'price')
            $numFormat = '#,##0.00';
        else if ($numFormat == 'dollar')
            $numFormat = '[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00';
        else if ($numFormat == 'euro')
            $numFormat = '#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]';
        $ignoreUntil = '';
        $escaped = '';
        for ($i = 0, $ix = strlen($numFormat); $i < $ix; $i++) {
            $c = $numFormat[$i];
            if ($ignoreUntil == '' && $c == '[')
                $ignoreUntil = ']';
            else if ($ignoreUntil == '' && $c == '"')
                $ignoreUntil = '"';
            else if ($ignoreUntil == $c)
                $ignoreUntil = '';
            if ($ignoreUntil == '' && ($c == ' ' || $c == '-' || $c == '(' || $c == ')') && ($i == 0 || $numFormat[$i - 1] != '_'))
                $escaped .= "\\" . $c;
            else
                $escaped .= $c;
        }
        return $escaped;
    }

    /**
     * @param $haystack
     * @param $needle
     * @return false|int|string
     */
    public static function addToListGetIndex(&$haystack, $needle)
    {
        $existingIdx = array_search($needle, $haystack, $strict = true);
        if ($existingIdx === false) {
            $existingIdx = count($haystack);
            $haystack[] = $needle;
        }
        return $existingIdx;
    }

    /**
     * @param $dateInput
     * @return float|int|mixed
     */
    public static function convertDateTime($dateInput) //thanks to Excel::Writer::XLSX::Worksheet.pm (perl)
    {
        $seconds = 0;    # Time expressed as fraction of 24h hours in seconds
        $year = $month = $day = 0;
        $hour = $min = $sec = 0;

        $dateTime = $dateInput;
        if (preg_match("/(\d{4})-(\d{2})-(\d{2})/", $dateTime, $matches)) {
            list($junk, $year, $month, $day) = $matches;
        }
        if (preg_match("/(\d+):(\d{2}):(\d{2})/", $dateTime, $matches)) {
            list($junk, $hour, $min, $sec) = $matches;
            $seconds = ($hour * 60 * 60 + $min * 60 + $sec) / (24 * 60 * 60);
        }

        //using 1900 as epoch, not 1904, ignoring 1904 special case

        # Special cases for Excel.
        if ("$year-$month-$day" == '1899-12-31')
            return $seconds;    # Excel 1900 epoch
        if ("$year-$month-$day" == '1900-01-00')
            return $seconds;    # Excel 1900 epoch
        if ("$year-$month-$day" == '1900-02-29')
            return 60 + $seconds;    # Excel false leapday

        # We calculate the date by calculating the number of days since the epoch
        # and adjust for the number of leap days. We calculate the number of leap
        # days by normalising the year in relation to the epoch. Thus the year 2000
        # becomes 100 for 4 and 100 year leapdays and 400 for 400 year leapdays.
        $epoch = 1900;
        $offset = 0;
        $norm = 300;
        $range = $year - $epoch;

        # Set month days and check for leap year.
        $leap = (($year % 400 == 0) || (($year % 4 == 0) && ($year % 100))) ? 1 : 0;
        $mdays = array(31, ($leap ? 29 : 28), 31, 30, 31, 30, 31, 31, 30, 31, 30, 31);

        # Some boundary checks
        if ($year != 0 || $month != 0 || $day != 0) {
            if ($year < $epoch || $year > 9999)
                return 0;
            if ($month < 1 || $month > 12)
                return 0;
            if ($day < 1 || $day > $mdays[$month - 1])
                return 0;
        }

        # Accumulate the number of days since the epoch.
        $days = $day;    # Add days for current month
        $days += array_sum(array_slice($mdays, 0, $month - 1));    # Add days for past months
        $days += $range * 365;                      # Add days for past years
        $days += intval(($range) / 4);             # Add leapdays
        $days -= intval(($range + $offset) / 100); # Subtract 100 year leapdays
        $days += intval(($range + $offset + $norm) / 400);  # Add 400 year leapdays
        $days -= $leap;                                      # Already counted above

        # Adjust for Excel erroneously treating 1900 as a leap year.
        if ($days > 59) {
            $days++;
        }

        return $days + $seconds;
    }
}
