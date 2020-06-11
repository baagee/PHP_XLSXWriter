<?php
/**
 * Desc:
 * User: baagee
 * Date: 2020/6/11
 * Time: 下午4:26
 */

namespace BaAGee\Excel;

class ExcelSheet
{
    public $fileName  = '';
    public $sheetName = '';
    public $xmlName   = '';
    public $rowCount  = 0;
    /**
     * @var WriterBuffer
     */
    public $fileWriter      = null;
    public $columns         = [];
    public $mergeCells      = [];
    public $maxCellTagStart = 0;
    public $maxCellTagEnd   = 0;
    public $autoFilter      = '';
    public $freezeRows      = '';
    public $freezeColumns   = '';
    public $finalized       = false;

    public function __construct($sheetFileName, $sheetName, $sheetXmlName, $autoFilter, $freezeRows, $freezeColumns)
    {
        $this->fileName = $sheetFileName;
        $this->sheetName = $sheetName;
        $this->xmlName = $sheetXmlName;
        $this->fileWriter = new WriterBuffer($sheetFileName);
        $this->autoFilter = $autoFilter;
        $this->freezeRows = $freezeRows;
        $this->freezeColumns = $freezeColumns;
    }
}
