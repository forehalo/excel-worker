<?php namespace ExcelWorker\Readers;

use ExcelWorker\Helper\Helper;
use PHPExcel;
use ExcelWorker\Parsers\ExcelWorkerParser;
use ExcelWorker\Exception\ExcelWorkerException;
use PHPExcel_Cell;
use PHPExcel_Cell_DefaultValueBinder;
use PHPExcel_IOFactory;

/**
 * Class ExcelWorkerReader.php
 * @package     forehalo/excel-worker
 * @version     1.0.0
 * @copyright   Copyright (c) 2015 forehalo <http://www.forehalo.top>
 * @author      forehalo <forehalo@gmail.com>
 * @license     http://www.gnu.org/licenses/lgpl.html   LGPL
 */
class ExcelWorkerReader
{
    /**
     * PHPExcel object
     * @var PHPExcel
     */
    public $excel;

    /**
     * Excel reader object
     * @var PHPExcel_Reader_IReader
     */
    public $reader;

    /**
     * file name
     * @var string
     */
    protected $file = '';

    /**
     * The extension of file.
     * @var string
     */
    protected $ext = '';

    /**
     * Header (Always or Default the first row if specified) of file/sheet.
     * @var array
     */
    protected $header = [];

    /**
     * The sheets selected to load.
     * @var array
     */
    protected $selectedSheets = [];

    /**
     * All parsed content.
     * @var array
     */
    protected $parsed;

    /**
     * The path information of file.
     * @var array
     */
    protected $pathInfo;

    /**
     * Format to init IReader.
     * @var string
     */
    protected $format;

    /**
     * Columns to be shown.
     * @var array
     */
    protected $column = [];

    /**
     * Title.
     * @var string
     */
    protected $title;

    /**
     * Helper object
     * @var Helper
     */
    protected $helper;

    /**
     * Constructor
     */
    public function __construct()
    {
        $this->helper = new Helper;
    }

    /**
     * Get all the content by columns given.
     * @param array
     * @return array
     */
    public function get($column = [])
    {
        $this->_parseFile($column);

        return $this->parsed;
    }

    /**
     * Get all.
     * @return array
     */
    public function all()
    {
        return $this->get();
    }

    /**
     * Get one row by the given row number and sheetNumber.
     * @param int $row number of row
     * @param int $sheetNum index of sheet
     * @return array
     * @throws ExcelWorkerException     When a number given less than 1 or greater than count of row.
     */
    public function getRow($row, $sheetNum = -1)
    {
        //TODO
    }

    /**
     * Get the first row. Perhaps the second row if header exist.
     * @throws ExcelWorkerException     When parsed is empty.
     */
    public function getFirst()
    {
        //TODO
    }

    /**
     * Get one column by given column number except header.
     * @param int $colNum number of column
     * @return array
     * @throws ExcelWorkerException     When a number given less than 1 or greater than count of column.
     */
    public function getColumn($colNum)
    {
        //TODO
    }

    /**
     * Get cell content.
     * @param $row number of row
     * @param $col number of column
     * @return string cell content
     * @throws ExcelWorkerException     When index of cell given is invalid.
     */
    public function getCell($row, $col)
    {
        //TODO
    }

    /**
     * Get the header if specified.
     * @return array
     */
    public function getHeader()
    {
        //TODO
    }

    /**
     * Judge whether the given row exist in $content.
     * @param int $row number of row
     * @return bool
     */
    public function rowExist($row)
    {
        //TODO
    }

    /**
     * Judge Whether the given column exist in $content.
     * @param int $colNum number of column
     * @return bool
     */
    public function columnExist($colNum)
    {
        //TODO
    }

    /**
     * Judge Whether the Given index of a cell exist in $content.
     * @param int $row number of row
     * @param int $col number of column
     * @return bool
     */
    public function cellExist($row, $col)
    {
        //TODO
    }

    /**
     * Can the file be read.
     */
    public function canReader()
    {
        //TODO
    }

    /**
     * Get the sheets selected.
     * @return array
     */
    public function getSelectedSheets()
    {
        return $this->selectedSheets;
    }

    /**
     * Set selected sheets.
     * @param array $Sheets
     */
    public function setSelectedSheets($Sheets)
    {
        $this->selectedSheets = $Sheets;
    }

    /**
     * Inject the PHPExcel into $this and reset.
     * @param PHPExcel $excel
     */
    public function injectExcel($excel)
    {
        $this->excel = $excel;
        $this->_reset();
    }

    /**
     * Load file
     * @param string $file
     * @return $this
     */
    public function load($file)
    {
        //initialize
        $this->_init($file);

        if ($this->sheetSelected())
            $this->reader->setLoadSheetsOnly($this->selectedSheets);

        $this->excel = $this->reader->load($this->file);

        return $this;
    }

    /**
     * Initialize
     * @param $file
     */
    protected function _init($file)
    {
        $this->_setFile($file)
            ->setExt()
            ->setTitle()
            ->_setFormat()
            ->_setReader();
    }

    /**
     * Set file name.
     * @param string $file
     * @return $this
     */
    protected function _setFile($file)
    {
        if (is_file($file)) {
            $this->file = $file;
            $this->pathInfo = pathinfo($file);
        }
        return $this;
    }

    /**
     * Set file extension.
     * @param string $ext
     * @return $this
     */
    public function setExt($ext = '')
    {
        $this->ext = $ext ? $ext : $this->getExt();
        return $this;
    }

    /**
     * Get file extension.
     * @return string
     */
    protected function getExt()
    {
        return $this->pathInfo['extension'];
    }

    /**
     * Set title
     * @param string $title
     * @return $this
     */
    public function setTitle($title = '')
    {
        $this->title = $title ? $title : $this->getTitle();
        return $this;
    }

    /**
     * Get title
     * @return string
     */
    public function getTitle()
    {
        return $this->pathInfo['filename'];
    }

    /**
     * Use PHPExcel_IOFactory to generate reader.
     */
    protected function _setReader()
    {
        $this->reader = PHPExcel_IOFactory::createReader($this->format);
    }

    /**
     * Set format that PHPExcel use.
     * @return $this
     */
    protected function _setFormat()
    {
        $this->format = $this->helper->getFormatByExtension($this->ext);
        return $this;
    }

    /**
     * Reset resource.
     */
    protected function _reset()
    {
        $this->excel->disconnectWorksheets();
        $this->resetValueBinder();
        unset($this->parse);
    }

    /**
     * Reset ValueBinder
     */
    protected function resetValueBinder()
    {
        PHPExcel_Cell::setValueBinder(new PHPExcel_Cell_DefaultValueBinder());
    }

    /**
     * Parse file
     * @param $column
     */
    protected function _parseFile($column)
    {
        $column = array_merge($this->column, $column);

        $parse = new ExcelWorkerParser($this);
        $this->parsed = $parse->parseFile($column);
    }

    /**
     * Judge whether has sheet selected.
     * @return bool
     */
    protected function sheetSelected()
    {
        return count($this->selectedSheets) > 0;
    }
}