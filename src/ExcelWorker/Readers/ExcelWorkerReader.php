<?php namespace ExcelWorker\Readers;

use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Cell;
use PHPExcel_Cell_DefaultValueBinder;
use ExcelWorker\Exception\ExcelWorkerException;
use ExcelWorker\Parsers\ExcelWorkerParser;
use ExcelWorker\Helper\Helper;

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
     * Whether has a header.
     * @var bool
     */
    public $hasHeader;

    /**
     * The sheets selected to load.
     * @var array
     */
    protected $selectedSheets = [];

    /**
     * indices of sheets selected.
     * @var array
     */
    protected $selectedSheetIndices = [];

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
    protected $columns = [];

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
     * number of row to ignore
     * @var int
     */
    protected $skip = 0;

    /**
     * Number of columns to take.
     * @var int
     */
    protected $take = -1;

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
    public function get($columns = [])
    {
        $this->_parseFile($columns);

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
    public function getRow($row)
    {
    }

    /**
     * Get the first row. Perhaps the second row if header exist.
     * @throws ExcelWorkerException     When parsed is empty.
     */
    public function getFirst()
    {
        $row = $this->hasHeader ? 2 : 1;
        $this->getRow($row);
    }

    /**
     * Set header.
     * @param string $header
     */
    public function setHeader($header)
    {
        $this->header = $header;
    }

    /**
     * Get the header if specified.
     * @return array
     */
    public function getHeader()
    {
        return $this->header;
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
     * @return ExcelWorkerReader
     */
    public function setSelectedSheets($Sheets)
    {
        $this->selectedSheets = $Sheets;
        return $this;
    }

    /**
     * Set selected sheets by index
     * @param $sheets
     * @return ExcelWorkerReader
     */
    public function setSelectedSheetIndices($sheets)
    {
        $this->selectedSheetIndices = $sheets;
        return $this;
    }

    /**
     * Get Selected sheets.
     */
    public function getSelectedSheetIndices()
    {
        return $this->selectedSheetIndices;
    }

    /**
     * Judge whether a sheet selected.
     * @param $sheet
     * @return bool
     */
    public function isSelected($sheet)
    {
        $selectedSheets = $this->getSelectedSheetIndices();
        if(empty($selectedSheets)) return true;

        return in_array($sheet, $selectedSheets);
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
     * @param bool $hasHeader
     * @return $this
     */
    public function load($file, $hasHeader)
    {
        //initialize
        $this->_init($file);

        $this->hasHeader = $hasHeader;

        if ($this->sheetSelected())
            $this->reader->setLoadSheetsOnly($this->selectedSheets);

        $this->excel = $this->reader->load($this->file);

        return $this;
    }

    /**
     * Skip $num row.
     * @param int $skip
     * @return ExcelWorkerReader
     */
    public function skip($skip = 0)
    {
        $this->skip = $skip;
        return $this;
    }

    /**
     * Get skip.
     * @return int
     */
    public function getSkip()
    {
        return $this->skip;
    }

    /**
     * Set number of columns to take.
     * @param int $take
     * @return $this
     */
    public function take($take = -1)
    {
        $this->take = $take;
        return $this;
    }

    /**
     * Get take.
     * @return int
     */
    public function getTake()
    {
        return $this->take;
    }

    /**
     * Set limit(skip and number to take)
     * @param $skip
     * @param $take
     * @return ExcelWorkerReader
     */
    public function limit($skip = 0, $take = -1)
    {
        $this->skip($skip);
        return $this->take($take);
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
        unset($this->parsed);
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
     * @param $columns
     */
    protected function _parseFile($columns)
    {
        $columns = array_merge($this->columns, $columns);

        $parse = new ExcelWorkerParser($this);
        $this->parsed = $parse->parseFile($columns);
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