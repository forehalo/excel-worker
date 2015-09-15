<?php namespace ExcelWorker\Readers;

use PHPExcel;
use ExcelWorker\Parsers\ExcelWorkerParser;
use ExcelWorker\Exception\ExcelWorkerException;
use PHPExcel_Cell;
use PHPExcel_Cell_DefaultValueBinder;
use PHPExcel_IOFactory;
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
     *
     * @var PHPExcel
     */
    public $excel;

    /**
     * Excel reader object
     *
     * @var PHPExcel_Reader_IReader
     */
    public $reader;
    /**
     * All cells of the file except the header(first row if specified).
     *
     * @var array $content
     */
    protected $content = [];

    /**
     * file name
     *
     * @var string
     */
    protected $file = '';

    /**
     * The extension of file.
     *
     * @var string $extension
     */
    protected $extension = '';

    /**
     * Header (Always or Default the first row if specified) of file/sheet.
     *
     * @var array $header
     */
    protected $header = [];


    private $selectedSheets = [];

    private $parse;

    private $pathInfo;

    private $format;

    /**
     * Constructor
     * Load file into $content and parse the first row into $header.
     *
     * @param ExcelWorkerParser $parser
     */
    public function __construct(ExcelWorkerParser $parser)
    {
        $this->parse = $parser;
    }

    /**
     * Get all the content as a index array.
     *
     * @return array
     * @throws ExcelWorkerException     When nothing loaded.
     */
    public function get(PHPExcel $excel)
    {

    }

    /**
     * Get one row by the given row number.
     *
     * @param int $row number of row
     * @return array
     * @throws ExcelWorkerException     When a number given less than 1 or greater than count of row.
     */
    public function getRow($row, $sheetNum = -1)
    {
        if ($this->content == []){

        }
    }

    /**
     * Get the first row. Perhaps the second row if header exist.
     *
     * @throws ExcelWorkerException
     */
    public function getFirst()
    {
    }

    /**
     * Get one column by given column number except header.
     *
     * @param int $col number of column
     * @return array
     * @throws ExcelWorkerException     When a number given less than 1 or greater than count of column.
     */
    public function getColumn($col)
    {
    }

    /**
     * Get cell content.
     *
     * @param $row number of row
     * @param $col number of column
     * @return string cell content
     * @throws ExcelWorkerException     When index of cell given is invalid.
     */
    public function getCell($row, $col)
    {
    }

    /**
     * Get the header if specified.
     *
     * @return array
     */
    public function getHeader()
    {
    }

    /**
     * Judge whether the given row exist in $content.
     *
     * @param int $row number of row
     * @return bool
     */
    public function rowExist($row)
    {
    }

    /**
     * Judge Whether the given column exist in $content.
     *
     * @param int $col number of column
     * @return bool
     */
    public function columnExist($col)
    {
    }

    /**
     * Judge Whether the Given index of a cell exist in $content.
     *
     * @param int $row number of row
     * @param int $col number of column
     * @return bool
     */
    public function cellExist($row, $col)
    {
    }

    public function canReader()
    {
    }

    /**
     * @return array
     */
    public function getSelectedSheets()
    {
        return $this->selectedSheets;
    }

    /**
     * @param int $Sheets
*/
    public function setSelectedSheets($Sheets)
    {
        $this->selectedSheets = $Sheets;
    }

    public function injectExcel($excel)
    {
        $this->excel = $excel;
        $this->_reset();
    }

    public function load($file)
    {
        $this->_init($file);

        if ($this->sheetSeleted())
            $this->reader->setLoadSheetsOnly($this->selectedSheets);
        $this->excel = $this->reader->load($this->file);

        return $this;
    }

    protected function _init($file)
    {
        $this->_setFile($file)
             ->setExtension()
             ->setTitle()
             ->_setFormat()
             ->_setReader();
    }

    protected function _setFile($file)
    {
        if (realpath($file) && is_file($file)){
            $this->file = $file;
            $this->pathInfo = pathinfo($file);
        }
        return $this;
    }

    public function setExtension($ext = '')
    {
        $this->extension = $ext ? $ext : $this->getExt();
        return $this;
    }

    private function getExt()
    {
        return $this->pathInfo['extension'];
    }

    public function setTitle($title = '')
    {
        $this->$title = $title ? $title : $this->getTitle();
        return $this;
    }

    /**
     * @return string
     */
    public function getTitle()
    {
        return $this->pathInfo['basename'];
    }

    protected function _setReader()
    {
        $this->reader = PHPExcel_IOFactory::createReader($this->format);
    }

    protected function _setFormat()
    {
        $this->format = Helper::getFormatByExtension($this->extension);
        return $this;
    }

    protected function _reset()
    {
        $this->excel->disconnectWorksheets();
        $this->resetValueBinder();
        unset($this->parse);
    }

    private function resetValueBinder()
    {
        PHPExcel_Cell::setValueBinder(new PHPExcel_Cell_DefaultValueBinder());
    }

    protected function sheetSeleted()
    {
        return count($this->selectedSheets) > 0;
    }


}