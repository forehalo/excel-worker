<?php namespace ExcelWorker\Writers;

use ExcelWorker\Helper\Helper;
use PHPExcel;
use PHPExcel_IOFactory;

/**
 * Class BaseWriter.php
 * @package     forehalo/excel-worker
 * @version     1.0.0
 * @copyright   Copyright (c) 2015 forehalo <http://www.forehalo.top>
 * @author      forehalo <forehalo@gmail.com>
 * @license     http://www.gnu.org/licenses/lgpl.html   LGPL
 */
class ExcelWorkerWriter
{
    /**
     * Path to storage file.
     * @var string
     */
    protected $path;

    /**
     * PHPExcel object
     * @var PHPExcel
     */
    protected $excel;

    /**
     * Whether excel file has header(always the first line).
     * @var bool
     */
    protected $hasHeader;

    /**
     * file name
     * @var string
     */
    protected $file;

    /**
     * title
     * @var string
     */
    protected $title;

    /**
     * file extension
     * @var string
     */
    protected $ext;

    /**
     * weiter
     * @var PHPExcel_Writer_IWriter
     */
    protected $writer;

    /**
     * The format PHPExcel will use.
     * @var string
     */
    protected $format;

    /**
     * Constructor
     */
    public function __construct()
    {
        $this->helper = new Helper();
    }

    /**
     * Writer the first line if header exist.
     * @param $header
     * @return ExcelWorkerWriter
     */
    public function writeHeader($header)
    {
        return $this->writeRow($header);
    }

    /**
     * Write data in one row given.
     * @param array $data
     * @param int $rowNum
     * @return $this
     */
    public function writeRow($data = [], $rowNum = 1)
    {
        $column = 'A';

        foreach ($data as $item) {
            $cell = $column . $rowNum;
            $this->excel->setActiveSheetIndex(0)->setCellValue($cell, $item);
            $column++;
        }
        return $this;
    }

    /**
     * Write data in one column given.
     * @param array $data
     * @param mixed $column
     * @return $this
     */
    public function writeColumn($data = [], $column = 'A')
    {
        if(is_numeric($column)) {
            $column = $this->getColumn($column);
        }
        $rowNum = $this->hasHeader ? 2 : 1;
        foreach ($data as $item) {
            $cell = $column . $rowNum;
            $this->excel->setActiveSheetIndex(0)->setCellValue($cell, $item);
            $rowNum++;
        }
        return $this;
    }

    /**
     * Save file.
     * @param string $ext
     * @param bool|false $path
     */
    public function save($ext = 'xlsx', $path = false)
    {
        $this->setStoragePath($path);

        $this->ext = $ext;

        $file = $this->path . '/' . $this->file . '.' .$this->ext;

        $this->_setFormat();
        $this->_setWriter();

        $this->writer->save($file);
    }

    /**
     * Set format
     */
    protected function _setFormat()
    {
        $this->format = $this->helper->getFormatByExtension($this->ext);
    }

    /**
     * Set writer
     * If extension is csv, set some parameters.
     */
    protected function _setWriter()
    {
        $this->writer = PHPExcel_IOFactory::createWriter($this->excel, $this->format);

        if ($this->format == 'CSV') {
            $this->writer->setDelimiter(',');
            $this->writer->setEnclosure('"');
            $this->writer->setLineEnding("\r\n");
        }
    }

    /**
     * Inject PHPExcel into $this.
     * @param PHPExcel $excel
     */
    public function injectExcel($excel)
    {
        $this->excel = $excel;
    }

    /**
     * Set file name.
     * @param string $file
     */
    public function setFileName($file)
    {
        $this->file = $file;
    }

    /**
     * Set title.
     * @param string $file
     */
    public function setTitle($file)
    {
        $this->title = $file;
        $this->excel->getProperties()->setTitle($this->title);
    }

    /**
     * Set storage path.
     * @param string $path
     */
    protected function setStoragePath($path)
    {
        $path = $path ? $path : '../result';

        $this->path = rtrim($path, '/');

        if(!is_dir($this->path))
            mkdir($path, 0777);

        if(!is_writable($this->path))
            chmod($this->path, 0777);
    }

    /**
     * Get the column name uses alphabet by column number.
     * @param $column
     * @return string
     */
    public function getColumn($column)
    {
        $n = $column;
        $ret = '';
        while ($n) {
            $ret = chr(($n - 1) % 26 + 65) . $ret;
            $n = floor(($n - 1) / 26);
        }
        return $ret;
    }
}