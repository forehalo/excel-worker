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
    protected $path;

    protected $excel;

    protected $hasHeader;

    private $file;

    private $title;

    private $ext;

    protected $writer;

    private $format;

    public function __construct()
    {
        $this->helper = new Helper();
    }

    public function writeHeader($header)
    {
        return $this->writeRow($header);
    }

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

    public function writeColumn($data = [], $col = 'A')
    {
        $rowNum = $this->hasHeader ? 2 : 1;
        foreach ($data as $item) {
            $cell = $col . $rowNum;
            $this->excel->setActiveSheetIndex(0)->setCellValue($cell, $item);
            $rowNum++;
        }
        return $this;
    }

    public function save($ext = 'xlsx', $path = false)
    {
        $this->setStoragePath($path);

        $this->ext = $ext;

        $file = $this->path . '/' . $this->file . '.' .$this->ext;

        $this->_setFormat();
        $this->_setWriter();

        $this->writer->save($file);
    }

    protected function _setFormat()
    {
        $this->format = $this->helper->getFormatByExtension($this->ext);
    }

    protected function _setWriter()
    {
        $this->writer = PHPExcel_IOFactory::createWriter($this->excel, $this->format);

        if ($this->format == 'CSV') {
            $this->writer->setDelimiter(',');
            $this->writer->setEnclosure('"');
            $this->writer->setLineEnding("\r\n");
        }
    }

    public function injectExcel($excel)
    {
        $this->excel = $excel;
    }

    public function setFileName($file)
    {
        $this->file = $file;
    }

    public function setTitle($file)
    {
        $this->title = $file;
        $this->excel->getProperties()->setTitle($this->title);
    }

    private function setStoragePath($path)
    {
        $path = $path ? $path : '../result';

        $this->path = rtrim($path, '/');

        if(!is_dir($this->path))
            mkdir($path, 0777);

        if(!is_writable($this->path))
            chmod($this->path, 0777);
    }
}