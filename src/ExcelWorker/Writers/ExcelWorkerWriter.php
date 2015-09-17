<?php namespace ExcelWorker\Writers;

use PHPExcel;

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
    protected $path = '../../result/';

    protected $excel;

    protected $writer;

    protected $ext = [
        'xlsx' => 'Excel2007',
        'xls' => 'Excel5',
        'csv' => 'CSV',
        'html' => 'HTML',
    ];

    public function __construct(PHPExcel $excel)
    {
        $this->excel = $excel;
    }

    public function writeHeader($header)
    {
        return self::writeRow(1, $header);
    }

    public function writeRow($rowNum = 1, $data = [])
    {
        $column = 'A';

        foreach ($data as $item) {
            $cell = $column . $rowNum;
            $this->excel->setActiveSheetIndex(0)
                ->setCellValue($cell, $item);
            $column++;
        }
        return $this;
    }

    public function writeColumn($col = 'A', $data = [], $append = false)
    {
    }

    public function save($file = null, $ext = 'xlsx')
    {
        $this->writer = \PHPExcel_IOFactory::createWriter($this->excel, $this->ext[$ext]);
        if (!isset($file)) {
            if (!is_dir($this->path))
                mkdir($this->path);
            $this->writer->save($this->path . date('Y-m-d') . 'xlsx');
        } else {
            $this->writer->save($file);
        }
    }
}