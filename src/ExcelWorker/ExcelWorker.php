<?php namespace ExcelWorker;

use ExcelWorker\Readers\ExcelWorkerReader;
use ExcelWorker\Writers\ExcelWorkerWriter;
use PHPExcel;
/**
 * Class ExcelWorker.php
 * @package     forehalo/excel-worker
 * @version     1.0.0
 * @copyright   Copyright (c) 2015 forehalo <http://www.forehalo.top>
 * @author      forehalo <forehalo@gmail.com>
 * @license     http://www.gnu.org/licenses/lgpl.html   LGPL
 */
class ExcelWorker
{
    private $excel;

    private $reader;

    private $writer;

    public function __construct()
    {
        $this->excel = new PHPExcel();
        $this->reader = new ExcelWorkerReader();
        $this->writer = new ExcelWorkerWriter($this->excel);
    }

    public function create($file)
    {

    }

    public function load($file)
    {
        $this->reader->injectExcel($this->excel);
        $this->reader->load($file);

        return $this->reader;
    }
}