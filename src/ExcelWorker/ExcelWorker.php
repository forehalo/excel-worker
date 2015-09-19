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
    /**
     * PHPExcel object
     * @var PHPExcel
     */
    private $excel;

    /**
     * Reader
     * @var ExcelWorkerReader
     */
    private $reader;

    /**
     * Writer
     * @var ExcelWorkerWriter
     */
    private $writer;

    /**
     * Constructor
     */
    public function __construct()
    {
        $this->excel = new PHPExcel();
        $this->reader = new ExcelWorkerReader();
        $this->writer = new ExcelWorkerWriter();
    }

    /**
     * Create a new file.
     * @param $file
     * @return ExcelWorkerWriter
     */
    public function create($file)
    {
        //Inject PHPExcel object into writer
        $this->writer->injectExcel($this->excel);

        $this->writer->setFilename($file);
        $this->writer->setTitle($file);

        return $this->writer;
    }

    /**
     * Load a exist file.
     * @param $file
     * @return ExcelWorkerReader
     */
    public function load($file)
    {
        //Inject PHPExcel object into reader
        $this->reader->injectExcel($this->excel);
        $this->reader->load($file);

        return $this->reader;
    }


    public function setSheetSelected($sheets = [])
    {
        $sheets = is_array($sheets) ? $sheets : func_get_args();
        $this->reader->setSelectedSheets($sheets);

        return $this;
    }
}