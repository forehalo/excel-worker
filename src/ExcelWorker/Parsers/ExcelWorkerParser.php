<?php namespace ExcelWorker\Parsers;

use ExcelWorker\Readers\ExcelWorkerReader;

/**
 * Class ExcelWorkerParser.php
 * @package     forehalo/excel-worker
 * @version     1.0.0
 * @copyright   Copyright (c) 2015 forehalo <http://www.forehalo.top>
 * @author      forehalo <forehalo@gmail.com>
 * @license     http://www.gnu.org/licenses/lgpl.html   LGPL
 */
class ExcelWorkerParser
{

    protected $excel;

    private $reader;

    private $column;

    private $hasParsed = false;

    private $defaultStartRow = 1;

    private $worksheet;

    private $row;

    private $cell;

    public function __construct(ExcelWorkerReader $reader)
    {
        $this->reader = $reader;
        $this->excel = $this->reader->excel;
    }

    public function parseFile($column = [])
    {
        $content = [];
        $this->setSelectedColumn($column);

        if (!$this->hasParsed) {
            $this->worksheet = $this->excel->getWorksheetIterator()->current();

            $worksheet = $this->parseWorksheet(0);
            $content[0] = $worksheet;
        }

        $this->hasParsed = true;
        return $content;
    }

    private function setSelectedColumn($column)
    {
        $this->column = $column;
    }

    private function parseWorksheet()
    {
        $content = [];
        $rows = $this->worksheet->getRowIterator($this->getStartRow());
        $i = 0;
        foreach ($rows as $this->row) {
            $content[$i] = $this->parseRow();
            $i++;
        }
        return $content;
    }

    private function parseRow()
    {
        $content = [];
        $cells = $this->row->getCellIterator();
        $i = 0;
        foreach ($cells as $this->cell) {
            $content[$i] = $this->cell->getValue();
            $i++;
        }
        return $content;

    }

    private function getStartRow()
    {
        $startRow = $this->defaultStartRow;

//        if($this->reader->hasHeader())
//            $startRow++;
//        $skip = $this->reader->getSkip();
//        if ($skip > 0)
//            $startRow += $skip;

        return $startRow;
    }

}