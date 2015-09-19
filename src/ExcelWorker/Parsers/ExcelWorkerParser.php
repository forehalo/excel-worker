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

    /**
     * PHPExcel object
     * @var \PHPExcel
     */
    protected $excel;

    /**
     * ExcelWorkerReader object
     * @var ExcelWorkerReader
     */
    protected $reader;

    /**
     * Columns to be parsed.
     * @var array
     */
    protected $column;

    /**
     * Whether has been parsed.
     * @var bool
     */
    protected $isParsed = false;

    /**
     * Default start row to be parsed.
     * @var int
     */
    protected $defaultStartRow = 1;

    /**
     * The current worksheet.
     * @var worksheet
     */
    protected $worksheet;

    /**
     * The current row.
     * @var Row
     */
    protected $row;

    /**
     * The current cell.
     * @var Cell
     */
    protected $cell;

    /**
     * Constructor
     * @param ExcelWorkerReader $reader
     */
    public function __construct(ExcelWorkerReader $reader)
    {
        $this->reader = $reader;
        $this->excel = $this->reader->excel;
    }

    /**
     * Parse file.
     * @param array $column
     * @return array
     */
    public function parseFile($column = [])
    {
        $content = [];
        $this->setSelectedColumn($column);
        if (!$this->isParsed) {
            foreach ($this->excel->getWorksheetIterator() as $this->worksheet) {
                $worksheet = $this->parseWorksheet();
                $title = $this->worksheet->getTitle();
                $content[$title] = $worksheet;
            }
        }

        $this->isParsed = true;
        return $content;
    }

    /**
     * Set column to be parsed.
     * @param $column
     */
    protected function setSelectedColumn($column)
    {
        $this->column = $column;
    }

    /**
     * Parse worksheet
     * @return array
     */
    protected function parseWorksheet()
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

    /**
     * Parse row
     * @return array
     */
    protected function parseRow()
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

    /**
     * get start row.
     * @return int
     */
    protected function getStartRow()
    {
        $startRow = $this->defaultStartRow;
        /*
        if($this->reader->hasHeader())
            $startRow++;
        $skip = $this->reader->getSkip();
        if ($skip > 0)
            $startRow += $skip;
        */
        return $startRow;
    }

}