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
    protected $columns;

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
     * header
     * @var array
     */
    protected $header = [];

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
     * @param array $columns
     * @return array
     */
    public function parseFile($columns = [])
    {
        $content = [];
        $this->setSelectedColumn($columns);
        if (!$this->isParsed) {
            $iterator = $this->excel->getWorksheetIterator();

            foreach ($iterator as $this->worksheet) {

                if ($this->reader->isSelected($iterator->key())) {

                    $worksheet = $this->parseWorksheet();
                    if (!empty($worksheet)) {
                        $title = $this->worksheet->getTitle();
                        $content[$title] = $worksheet;
                    }
                }
            }
        }

        $this->isParsed = true;
        return $content;
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
            $parsed = $this->parseRow();
            if (!empty($parsed))
                $content[$i] = $parsed;
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
        if ($this->reader->hasHeader) {
            $this->initHeader();
        }
        $cells = $this->row->getCellIterator();
        $i = 0;
        foreach ($cells as $this->cell) {
            $header = $this->reader->hasHeader ? $this->header[$i] : $i;
            if ($this->needParsed($header)) {
                $content[$header] = $this->cell->getValue();
            }
            $i++;
        }
        return $content;

    }

    protected function initHeader()
    {
        $row = $this->worksheet->getRowIterator(1)->current();
        $header = [];
        foreach ($row->getCellIterator() as $item) {
            $header[] = $item->getValue();
        }
        $this->header = $header;
    }

    protected function needParsed($index)
    {
        if (empty($this->columns))
            return true;
        return in_array($index, $this->getSelectedColumns(), true);
    }

    /**
     * Set columns to be parsed.
     * @param array $columns
     */
    protected function setSelectedColumn($columns)
    {
        $this->columns = $columns;
    }

    /**
     * Get selected columns
     * @return array
     */
    protected function getSelectedColumns()
    {
        return $this->columns;
    }

    /**
     * get start row.
     * @return int
     */
    protected function getStartRow()
    {
        $startRow = $this->defaultStartRow;
        if ($this->reader->hasHeader)
            $startRow++;
//
//        $skip = $this->reader->getSkip();
//        if ($skip > 0)
//            $startRow += $skip;
        return $startRow;
    }

}