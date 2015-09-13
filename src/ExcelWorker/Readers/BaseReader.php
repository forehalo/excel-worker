<?php namespace ExcelWorker\Readers;

use ExcelWorker\Exception;
use ExcelWorker\Exception\ExcelWorkerException;

/**
 * Class BaseReader.php
 * @package     forehalo/excel-worker
 * @version     1.0.0
 * @copyright   Copyright (c) 2015 forehalo <http://www.forehalo.top>
 * @author      forehalo <forehalo@gmail.com>
 * @license     http://www.gnu.org/licenses/lgpl.html   LGPL
 */
abstract class BaseReader implements IReader
{
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

    /**
     * Constructor
     * Load file into $content and parse the first row into $header.
     *
     * @param string $file
     */
    public function __construct($file = null)
    {
        if (isset($file)) {
            $this->file = $file;
            $this->load($file);
        }
    }

    /**
     * Get all the content as a index array.
     *
     * @return array
     * @throws ExcelWorkerException     When nothing loaded.
     */
    public function get()
    {
        if ($this->content == []) {
            throw new ExcelWorkerException('No content in this file', EMPTY_FILE);
        }
        return $this->content;
    }

    /**
     * Get one row by the given row number.
     *
     * @param int $row number of row
     * @return array
     * @throws ExcelWorkerException     When a number given less than 1 or greater than count of row.
     */
    public function getRow($row)
    {
        if (!$this->rowExist($row - 1)) {
            throw new ExcelWorkerException('Wrong row number', WRONG_ROW);
        }
        return $this->content[$row - 1];
    }

    /**
     * Get the first row. Perhaps the second row if header exist.
     *
     * @throws ExcelWorkerException
     */
    public function getFirst()
    {
        $this->getRow(1);
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
        if (!$this->columnExist($col - 1)) {
            throw new ExcelWorkerException('Wrong column number', WRONG_COLUMN);
        }
        return array_column($this->content, $col - 1);
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
        if (!$this->cellExist($row, $col)) {
            throw new ExcelWorkerException('Wrong cell index', WRONG_CELL);
        }
        return $this->content[$row - 1][$col - 1];
    }

    /**
     * Get the header if specified.
     *
     * @return array
     */
    public function getHeader()
    {
        if ($this->header != [])
            return $this->header;
    }

    /**
     * Judge whether the given row exist in $content.
     *
     * @param int $row number of row
     * @return bool
     */
    public function rowExist($row)
    {
        return array_key_exists($row, $this->content);
    }

    /**
     * Judge Whether the given column exist in $content.
     *
     * @param int $col number of column
     * @return bool
     */
    public function columnExist($col)
    {
        return array_key_exists($col, $this->content[0]);
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
        return isset($this->content[$row][$col]);
    }

    public function canReader()
    {
        if (!file_exists($this->file) || !is_readable($this->file)) {
           throw new ExcelWorkerException("File: $this->file is not exist or readable", FAIL_TO_READ);
        }
        return true;
    }
}