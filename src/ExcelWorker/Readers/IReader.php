<?php namespace ExcelWorker\Readers;

/**
 * Interface IReader.php
 * @package     forehalo/excel-worker
 * @version     1.0.0
 * @copyright   Copyright (c) 2015 forehalo <http://www.forehalo.top>
 * @author      forehalo <forehalo@gmail.com>
 * @license     http://www.gnu.org/licenses/lgpl.html   LGPL
 */
interface IReader
{
    public function load($file);
    public function get();
    public function getRow($row);
    public function getColumn($col);
    public function getCell($row, $col);
    public function rowExist($row);
    public function columnExist($col);
    public function cellExist($row, $col);

}