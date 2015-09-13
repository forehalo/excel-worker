<?php namespace ExcelWorker;

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
    public function __construct()
    {
        $this->excel = new PHPExcel();
    }
}