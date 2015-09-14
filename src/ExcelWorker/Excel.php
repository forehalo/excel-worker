<?php namespace ExcelWorker;

use ExcelWorker\ExcelWorker;

/**
 * Class Excel.php
 * @package     forehalo/excel-worker
 * @version     1.0.0
 * @copyright   Copyright (c) 2015 forehalo <http://www.forehalo.top>
 * @author      forehalo <forehalo@gmail.com>
 * @license     http://www.gnu.org/licenses/lgpl.html   LGPL
 */
class Excel
{
    private static $excel = null;

    private function __construct()
    {
        self::$excel = new ExcelWorker();
    }

    private static function getInstance(){
        if (!(self::$excel instanceof ExcelWorker)) {
            self::$excel = new ExcelWorker();
        }
        return self::$excel;
    }

    public static function create($file)
    {
        self::getInstance()->create($file);
    }

    public static function load($file)
    {
        self::getInstance()->load($file);
        return self::getInstance()->get();
    }
}