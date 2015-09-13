<?php namespace ExcelWorker;

use Exception;

const FILE_NOT_FOUND        = 10;
const PATH_NOT_FOUND        = 11;
const WRONG_EXTENSION       = 12;
const FAIL_TO_OPEN          = 20;
const FAIL_TO_WRITE         = 21;
const FAIL_TO_READ          = 22;
const WRONG_ROW             = 30;
const WRONG_COLUMN          = 31;
const WRONG_CELL            = 32;

/**
 * Class ExcelWorkerException
 * @package     forehalo/excel-worker
 * @version     1.0.0
 * @copyright   Copyright (c) 2015 forehalo <http://www.forehalo.top>
 * @author      forehalo <forehalo@gmail.com>
 * @license     http://www.gnu.org/licenses/lgpl.html   LGPL
 */
class ExcelWorkerException extends Exception
{
    /**
     * Exception handler.
     *
     * @param $code
     * @param $message
     * @param $file
     * @param $line
     * @throws Exception
     */
    public function ExceptionHandler($code, $message, $file, $line)
    {
        $e = new self($message, $code);
        $e->line = $line;
        $e->file = $file;
        throw $e;
    }


}