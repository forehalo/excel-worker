<?php namespace ExcelWorker\Exceptions;

use Exception;

class ExcelWorkerException extends Exception {
    public function ExceptionHandler($code, $desc, $file, $line)
    {
        $e = new Exception($desc, $code);
        $e->line = $line;
        $e->file = $file;
        throw $e;
    }
}