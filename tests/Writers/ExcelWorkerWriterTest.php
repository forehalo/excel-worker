<?php

use ExcelWorker\ExcelWorker;

class ExcelWorkerWriterTest extends PHPUnit_Framework_TestCase
{
    public function setUp()
    {
        $this->worker = new ExcelWorker();
        $this->writer = $this->worker->create(__DIR__ . '/files/test.xlsx');
    }
}