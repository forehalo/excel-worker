<?php

use ExcelWorker\ExcelWorker;
class ExcelWorkerTest extends PHPUnit_Framework_TestCase
{
    public function testConstructor()
    {
        $worker = new ExcelWorker();
        $this->assertInstanceOf('ExcelWorker\ExcelWorker', $worker);
    }
}