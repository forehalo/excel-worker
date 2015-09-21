<?php

use ExcelWorker\ExcelWorker;

class MultiSheetsTest extends PHPUnit_Framework_TestCase
{
    public function setUp()
    {
        $this->worker = new ExcelWorker();
        $this->reader = $this->worker->load(__DIR__ . '/files/multiSheets.xlsx', true);
        $this->assertInstanceOf('ExcelWorker\Readers\ExcelWorkerReader', $this->reader);
    }

    public function testGetSheet()
    {
        $this->assertEquals(['Sheet1', 'Sheet2'], array_keys($this->reader->get()));
    }

    public function testSelectSheetByIndex()
    {
        $this->assertEquals(['Sheet2'], array_keys($this->worker->setSelectedSheetsByIndex([1])->load(__DIR__ . '/files/multiSheets.xlsx', true)->get()));
    }

    public function testSelectSheetByName()
    {
        $this->assertEquals(['Sheet1'], array_keys($this->worker->setSelectedSheets(['Sheet1'])->load(__DIR__ . '/files/multiSheets.xlsx', true)->get()));
    }
}