<?php

use ExcelWorker\ExcelWorker;

class ExcelWorkerWriterTest extends PHPUnit_Framework_TestCase
{
    public function setUp()
    {
        $this->worker = new ExcelWorker();
        $this->writer = $this->worker->create('test');
    }

    public function testWriteRow()
    {
        $rowNum = 2;
        $this->writer->writeRow(['a', 'b', 'c'], $rowNum)->save('xlsx', __DIR__ . '/files/');
    }

    public function testGetColumn()
    {
        $colNum = 50;
        $this->assertEquals('AX', $this->writer->getColumn($colNum));
    }

    public function testWriteColumn()
    {
        $colNum = 50;
        $col = 'AY';
        $this->writer->writeColumn(['a', 'b', 'c'], $colNum);
        $this->writer->writeColumn(['a', 'b', 'c'], $col);
        $this->writer->save('xlsx', __DIR__ . '/files/');
    }

    public function testWriteCSV()
    {
        $this->writer->writeRow(['a', 'b', 'c'])->save('csv', __DIR__ . '/files/');
    }

}