<?php

use ExcelWorker\ExcelWorker;
use ExcelWorker\Readers\ExcelWorkerReader;

class ExcelWorkerReaderTest extends PHPUnit_Framework_TestCase
{
    public $reader;

    public $worker;

    public function setup()
    {
        $this->worker = new ExcelWorker();
        $this->reader = $this->worker->load(__DIR__ . '/files/test.xlsx', true);
        $this->assertInstanceOf('ExcelWorker\Readers\ExcelWorkerReader', $this->reader);
    }

    public function testGet()
    {
        $expected = [
            'Sheet1' => [
                [
                    'h1' => 1,
                    'h2' => 2,
                    'h3' => 3,
                    'h4' => 4
                ],
                [
                    'h1' => 5,
                    'h2' => 6,
                    'h3' => 7,
                    'h4' => 8
                ],
                [
                    'h1' => 9,
                    'h2' => 10,
                    'h3' => 11,
                    'h4' => 12
                ]
            ]
        ];
        $this->assertEquals($expected, $this->reader->get());
        $this->assertEquals($expected, $this->reader->all());
    }

    public function testSkip()
    {
        $expected = [
            'Sheet1' => [
                [
                    'h1' => 5,
                    'h2' => 6,
                    'h3' => 7,
                    'h4' => 8
                ],
                [
                    'h1' => 9,
                    'h2' => 10,
                    'h3' => 11,
                    'h4' => 12
                ]
            ]
        ];
        $this->assertEquals($expected, $this->reader->skip(1)->get());
    }

    public function testTake()
    {
        $expected = [
            'Sheet1' => [
                [
                    'h1' => 5,
                    'h2' => 6,
                    'h3' => 7,
                    'h4' => 8
                ]
            ]
        ];
        $this->assertEquals($expected, $this->reader->skip(1)->take(1)->get());
    }

    public function testLimit()
    {
        $expected = [
            'Sheet1' => [
                [
                    'h1' => 5,
                    'h2' => 6,
                    'h3' => 7,
                    'h4' => 8
                ]
            ]
        ];
        $this->assertEquals($expected, $this->reader->limit(1, 1)->get());
    }
}