#Excel-Wroker V2.0

a simple library to work with excels based on PHPExcel.

##Feature

- muti-workesheets, possible:

```php
$worker->setSelectedSheets(['Sheet1', 'Sheet2'])->load('filename')->get();
```

- specify columns, possible: 

```php
$worker->load('filename', true)->get(['col1', 'col2']);
$worker->load('filename')->get([0, 1, 2]);
```

- limit start and end row, possible:

```php
//skip the first 5 rows.
$reader->skip(5)->get();
//only fetch 6 rows
$reader->take(6)->get();
//skip the first 7 rows and fetch the following 8 rows
$reader->skip(7)->take(8)->get();
//or
$reader->limit(7, 8)->get();

//attention
$reader->limit(-1, 8) // no skip
$reader->limit(7, -1) // no take
```


##Installation

require this package in your `composer.json` and update composer. It will also download PHPExcel for you.

	"forehalo/excel-worker":"2.0.*"

##Usage

Before usage you should include `vendor/autoload.php` into your file.

```php
use ExcelWorker\ExcelWorker;
$worker = new ExcelWorker();

//Export
$rowNum = 2;
$writer = $worker->create('file'); //create() will return a Writer instance
$writer->writerRow(['a', 'b', 'c'], $rowNum);
$writer->writeColumn(['a', 'b', 'c'], $col); //both $col = 50 or $col = 'AX' could be accpeted.
$writer->save('xlsx');  //you may use ->save('xlsx', 'path') to specify the storage path.

//Import
//The second parameter is a bool value to tell whether header exists(probably the first row), default is false.
$worker->load('./path/filename.xlsx', true)->get();

//The load() method returns a Reader instance, so you could use as:
$reader = $worker->load('filename.xlsx');
$reader->get();
```

##License

![](http://i.imgur.com/8ZtPnc7.png)

http://www.gnu.org/licenses/lgpl.txt
