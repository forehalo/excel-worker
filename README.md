#Excel-Wroker

a simple library to work with excels based on PHPExcel.

##Installation

require this package in your `composer.json` and update composer. It will also download PHPExcel for you.

	"forehalo/excel-worker":"1.0.0"

##Usage

Before usage you should include `vendor/autoload.php` into your file.

```php
use ExcelWorker\ExcelWorker;
$worker = new ExcelWorker();

//Export
$worker->create('filename')
	   ->WriterRow([
			'a',
			'b',
			'c'
		])->save('xlsx');  //you may use ->save('xlsx', 'path') to specify the storage path.

//Import
$worker->load('./path/filename.xlsx')->all();
```

##further

This library is still in developing. More feature will be add in, and bug will be fixed.

Thanks for using.

##License

![](http://i.imgur.com/8ZtPnc7.png)

http://www.gnu.org/licenses/lgpl.txt