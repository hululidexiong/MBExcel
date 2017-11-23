## Example： ##
```php
use MBExcel\Factory;

define('APP_ROOT' , __DIR__);

//传入excel的保存路径
$excel = new Factory( APP_ROOT );

$data = array(
    array(
        'name' => '武帅',
        'phone' => '13334343344',
    ),
    array(
        'name' => '乔辉',
        'phone' => '13334343345',
    )
);
$e = $excel->easyExport($data);
//var_dump($e);

//数据量大的时候允许分段写入
$e = $excel->easyExport($data , 'first');
$e = $excel->easyExport($data , 'continue' , $e['ident']);
$e = $excel->easyExport(array() , 'end' , $e['ident']);
//var_dump($e);

```