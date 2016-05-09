

phpexcle 写入扩展类依赖phpexcel(https://github.com/PHPOffice/PHPExcel)

用法例子
```
$arr = [
    [
        'name'=>'id',
        'text'=>'ID'
    ],
    [
        'name'=>'name',
        'text'=>'名称'
    ],
    [
        'name'=>'price',
        'text'=>'价格'
    ],
];
$fileName = 'demo.xlsx';
(new Excel())->setFieldMap($arr)->download($fileName, $res['rows']);
```