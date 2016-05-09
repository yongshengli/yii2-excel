excel
============================================
phpexcle 写入扩展类依赖phpexcel(https://github.com/PHPOffice/PHPExcel)

用法例子
```
$fieltMap = [
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
$data = [
    [
        'id' => 1,
        'name'=>'名称1',
        'price'=>13,
    ],
    [
        'id' => 2,
        'name'=>'名称2',
        'price'=>13,
    ],
    [
        'id' => 3,
        'name'=>'名称3',
        'price'=>23,
    ],
];
$fileName = 'demo.xlsx';
(new Excel())->setFieldMap($fieltMap)->download($fileName, $data);
```