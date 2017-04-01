yii2-excel 用于excel,cvs数据导出
============================================
phpexcle 写入扩展类 依赖phpexcel(https://github.com/PHPOffice/PHPExcel) 和yii 2.0

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
//(new Excel())->setFieldMap($fieltMap)->setData($data)->download($fileName);
//(new Excel())->setFieldMap($fieldMap1)->setData($data1)->setFieldMap($fieldMap2)->setData($data2)->download('filename');
```