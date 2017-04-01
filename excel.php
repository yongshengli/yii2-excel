<?php
/**
 * Created by PhpStorm.
 * User: david
 * Date: 2017/3/31
 * Time: 18:05
 * Email:liyongsheng@meicai.cn
 */

namespace app\components;
use yii\base\Component;
use PHPExcel;
use PHPExcel_Cell;
use yii\base\Exception;

/**
 * Class Excel
 *
 * @property int $activeSheetIndex
 * @method getSheet($pIndex = 0) 根据指定索引获取工作簿
 * @method getActiveSheet() 获取当前工作簿
 * @method getSheetByName($pName = '')
 */
class Excel extends Component
{
    /**
     * 当前行
     * @var int
     */
    public $row = [];

    /**
     * @var \PHPExcel
     */
    public $phpExcel = null;

    /**
     * 字段map
     * @var array
     */
    private $_fieldMap =[];

    public function init()
    {
        parent::init();
        $this->phpExcel = new PHPExcel();
    }

    /**
     * @param array $arr
     * @param int $activeSheetIndex 当前工作簿
     * @return $this
     */
    public function setFieldMap(array $arr, $activeSheetIndex=null)
    {
        if($activeSheetIndex !==null){
            $this->activeSheetIndex = $activeSheetIndex;
        }
        for ($i = 0; $i < count($arr); $i++) {
            $arr[$i]['cell'] = PHPExcel_Cell::stringFromColumnIndex($i);
        }
        $this->_fieldMap[$this->activeSheetIndex] = $arr;
        $this->row[$this->activeSheetIndex] = 1;
        return $this;
    }

    /**
     * 获取 字段映射map
     * @return mixed
     */
    public function getFieldMap()
    {
        if(isset($this->_fieldMap[$this->activeSheetIndex])){
            return $this->_fieldMap[$this->activeSheetIndex];
        }else{
            return null;
        }
    }
    /**
     * 设置空行
     * 设置当前行
     *
     * @param int $num
     * @return $this
     */
    public function rowIncrease($num=1)
    {
        $this->row[$this->activeSheetIndex] +=$num;
        return $this;
    }

    /**
     * 当前行
     */
    public function currentRow()
    {
        return $this->row[$this->activeSheetIndex];
    }
    /**
     * 写入头信息
     * @return $this
     */
    public function writeHeader()
    {
        $worksheetObj = $this->phpExcel->setActiveSheetIndex($this->activeSheetIndex);
        $fieldNum  = count($this->fieldMap);
        for ($i = 0; $i < $fieldNum; $i++) {
            $worksheetObj->setCellValue($this->fieldMap[$i]['cell'] . $this->currentRow(), $this->fieldMap[$i]['text']);
        }
        $this->rowIncrease();
        return $this;
    }
    /**
     * 设置数据
     * @param array $data
     * @param bool $withHeader 是否需要写入头
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function setData(array $data=array(), $withHeader=true)
    {
        if(empty($data)){
            return $this;
        }
        if($withHeader) {
            $this->writeHeader();
        }
        $worksheetObj = $this->phpExcel->setActiveSheetIndex($this->activeSheetIndex);
        foreach ($data as $item) {
            foreach ($this->fieldMap as $field) {
                if (isset($item[$field['name']])) {
                    if(isset($field['func'])){
                        $item[$field['name']] = call_user_func_array($field['func'], [$item[$field['name']]]);
                    }
                    $worksheetObj->setCellValue($field['cell'] . $this->currentRow(), $item[$field['name']]);
                }
            }
            $this->rowIncrease();
        }
        return $this;
    }
    /**
     * 导出
     * @param string $fileName
     * @param array $data
     * @throws Exception
     * @throws \PHPExcel_Exception
     * @throws \PHPExcel_Reader_Exception
     */
    public function download($fileName, $data=array())
    {
        if(empty($this->_fieldMap)){
            throw new Exception('fieldMap 未设置');
        }
        // Redirect output to a client’s web browser (Excel2007)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="' . $fileName . '"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');
        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $this->setData($data);

        $objWriter = \PHPExcel_IOFactory::createWriter($this->phpExcel, 'Excel2007');
        $objWriter->save('php://output');
    }
    /**
     * @param $activeSheetIndex
     * @return $this
     */
    public function setActiveSheetIndex($activeSheetIndex=0)
    {
        $this->phpExcel->setActiveSheetIndex($activeSheetIndex);
        return $this;
    }

    /**
     * 设置当前活动工作簿索引
     * @param string $pValue
     * @throws \PHPExcel_Exception
     * @return $this
     */
    public function setActiveSheetIndexByName($pValue = '')
    {
        $this->phpExcel->setActiveSheetIndexByName($pValue);
        return $this;
    }

    /**
     * 获取当前工作簿索引
     */
    public function getActiveSheetIndex(){
        return $this->phpExcel->getActiveSheetIndex();
    }

    /**
     * 添加工作簿
     * @param \PHPExcel_Worksheet $pSheet
     * @param null $iSheetIndex
     * @return $this
     * @throws \PHPExcel_Exception
     */
    public function addSheet(\PHPExcel_Worksheet $pSheet, $iSheetIndex = null)
    {
        $this->phpExcel->addSheet($pSheet, $iSheetIndex);
        $this->phpExcel->setActiveSheetIndex($iSheetIndex);
        return $this;
    }
    /**
     * 魔术方法
     * @param string $method
     * @param array $params
     * @return mixed
     */
    public function __call($method, $params)
    {
        return call_user_func_array([$this->phpExcel, $method], $params);
    }
}