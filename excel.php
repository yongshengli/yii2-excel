<?php

/**
 * Created by PhpStorm.
 * User: david
 * Date: 16/5/5
 * Time: 17:44
 * Email:liyongsheng@meicai.cn
 */
class Excel
{
    private $_fieldMap =[];

    /**
     * @param $arr
     * @return $this
     */
    public function setFieldMap($arr)
    {
        $this->_fieldMap = $arr;
        return $this;
    }

    /**
     * 导出
     * @param string $fileName
     * @param array $data
     * @throws CException
     * @throws PHPExcel_Exception
     * @throws PHPExcel_Reader_Exception
     */
    public function download($fileName, $data=array())
    {
        if(empty($this->_fieldMap)){
            throw new CException('fieldMap 未设置');
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

        $objPHPExcel = new PHPExcel();
        $worksheetObj = $objPHPExcel->setActiveSheetIndex(0);
        for ($i = 0; $i < count($this->_fieldMap); $i++) {
            $this->_fieldMap[$i]['cell'] = PHPExcel_Cell::stringFromColumnIndex($i);
            $worksheetObj->setCellValue($this->_fieldMap[$i]['cell'] . '1', $this->_fieldMap[$i]['text']);
        }
        $row = 2;
        foreach ($data as $item) {
            foreach ($this->_fieldMap as $field) {
                if (isset($item[$field['name']])) {
                    $worksheetObj->setCellValue($field['cell'] . $row, $item[$field['name']]);
                }
            }
            $row++;
        }

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->save('php://output');
    }
}