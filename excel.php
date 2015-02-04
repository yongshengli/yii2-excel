<?php

/**
 * cls_php_excel.php
 * @author: liyongsheng
 * @email： liyongsheng@huimai365.com
 * @date: 2015/1/15
*/

require_once 'PHPExcel.php';
require_once 'manage/includes/az.php';

/**
 * Class excel
 */
class excel extends PHPExcel
{

    private $azMap = array();

    /**
     * @param array $data
     */
    public function setAZMap(array $data)
    {
        $this->azMap = AZ::getAZ($data);
    }

    /**
     * 生成excel 标题行
     * @param array|arrayObject $data
     * @param int $line 标题所在的行
     * @throws Exception
     */
    public function createTitle($data, $line=1)
    {
        if(empty($this->azMap)){
            $this->setAZMap($data);
        }
        $activeSheet = $this->setActiveSheetIndex(0);
        foreach($data as $key=>$val){
            $activeSheet->setCellValue($this->azMap[$key].$line, $val);
        }
    }

    /**
     * 生成body
     * @param array|arrayObject $data
     * @param int $line 内容的起始行
     * @throws Exception
     */
    public function createBody($data, $line=2)
    {
        $activeSheet = $this->setActiveSheetIndex(0);

        foreach($data as $row){
            foreach ($row as $key=>$val){
                if(isset($this->azMap[$key])){
                    $activeSheet->setCellValue($this->azMap[$key].$line, $val);
                }
            }
            $line++;
        }
    }

}