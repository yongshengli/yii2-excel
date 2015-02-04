<?php

/**
 * az.php
 * @author: liyongsheng
 * @email： liyongsheng@huimai365.com
 * @date: 2015/1/15
*/

/**
 * Class AZ
 * 将数字转化为AZ进制 最大支持 到 702 ZZ
 */
class AZ
{

    static $dic = array('A', 'B', 'C','D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q','R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z');

    /**
     * @param array $data
     * @return array
     */
    public static function getAZ(array $data)
    {
        $arr = array();
        $i = 0;
        while(list($key, $val) = each($data)){
            if($i>=702){
                break;
            }
            if($i<26){
                $arr[$key] = self::$dic[$i];
            }else{
                $k = $i%26;
                $d = intval(floor($i/26-1));
                $arr[$key] =  self::$dic[$d]. self::$dic[$k];
            }
            $i++;
        }
        return $arr;
    }

    /**
     * 将整数转换为AZ
     * @param int $int
     * @return bool|string
     */
    public static function int2AZ($int)
    {
        if($int>=702){
            return false;
        }
        if($int<26){
            return  self::$dic[$int];
        }else{
            $k = $int%26;
            $d = intval(floor($int/26-1));
            return self::$dic[$d]. self::$dic[$k];
        }
    }
}