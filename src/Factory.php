<?php
/**
 * Created by PhpStorm.
 * User: mhx
 * Date: 2017/11/22
 * Time: 11:11
 */

namespace MBExcel;

use BBear\Tools\Base\Tools;

class Factory
{
    private $_FilePath;
    private $_Path;
    private $_app_dir;

    function __construct( $app_dir = __DIR__){

        $this->_app_dir = $app_dir;
        $this->_Path = DIRECTORY_SEPARATOR. 'excel';
        //日志文件夹
        $this->_FilePath = $this->_app_dir .$this->_Path;
    }



//-----------------------------------------------------------------------------------------------

    /**
     * 导出为web格式的excel文件 无法直接导入
     * @param $column array 纵列名称 一维数组
     * @param $data array 对应的数据 二维数组 注： 顺序要和 $column 对应
     * @param $step string 生成步骤 分段生成文件 传 frist  ... continue ....end  或者 一次生成文件 full
     * @param $ident string 当 $step 为frist 时 返回数组中会 包含 ident 参数,用于下一次调用时传入
     * @param $filename string 文件名称
     * @return array
     */
    function easyExport($data,$step='full',$ident = '',$filename=''){

        $column = @array_keys($data[0]);
        $r = array();
        switch($step){
            case 'full':
                $content = $this->TableHeader();
                $content .= $this->TableColumn($column);
                $content .= $this->TableData($data);
                $content .= $this->TableEnd();
                $p = $this->Mkdir('',$filename);
                $this->WriteFile($p['fullPath'],$content);
                break;
            case 'first':
                $content = $this->TableHeader();
                $content .= $this->TableColumn($column);
                $content .= $this->TableData($data);
                $p = $this->Mkdir('',$filename);
                $this->WriteFile($p['fullPath'],$content);
                break;
            case 'end':
                $content = $this->TableData($data);
                $p = $this->Mkdir($ident,$filename);
                $this->WriteFile($p['fullPath'],$content,1);
                break;
            case 'continue':
                $content = $this->TableData($data);
                $content .= $this->TableEnd();
                $p = $this->Mkdir($ident,$filename);
                $this->WriteFile($p['fullPath'],$content,1);
                break;
            default:
                throw new \Exception(' step is invalid');
        }
        $r = $p;
        return $r;
    }


    /**
     * 读取Excel 成Array格式
     * @param $inputFileName Excel文件(包含路径)
     * @return array
     */
    function reader($inputFileName){
        //$inputFileName = $this->_app_dir.'/1319.xls';  正常Excel
        $objPHPExcel = \PHPExcel_IOFactory::load($inputFileName);
        $sheetData = $objPHPExcel->getActiveSheet()->toArray(null,true,true,true);
        //  $sheetData = $this->read( $this->_app_dir.'/1319.xls', 'xls');
        return $sheetData;
    }
//---------------------------------End of public function---------------------------------------------

    private function nDirectory($path = ''){
        $path = $path ? DIRECTORY_SEPARATOR.$path : '';
        if(!is_dir($this->_FilePath.$path)){
            mkdir($this->_FilePath.$path,0755,true);
        }

//        if(!is_dir($this->_FilePath)){
//            mkdir($this->_FilePath,0755);
//        }
//        if($path && !is_dir($this->_FilePath.$path)){
//            mkdir($this->_FilePath.$path,0755);
//        }
    }

    private function TableHeader(){
        $tableHeader = <<<EOF
<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel"
      xmlns="http://www.w3.org/TR/REC-html40">
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd"
>
<html>
<head>
    <meta http-equiv="Content-type" content="text/html;charset=UTF-8"/>
    <style id="Classeur1_16681_Styles"></style>
</head>
<body>
<div id="Classeur1_16681" align=center x:publishsource="Excel">
    <table x:str border=1 cellpadding=0 cellspacing=0 width=100% style="border-collapse: collapse">
EOF;
        return $tableHeader;
    }

    private function TableColumn($column){

        if(empty($column)){
            throw new \Exception('Not found TableColumn ,because $column is invalid');
        }
        //输出表头
        $content = '';
        $content .= '<tr>';
        foreach ($column as $nk => $nv) {
            $content .= '<td class=xl2216681 nowrap width="120">' . $nv . "</td>";
        }
        $content .= '</tr>';
        return $content;
    }

    private function TableData($data){

        $content = '';
        foreach($data as $item){
            $content .= '<tr>';
            foreach ($item as $tableItem) {
                $content .= '<td class=xl2216681 nowrap>' . $tableItem . '</td>';
            }
            $content .= '</tr>';
        }
        return $content;
    }

    private function TableEnd(){
        $content = <<<EOF
</table>
</div>
</body>
</html>
EOF;
        return $content;
    }


    private function Mkdir($ident = '' , $filename = ''){
        $r = array();
        if(!$ident){
            if(!$filename){
                $filename = rand(1000,9999).'.xls';
            }
            $path = date('YmdHis',time());
        }else{
            $Auth = explode("\t", Tools::sDecode( $ident));
            if($Auth[0] != 1){
                C::t('G')->rJson(array('Error' => 1 , 'Msg' => 'ident 不合法'));
            }
            $ident = $Auth[1];
            $i = explode(DIRECTORY_SEPARATOR,$ident);
            $path = $i[0];
            $filename = $i[1];
        }
        $fullfile = $path.DIRECTORY_SEPARATOR.$filename;
        $ident = Tools::sEncode(implode("\t",array(1,$fullfile)));
        $r['ident'] = $ident;
        $r['path'] = $path;
        $this->nDirectory($r['path']);
        $r['url'] = $this->_Path.DIRECTORY_SEPARATOR.$fullfile;
        $r['fullPath'] = $this->_FilePath.DIRECTORY_SEPARATOR.$fullfile;
        return $r;
    }
    private function WriteFile($fullPath,$content,$isAppend = false){
        if($isAppend){
            echo $fullPath;
            file_put_contents($fullPath,$content,FILE_APPEND);
        }else{
            file_put_contents($fullPath,$content);
        }
    }

//    private function read($filename, $file_type, $encode='utf-8'){
////        $objReader = PHPExcel_IOFactory::createReader('Excel5');
//        if(strtolower ( $file_type )=='xls')//判断excel表类型为2003还是2007
//        {
//            $objReader = PHPExcel_IOFactory::createReader('Excel5');
//        }elseif(strtolower ( $file_type )=='xlsx')
//        {
//            $objReader = PHPExcel_IOFactory::createReader('Excel2007');
//        }
//        $objReader->setReadDataOnly(true);
//        $objPHPExcel = $objReader->load($filename);
//        $objWorksheet = $objPHPExcel->getActiveSheet();
//        $highestRow = $objWorksheet->getHighestRow();
//        $highestColumn = $objWorksheet->getHighestColumn();
//        $highestColumnIndex = PHPExcel_Cell::columnIndexFromString($highestColumn);
//        $excelData = array();
//        for ($row = 1; $row <= $highestRow; $row++) {
//            for ($col = 0; $col < $highestColumnIndex; $col++) {
//                $excelData[$row][] =(string)$objWorksheet->getCellByColumnAndRow($col, $row)->getValue();
//            }
//        }
//        return $excelData;
//    }
}