```php
<?php
/**
 * 
 * 订单数据服务费
 * 
 */

class DataServiceFeeController extends BaseController{
    
    protected $_pageSize = 15;
    protected $_orderHeadModel;
    protected $moth_e = array(
        1  => "January",
        2  => "February",
        3  => "March",
        4  => "April",
        5  => "May",
        6  => "June",
        7  => "July",
        8  => "August",
        9  => "September",
        10 => "October",
        11 => "November",
        12 => "December");
    public function __construct()
    {
        $this->_orderHeadModel = new OrderHeadModel();
        return parent::__construct("finance");
    }
    /**
     * 
     * 订单数据服务费列表
     * 
     */
    public function actionIndex(){
        $ecmsTaxModel = new EcmsTaxReceiptModel();
        $dataserviceModel = new DataServiceFeeModel();
        $plamfromType = CDict::$platform_type;
        $getData = $this->_get();
        $curr_page = $getData['page'] ? $getData['page'] : 1;
        $platform_type = $this->_get('platform_type');
        // 平台类型
        $platfromType = CDict::$platform_type;
        $whereSql = '';
        if($getData['platform_type'] !='0'&& $getData['platform_type']){
            $whereSql .= "AND ORDERTYPE = '".$getData['platform_type']."'";
        }
        
        if(strlen($getData['gbillno'])>0){
            $whereSql .= "AND GENERALBILLNO LIKE '%".$getData['gbillno']."%'";
        }
        
        if(strlen($getData['start-time'])>0){
            $whereSql .= " AND TO_CHAR(REALARRIVALDATE, 'YYYY-MM-DD HH24:MI:SS') > '" . $getData['start-time'] . " 00:00:00'";
        }
        
        if(strlen($getData['end-time'])>0){
            $whereSql .= " AND TO_CHAR(REALARRIVALDATE, 'YYYY-MM-DD HH24:MI:SS') < '" . $getData['end-time'] . " 24:00:00'";
        }
        $data = $dataserviceModel->getList($whereSql, $curr_page,$this->_pageSize);
        $count = $dataserviceModel->getNowhereCount($whereSql);
        $page = new pager ($count, $curr_page, $this->_pageSize);
	    $pageStr = $page->GetPagerContent ();
        //var_dump($data);
        //var_dump($getData);
        //var_dump($whereSql);
	    $count1 = $this->_orderHeadModel->get_count(" AND HANDLESTATE='3' AND BILLSTATUS='0' ");
	    $count2 = $this->_orderHeadModel->get_count(" AND HANDLESTATE='3' AND BILLSTATUS='1' ");
	    $count3 = $this->_orderHeadModel->get_count(" AND HANDLESTATE='3' AND BILLSTATUS='2' ");
	    $count4 = $ecmsTaxModel->get_count(" AND TYPE='ECMS'");
	    $count5 = $ecmsTaxModel->get_count(" AND TYPE='UPS'");
	    $count6 = $dataserviceModel->getNowhereCount("");
	    $this->assign("count1", $count1);
	    $this->assign("count2", $count2);
	    $this->assign("count3", $count3);
	    $this->assign("count4", $count4);
	    $this->assign("count5", $count5);
	    $this->assign("count6", $count6);
        array_unshift($plamfromType, "选择平台类型");
        $this->assign('page',$pageStr);
        $this->assign('count',$count);
        $this->assign('platform_type',$getData['platform_type']);
        $this->assign('gbillno',$getData['gbillno']);
        $this->assign('starttime',$getData['start-time']);
        $this->assign('endtime',$getData['end-time']);
        $this->assign('data', $data);
        $this->assign("platformType", $plamfromType);
        $this->display('dataservicefee/dataservicefee_index.html');
    }
    
    /**
     * 
     * 订单列表导出Excel
     * 
     */
    public function actionDoExport($stamp = false,$sheet =false){
        header("Content-Type:text/html;charset=GB2312");
        Yii::$enableIncludePath=false;
        Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
        $obj_phpexcel = new PHPExcel();
        $dataserviceModel = new DataServiceFeeModel();
        $getData = $this->_get();
        $month = explode('-',$getData['start-time'])[1];
        $whereSql = '';
        $styleThinBlackBorderOutline = array(
            'borders' => array (
                'outline' => array (
                    'style' => PHPExcel_Style_Border::BORDER_THIN,   //设置border样式
                    //'style' => PHPExcel_Style_Border::BORDER_THICK,  另一种样式
                    'color' => array ('argb' => 'FF000000'),          //设置border颜色
                ),
            ),
        );
        if($getData['platform_type'] !='0'){
            $whereSql .= "AND ORDERTYPE = '".$getData['platform_type']."'";
        }
        
        if(strlen($getData['gbillno'])>0){
            $whereSql .= "AND GENERALBILLNO LIKE '%".$getData['gbillno']."%'";
        }
        
        if(strlen($getData['start-time'])>0){
            $whereSql .= " AND TO_CHAR(REALARRIVALDATE, 'YYYY-MM-DD HH24:MI:SS') > '" . $getData['start-time'] . " 00:00:00'";
        }
        
        if(strlen($getData['end-time'])>0){
            $whereSql .= " AND TO_CHAR(REALARRIVALDATE, 'YYYY-MM-DD HH24:MI:SS') < '" . $getData['end-time'] . " 24:00:00'";
        }
        $getData['platform_type'] = $getData['platform_type']!='0'?$getData['platform_type']:'全部平台';
        $data = $dataserviceModel->getNowhereList($whereSql);
        if(!$sheet){
            $operateExcel = new OperateExcel();
            $operateExcel=PHPExcel_IOFactory::load(Yii::app()->basePath.'\tmp\temp\ServiceTemplate.xlsx');
            $writer=PHPExcel_IOFactory::createWriter($operateExcel,'Excel5');
            $sheet=$operateExcel->getSheet(0);
        }
        
        $sheet->setTitle("Worksheet");
        $i = 1;
        $number = 0;
        $row = 2;
        foreach($data as $key=>$val){
            $i++;
            $row++;
            $sheet->setCellValueByColumnAndRow(0,$i,$key+1);
            $sheet->setCellValueByColumnAndRow(1,$i,$val['GENERALBILLNO']);
            $sheet->setCellValueByColumnAndRow(4,$i,$val['NUM']);
            $sheet->setCellValueByColumnAndRow(2,$i,$val['FLIGHTNO']);
            $sheet->setCellValueByColumnAndRow(3,$i,$val['REALARRIVALDATE']);
            $number += $val['NUM'];
            $sheet->getRowDimension($i)->setRowHeight(20);
            //设置边框
            $sheet->getStyle("A{$i}")->applyFromArray($styleThinBlackBorderOutline);
            $sheet->getStyle("B{$i}")->applyFromArray($styleThinBlackBorderOutline);
            $sheet->getStyle("C{$i}")->applyFromArray($styleThinBlackBorderOutline);
            $sheet->getStyle("D{$i}")->applyFromArray($styleThinBlackBorderOutline);
            $sheet->getStyle("E{$i}")->applyFromArray($styleThinBlackBorderOutline);
        }
        $sheet->getRowDimension($row)->setRowHeight(20);
        $sheet->mergeCells("A{$row}:D{$row}");
        $sheet->getCell("A{$row}")->setValue('合计');
        $sheet->getStyle("A{$row}:E{$row}")->applyFromArray($styleThinBlackBorderOutline);
        $sheet->setCellValueByColumnAndRow(4,$row,$number);
        //设置字体
        $sheet->getStyle("A{$row}:E{$row}")->getFont()->setName('宋体');
        $sheet->getStyle("A{$row}:E{$row}")->getFont()->setSize(16);
        $sheet->getStyle("A{$row}:E{$row}")->getFont()->setBold(true);
        $sheet->getStyle("A{$row}:E{$row}")->getFont()->getColor()->setRGB('0000');
        $sheet->getRowDimension($row)->setRowHeight(40);
        if(!$stamp){
            $file=uniqid().'.xls';
            $file_name = $getData['platform_type']."__".$month."月份到港数据";
            $show_name="{$file_name}.xls";
            $write=PHPExcel_IOFactory::createWriter($operateExcel,'Excel5');
            $write->save($file);
            //解决乱码
            ob_end_clean();
            //弹出下载对话框
            header('Content-Disposition: attachment; filename='.$show_name);
            header('Content-Length: ' . filesize($file));
            //生成缓存文件
            readfile($file);
            //删除缓存文件
            unlink($file);
        }else{
            return true;
        }
        
    }
    
    /**
     * 
     * 导出订单数据服务费报表
     * 
     */
    public function actionExportServiceCharge(){
        header("Content-Type:text/html;charset=GB2312");
        Yii::$enableIncludePath=false;
        Yii::import('application.extensions.PHPExcel.PHPExcel', 1);
        $obj_phpexcel = new PHPExcel();
        $dataserviceModel = new DataServiceFeeModel();
        $getData = $this->_get();
        $time = date('F/d',time());
        //echo $getData['start-time'];
        $whereSql = '';
        if($getData['platform_type'] !='0'){
            $whereSql .= "AND ORDERTYPE = '".$getData['platform_type']."'";
        }
        
        if(strlen($getData['gbillno'])>0){
            $whereSql .= "AND GENERALBILLNO LIKE '%".$getData['gbillno']."%'";
        }
        
        if(strlen($getData['start-time'])>0){
            $whereSql .= " AND TO_CHAR(REALARRIVALDATE, 'YYYY-MM-DD HH24:MI:SS') > '" . $getData['start-time'] . " 00:00:00'";
        }
        
        if(strlen($getData['end-time'])>0){
            $whereSql .= " AND TO_CHAR(REALARRIVALDATE, 'YYYY-MM-DD HH24:MI:SS') < '" . $getData['end-time'] . " 24:00:00'";
        }
        $data = $dataserviceModel->getNowhereList($whereSql);
        foreach($data as &$val){
            $val['price'] = $getData['charge'];
            $val['charge'] = $val['NUM']*$getData['charge'];
        }
        $opexcel = new OperateExcel();
        $opexcel=PHPExcel_IOFactory::load(Yii::app()->basePath.'\tmp\temp\ServiceFeeTemplate.xlsx');
        $writer=PHPExcel_IOFactory::createWriter($opexcel,'Excel5');
        $this->actionDoExport('OK',$opexcel->getSheet(1));
        $sheet1=$opexcel->getSheet(0);
        
        //设置sheet标题
        $sheet1->setTitle("Debit note");
        //设置客户资料
        $sheet1->getCell('A10')->setValue($getData['platform_type']);
        //设置日期
        $sheet1->getCell('F7')->setValue($time);
        $roll = 14;
        $styleThinBlackBorderOutline = array(
            'borders' => array (
                'outline' => array (
                    'style' => PHPExcel_Style_Border::BORDER_THIN,   //设置border样式
                    //'style' => PHPExcel_Style_Border::BORDER_THICK,  另一种样式
                    'color' => array ('argb' => 'FF000000'),          //设置border颜色
                ),
            ),
        );
        //金额
        $price = 0;
        //总数
        $countnum = 0;
        foreach($data as $res){
            $price += intval($res['charge']);
            $countnum += intval($res['NUM']);
        }
         //设置数量
        $sheet1->getCell('D15')->setValue($countnum);
        //设置单价
        $sheet1->getCell('E15')->setValue($getData['charge']);
        //设置总价
        $sheet1->getCell('F15')->setValue($countnum*$getData['charge']);
       /*  foreach($data as $key=>$res){
            $roll++;
            //合并单元格
            $sheet1->mergeCells("B{$roll}:C{$roll}");
            //设置字体
            $sheet1->getStyle("A{$roll}:F{$roll}")->getFont()->setName('宋体');
            $sheet1->getStyle("A{$roll}:F{$roll}")->getFont()->setSize(8);
            $sheet1->getStyle("A{$roll}:F{$roll}")->getFont()->setBold(true);
            $sheet1->getStyle("A{$roll}:F{$roll}")->getFont()->getColor()->setRGB('0000');
            //序号
            $sheet1->setCellValueByColumnAndRow(0,$roll,$key+1);
            //收费项目
            $sheet1->setCellValueByColumnAndRow(1,$roll,'数据处理费');
            //数量
            $sheet1->setCellValueByColumnAndRow(3,$roll,$res['NUM']);
            //单价
            $sheet1->setCellValueByColumnAndRow(4,$roll,'￥'.number_format($res['price'],2));
            //金额
            $sheet1->setCellValueByColumnAndRow(5,$roll,'￥'.number_format($res['charge'],2));
            $price += intval($res['charge']);
            $sheet1->setCellValueByColumnAndRow(5,$roll+1,$price);
            //设置边框
            $sheet1->getStyle("A{$roll}")->applyFromArray($styleThinBlackBorderOutline);
            $sheet1->getStyle("B{$roll}:C{$roll}")->applyFromArray($styleThinBlackBorderOutline);
            $sheet1->getStyle("D{$roll}")->applyFromArray($styleThinBlackBorderOutline);
            $sheet1->getStyle("E{$roll}")->applyFromArray($styleThinBlackBorderOutline);
            $sheet1->getStyle("F{$roll}")->applyFromArray($styleThinBlackBorderOutline);
            //生成一行
            $sheet1->insertNewRowBefore($roll+1,1);
        }
            
            //删除一行
            $sheet1->removeRow($roll+1,1); */
        
            $filename=uniqid().'.xls';
            $file_name = "顺益订单服务费(".$getData['platform_type']."_".$getData['start-time']."_".$getData['end-time'].")";
            $show_name="{$file_name}.xls";
            $write1=PHPExcel_IOFactory::createWriter($opexcel,'Excel5');
            $write1->save($filename);
            //解决乱码
            ob_end_clean();
            //弹出下载对话框
            header('Content-Disposition: attachment; filename='.$show_name);
            header('Content-Length: ' . filesize($filename));
            //生成缓存文件
            readfile($filename);
            //删除缓存文件
            unlink($filename);
    }
    
}

```
