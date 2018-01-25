<!DOCTYPE html>
<html>
<head>
	<title>测试excel</title>
	<style type="text/css">
		a{color: black;}
		td{text-align: center;}
	</style>
</head>
<body>
	 <a href="index.php?id=1">导出</a>


		
</body>
</html>
<?php 
 	date_default_timezone_set('PRC');
   	$data=array();
   	for ($i=1; $i <51 ; $i++) { 
   		$data[$i]['keyword']='keyword'.$i;
   		$data[$i]['read_ip']='read_ip'.$i;
   		$data[$i]['down_ip']='down_ip'.$i;
   		$data[$i]['%%%']='%%%'.$i;
   		$data[$i]['chaneel']='chaneel'.$i;

   	}
    if (!empty($_GET['id'])) {
    	include 'excel/PHPExcel.class.php';
    	$objPHPExcel = new \PHPExcel();
	
		 
		// 表头
		$objPHPExcel->setActiveSheetIndex(0)
		->setCellValue('A1', 'ID')
		->setCellValue('B1', '关键词')
		->setCellValue('C1', '访问IP数')
		->setCellValue('D1', '下载IP数')
		->setCellValue('E1', '百分比')
		->setCellValue('F1', '渠道');
	    
		// 内容
		for ($i = 1, $len = count($data); $i < $len; $i++) {
			$objPHPExcel->getActiveSheet(0)->setCellValue('A' . ($i+2), $i);
			$objPHPExcel->getActiveSheet(0)->setCellValue('B' . ($i+2), $data[$i]['keyword']);
			$objPHPExcel->getActiveSheet(0)->setCellValue('C' . ($i+2), $data[$i]['read_ip']);
			$objPHPExcel->getActiveSheet(0)->setCellValue('D' . ($i+2), $data[$i]['down_ip']);
			$objPHPExcel->getActiveSheet(0)->setCellValue('E' . ($i+2), $data[$i]['%%%']);
			$objPHPExcel->getActiveSheet(0)->setCellValue('F' . ($i+2), $data[$i]['chaneel']);
		}
	
		// Rename sheet
		$objPHPExcel->getActiveSheet()->setTitle('关键词');
	
		// Set active sheet index to the first sheet, so Excel opens this as the first sheet
		$objPHPExcel->setActiveSheetIndex(0);
	
		// 输出
		//header('Content-Type: application/vnd.ms-excel');
		//header('Content-Disposition: attachment;filename="' .'关键词导出'. date("YmdHi") . '.xls"');
		//header('Cache-Control: max-age=0');
		ob_end_clean();//清除缓冲区,避免乱码
		 
		header("Content-type:application/octet-stream");
		header("Accept-Ranges:bytes");
		header("Content-type:applicationnd.ms-excel");
		header("Content-Disposition:attachment;filename=".'关键词导出'.date('YmdH').".xls");
		header("Pragma: no-cache");
		header("Expires: 0");
		 
		$objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
		 
		$objWriter->save('php://output');
    }
   //如果有数据库操作，那就是导入之后再添加到数据库
   /*insert data to databse*/	
  	function into_database($filename){
				header("content-type:text/html;charset=utf8");
				set_time_limit(0);
				include 'excel/PHPExcel.class.php';
	
				$objPHPExcel    = \PHPExcel_IOFactory::load($filename);
				$objWorksheet   = $objPHPExcel->getActiveSheet();
				$objActSheetArr = $objWorksheet->toArray('',false,true,false);
	
				//dump($objActSheetArr);exit;
				
	
				for($i = 1; $i < count($objActSheetArr); $i++ ){
	
				$data['keyword'] = $objActSheetArr[$i][0];//关键词
				$data['price']   = $objActSheetArr[$i][1];//均价
				$data['type']    = $objActSheetArr[$i][2];//渠道
				$data['ctime']   = time();//时间
				/*insert into table_name value $data*/
				}	
	}					
		

?>
