<?php
    function request_post($url = '', $param = '')
    {
        if (empty($url) || empty($param)) {
            return false;
        }

        $postUrl = $url;
        $curlPost = $param;
        // 初始化curl
        $curl = curl_init();
        curl_setopt($curl, CURLOPT_URL, $postUrl);
        curl_setopt($curl, CURLOPT_HEADER, 0);
        // 要求结果为字符串且输出到屏幕上
        curl_setopt($curl, CURLOPT_RETURNTRANSFER, 1);
        curl_setopt($curl, CURLOPT_SSL_VERIFYPEER, false);
        // post提交方式
        curl_setopt($curl, CURLOPT_POST, 1);
        curl_setopt($curl, CURLOPT_POSTFIELDS, $curlPost);
        // 运行curl
        $data = curl_exec($curl);
        curl_close($curl);

        return $data;
    }

    function createexcel($data1 , $data2){
        error_reporting(0);
        $name = '数据抓取_'.date('Y-m-d').'.xls';
        include_once './PHPExcel/Classes/PHPExcel.php';
        // Create new PHPExcel object
        $objPHPExcel = new PHPExcel();

    // Set document properties
        $objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
            ->setLastModifiedBy("Maarten Balliauw")
            ->setTitle("Office 2007 XLSX Test Document")
            ->setSubject("Office 2007 XLSX Test Document")
            ->setDescription("Test document for Office 2007 XLSX, generated using PHP classes.")
            ->setKeywords("office 2007 openxml php")
            ->setCategory("Test result file");


    // Add some data
        $objPHPExcel->setActiveSheetIndex(0)
            ->setCellValue('A1', 'showName')
            ->setCellValue('B1', '年份')
            ->setCellValue('C1', '数量')
            ->setCellValue('E1', '学科分类')
            ->setCellValue('F1', '数量');
        //加粗居中
        $objPHPExcel->getActiveSheet()->getStyle('A1:C1')->applyFromArray(
            array(
                'font' => array (
                    'bold' => true
                ),
                'alignment' => array(
                    'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER
                )
            )
        );
        if ($data1 && $data2){
            foreach ($data1 as $k => $v){
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('A'.($k+2), $v['showName'])
                    ->setCellValue('B'.($k+2), $v['value'])
                    ->setCellValue('C'.($k+2), $v['count']);
            }

            foreach ($data2 as $k => $v){
                $objPHPExcel->setActiveSheetIndex(0)
                    ->setCellValue('E'.($k+2), $v['showName'])
                    ->setCellValue('F'.($k+2), $v['count']);
            }
        }
    // Rename worksheet
        $objPHPExcel->getActiveSheet()->setTitle($name);


    // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $objPHPExcel->setActiveSheetIndex(0);


    // Redirect output to a client’s web browser (Excel5)
        header('Content-Type: application/vnd.ms-excel');
        header('Content-Disposition: attachment;filename="'.$name.'"');
        header('Cache-Control: max-age=0');
    // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

    // If you're serving to IE over SSL, then the following may be needed
        header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
        header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header ('Pragma: public'); // HTTP/1.0

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
        $objWriter->save('php://output');
        exit;
}

    $param['searchType'] = 'all';
    //$param['searchWord'] = '(%E4%BD%9C%E8%80%85%E5%8D%95%E4%BD%8D%3A(%22%E5%AE%89%E9%98%B3%22)*%E4%BD%9C%E8%80%85%E5%8D%95%E4%BD%8D%3A(%22%E5%8D%97%E4%BA%AC%22))';
    $param['searchWord'] = $_POST['searchWord'];
    $param['facetField'] = '$common_year';
    $param['isHit'] = '';
    $param['startYear'] = '';
    $param['endYear'] = '';
    $param['limit'] = '100';
    $param['single'] = 'true';

    $url = 'http://www.wanfangdata.com.cn/search/navigation.do';

    $result1 = request_post($url , $param);
    $data1 = json_decode($result1 , true);
    $data1 = $data1['facetTree'];

    $param2['searchType'] = 'all';
    //$param['searchWord'] = '(%E4%BD%9C%E8%80%85%E5%8D%95%E4%BD%8D%3A(%22%E5%AE%89%E9%98%B3%22)*%E4%BD%9C%E8%80%85%E5%8D%95%E4%BD%8D%3A(%22%E5%8D%97%E4%BA%AC%22))';
    $param2['searchWord'] = $_POST['searchWord'];
    $param2['facetField'] = '$subject_classcode_level';
    $param2['isHit'] = '';
    $param2['startYear'] = '';
    $param2['endYear'] = '';
    $param2['limit'] = '100';
    $param2['single'] = 'true';
    $result2 = request_post($url , $param2);
    $data2 = json_decode($result2 , true);
    $data2 = $data2['facetTree'];
    //var_dump($data2);
    createexcel($data1 , $data2);
    //var_dump($data);
    die;