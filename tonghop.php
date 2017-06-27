<?php 
//data -> excel ok
//data -> html ok
//data -> pdf ok
//excel -> html theo dinh dang table ok
//excel -> pdf ok
//html -> excel phai tao dinh dang table fail
//html -> pdf ok
	require_once __DIR__ . '/vendor/autoload.php';

	$objPHPExcel = new PHPExcel();

	//bảo mật bảng tính
	$objPHPExcel->getActiveSheet()->getProtection()->setSheet(true);

	//cài đặt font chữ mặc định
	$objPHPExcel->getActiveSheet()->getDefaultStyle()->getFont()->setName('Arial Cyr');

	//cấu hình file nhằm đảm bảo tính sở hữu cá nhân hoặc bản quyền hoặc quản lý
	//được dễ dàng hơn
	$objPHPExcel->getProperties()->setCreator('Gia Thanh Phat')
								 ->setLastModifiedBy('Gia Thanh Phat')
								 ->setTitle('Export PDF by PHPExcel')
								 ->setSubject('Export PDF by PHPExcel')
								 ->setDescription('Convert HTML to PDF by PHPExcel library')
								 ->setKeywords('PHPExcel')
								 ->setCategory('html to pdf')
								 ->setCompany('IVC VietNam');

	//đặt giá trị vào ô bằng tọa độ
	$objPHPExcel->getActiveSheet()->setCellValue('A2', 'giathanhphat');

	//lấy giá trị của một ô dữ liệu
	$objPHPExcel->getActiveSheet()->getCell('A2')->getValue();

	//lấy phép tính của ô có giá trị đó. Ví dụ ô 'A2' có giá trị 6 được thực hiện
	//bởi phép tính 'B2' * 2 ( thì đây là kết quả lấy được).
	$objPHPExcel->getActiveSheet()->getCell('A2')->getCalculatedValue();

	//thiết lập giá trị theo cột và hàng
	$objPHPExcel->getActiveSheet()->setCellValueByColumnAndRow('1', '8', 'le thi phuong vy');

	//lấy giá trị theo cột và hàng
	$objPHPExcel->getActiveSheet()->getCellByColumnAndRow('1', '8')->getValue();

	//lấy phép tính của ô có giá trị đó
	$objPHPExcel->getActiveSheet()->getCellByColumnAndRow('1', '8')->getCalculatedValue();

	//ví dụ đọc tất cả các giá trị trong bảng tính (file excel) để hiển thị ra bảng
	//tạo đối tượng để đọc file (ở đây đọc file excel 2007)
	$objReader = PHPExcel_IOFactory::createReader('Excel2007');
	//gán thuộc tính chỉ để đọc
	$objReader->setReadDataOnly(true);
	//load file (kèm đọc đã đọc dữ liệu)
	$objPHPExcel = $objReader->load('test.xlsx');
	//mở bảng tính và sử dụng
	$objWorksheet = $objPHPExcel->getActiveSheet();
	//lấy vị trí dòng chứa dữ liệu cuối cùng.
	$HighesRow = $objWorksheet->getHighestRow();//ví dụ 10
	//lấy vị trí cột chứa dữ liệu cuối cùng
	$HighesColumn = $objWorksheet->getHighestColumn();//vị trí 'F'
	//chuyển vị trí cột thành chuỗi
	$HighesColumnIndex = PHPExcel_Cell::columnIndexFromString($HighesColumn);//vị trí 5 <=> vị trí 'F'
	echo '<table>'."\N";
	For ($row = 0; $row <= $HighestRow; $row++) 
	{
		echo '<tr>'."\N";
		For ($col = 0; $col <= $HighestColumnIndex; $col++) 
		{
			echo '<td>'. $ObjWorksheet->getCellByColumnAndRow ($col, $row)->getValue().'</td>'."\N";
		}
		echo '</tr>'."\N";
	}
	echo '</table>'."\N";

	//tạo kết dính khi nhập dữ liệu có nghĩa là dữ liệu sẽ được chuyển sang dạng chung
	//mà nó đã được định nghĩa sẵn ví dụ nhập 10% -> 0.1, 21 December 1983 -> 1983-12-21,
	//định dạng giờ là H:i:s
	//nhớ thêm 3 file
	require_once 'PHPExcel.php'; 
	require_once 'PHPExcel/Cell/AdvancedValueBinder.php'; 
	require_once 'PHPExcel/IOFactory.php';
	// Set value binder 
	PHPExcel_Cell::setValueBinder( new PHPExcel_Cell_AdvancedValueBinder() );
	//tạo lại phpexcel
	$objPHPExcel = new PHPExcel();
	// Add some data, resembling some different data types
	$objPHPExcel->getActiveSheet()->setCellValue('A4', 'Percentage value:');
	$objPHPExcel->getActiveSheet()->setCellValue('B4', '10%');// Converts to 0.1 and sets percentage cell style
	$objPHPExcel->getActiveSheet()->setCellValue('A5', 'Date/time value:');
	$objPHPExcel->getActiveSheet()->setCellValue('B5', '21 December 1983');// Converts to date and sets date format cell style

	//thiết lập một kiểu dữ liệu cho ô
	 $objPHPExcel-> getActiveSheet () -> getCell ('A1') -> setValueExplicit ('25 ', PHPExcel_Cell_DataType :: TYPE_NUMERIC);

	 //nhập url vào ô
	 $objPHPExcel->getActiveSheet()->getCell('E26')->getHyperlink()->setUrl('http://www.phpexcel.net');

	 //thiết lập hướng trang và kích thước trang
	$objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE);
	$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);

	//căn giữa một trang theo chiều ngang hoặc chiều dọc
	$objPHPExcel->getActiveSheet()->getHeaderFooter()->setOddHeader('& C & HPlease coi tài liệu này là bí mật!');
	$objPHPExcel->getActiveSheet()->getHeaderFooter()->setOddFooter('& L & B'. $objPHPExcel->getProperties()->getTitle(). '& RPage & P of & N');

	//thêm hình ảnh cho header và footer
	$objDrawing = new PHPExcel_Worksheet_HeaderFooterDrawing();
	$objDrawing->setName('PHPExcel logo');
	$objDrawing->setPath('./images/phpexcel_logo.gif');
	$objDrawing->setHeight(36);
	$objPHPExcel->getActiveSheet()->getHeaderFooter()->addImage($objDrawing, PHPExcel_Worksheet_HeaderFooter::IMAGE_HEADER_LEFT);
	
	//thiết lập ngắt in trên một dòng
	$objPHPExcel->getActiveSheet()->setBreak( 'A10' , PHPExcel_Worksheet::BREAK_ROW );

	//thiết lập ngắt in trên một cột
	$objPHPExcel->getActiveSheet()->setBreak( 'D10' , PHPExcel_Worksheet::BREAK_COLUMN );

	//ẩn / hiện lưới khi in (dùng nhiều)
	$objPHPExcel->getActiveSheet()->setShowGridlines(true);

	//thiết lập hàng/cột để lặp lại ở trên cùng/bên trái
	$objPHPExcel->getActiveSheet()->getPageSetup()->setRowsToRepeatAtTopByStartAndEnd(1, 5);

	//chỉ định vùng in
	$objPHPExcel->getActiveSheet()->getPageSetup()->setPrintArea('A1:E5');

	//định dạng thuộc tính ô: như font, color, border...
	$objPHPExcel->getActiveSheet()->getStyle('B2')->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_RED);

	$objPHPExcel->getActiveSheet()->getStyle('B2')->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_RIGHT);

	$objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
	$objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
	$objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);
	$objPHPExcel->getActiveSheet()->getStyle('B2')->getBorders()->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THICK);

	$objPHPExcel->getActiveSheet()->getStyle('B2')->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
	$objPHPExcel->getActiveSheet()->getStyle('B2')->getFill()->getStartColor()->setARGB('FFFF0000');

	//định dạng một dãy ô
	$objPHPExcel->getActiveSheet()->getStyle('B3:B7')->getFill()
	->setFillType(PHPExcel_Style_Fill::FILL_SOLID)
	->getStartColor()->setARGB('FFFF0000');

	//định dạng ô bằng cách thiết lập mảng style
	$styleArray = array(
	'font' => array(
	'bold' => true,
	),
	'alignment' => array(
	'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_RIGHT,
	),
	'borders' => array(
	'top' => array(
	'style' => PHPExcel_Style_Border::BORDER_THIN,
	),
	),
	'fill' => array(
	'type' => PHPExcel_Style_Fill::FILL_GRADIENT_LINEAR,
	'rotation' => 90,
	'startcolor' => array(
	'argb' => 'FFA0A0A0',
	),
	'endcolor' => array(
	'argb' => 'FFFFFFFF',
	),
	),
	);
	$objPHPExcel->getActiveSheet()->getStyle('A3')->applyFromArray($styleArray);
	//hoặc một dãy ô
	$objPHPExcel->getActiveSheet()->getStyle('B3:B7')->applyFromArray($styleArray);

	//thiết lập mặc định bảng tính

	//điều kiện định dạng ô

	//đưa chú thích vào một ô

	//thiết lập bảo mật trên một bảng tính

	//thiết lập xác thực dữ liệu trên một ô

	//đặt chiều rộng cho cột (thủ công hoặc auto)
	$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(12);
	$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);

	//hiện/ẩn một cột

	//thiết lập độ cao của hàng

	//show/hide một hàng

	//gộp nhiều ô thành 1 ô
	$objPHPExcel->getActiveSheet()->mergeCells('A18:E22');

	//tách 1 ô thành nhiều ô
	$objPHPExcel->getActiveSheet()->unmergeCells('A18:E22');

	//chèn hàng
	$objPHPExcel->getActiveSheet()->insertNewRowBefore(7,2); // chèn 2 hàng mới ngay trước hàng thứ 7

	//chuyển đầu ra tới trình duyệt web -------------------------bắt buộc--------------------------------
	//ví dụ chuyển tệp excel2007 ra trình duyệt
	/* Here there will be some code where you create $objPHPExcel */
	// redirect output to client browser
	header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
	header('Content-Disposition: attachment;filename="myfile.xlsx"');
	header('Cache-Control: max-age=0');

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	$objWriter->save('php://output');

	//ví dụ chuyển tệp excel5 ra trình duyệt
	/* Here there will be some code where you create $objPHPExcel */
	// redirect output to client browser
	header('Content-Type: application/vnd.ms-excel');
	header('Content-Disposition: attachment;filename="myfile.xls"');
	header('Cache-Control: max-age=0');

	$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel5');
	$objWriter->save('php://output');

	//thiết lập chiều rộng cột mặc định
	$objPHPExcel->getActiveSheet()->getDefaultColumnDimension()->setWidth(12);

	//thiết lập chiều cao hàng mặc định
	$objPHPExcel->getActiveSheet()->getDefaultRowDimension()->setRowHeight(15);

	//thiết lập mức thu nhỏ/ phóng to của bảng tính
	$objPHPExcel->getActiveSheet()->getSheetView()->setZoomScale(75);

	//thiết lập tất cả các reader để đọc file
	$objReader = new PHPExcel_Reader_Excel2007();
	$objPHPExcel = $objReader->load("05featuredemo.xlsx");
	$objReader = new PHPExcel_Reader_Excel5();
	$objPHPExcel = $objReader->load("05featuredemo.xls");
	$objReader = new PHPExcel_Reader_CSV();
	$objPHPExcel = $objReader->load("05featuredemo.csv");
	$objReader = new PHPExcel_Reader_HTML();
	$objPHPExcel = $objReader->load("test.php");

	//thiết lập tất cả các writer để xuất file
	$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
	$objWriter->save("05featuredemo.xlsx");
	$objWriter = new PHPExcel_Writer_Excel5($objPHPExcel);
	$objWriter->save("05featuredemo.xls");
	$objWriter = new PHPExcel_Writer_CSV($objPHPExcel);
	$objWriter->save("05featuredemo.csv");
	$objWriter = new PHPExcel_Writer_HTML($objPHPExcel);
	$objWriter->save("05featuredemo.html");
	$objWriter = new PHPExcel_Writer_PDF($objPHPExcel);
	$objWriter->save("05featuredemo.pdf");

	//tổng hợp style định dạng
 ?>