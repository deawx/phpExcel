<?php 
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

	
 ?>