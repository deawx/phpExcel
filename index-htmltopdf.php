<?php
require_once __DIR__ . '/vendor/autoload.php';



//	Change these values to select the Rendering library that you wish to use
//		and its directory location on your server
//$rendererName = PHPExcel_Settings::PDF_RENDERER_TCPDF;
$rendererName = PHPExcel_Settings::PDF_RENDERER_MPDF;
// // $rendererName = PHPExcel_Settings::PDF_RENDERER_DOMPDF;
// //$rendererLibrary = 'tcPDF5.9';
$rendererLibrary = 'mpdf61';
// // $rendererLibrary = 'domPDF0.6.0beta3';
// // $rendererLibraryPath = '/php/libraries/PDF/' . $rendererLibrary;
$rendererLibraryPath = $rendererLibrary;

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Set document properties
$objPHPExcel->getProperties()->setCreator("Gia Thanh Phat")
							 ->setLastModifiedBy("Gia Thanh Phat")
							 ->setTitle("Convert PDF")
							 ->setSubject("PDF Test Document")
							 ->setDescription("Convert PDF using PHPExcel")
							 ->setKeywords("pdf phpexcel")
							 ->setCategory("Test result file");


$objReader = new PHPExcel_Reader_HTML();
$objPHPExcel = $objReader->load("test.php");
$objPHPExcel->getActiveSheet()->setCellValue('A12', 'Title');
$objPHPExcel->getActiveSheet()->getHeaderFooter()->setOddHeader('& C & H Please coi tài liệu này là bí mật!');
$objPHPExcel->getDefaultStyle()->getFont()->setName('Arial');
$objPHPExcel->getDefaultStyle()->getFont()->setSize(20);
$objPHPExcel->getDefaultStyle()->getFont()->getColor()->setRGB('ff0000');
// $objPHPExcel->getActiveSheet()->getPageSetup()->setOrientation(PHPExcel_Worksheet_PageSetup::ORIENTATION_LANDSCAPE); //theo chiều rộng
$objPHPExcel->getActiveSheet()->getPageSetup()->setPaperSize(PHPExcel_Worksheet_PageSetup::PAPERSIZE_A4);
$objPHPExcel->getActiveSheet()->getStyle('A1')->getFont()->getColor()->setRGB('a1a1a1');
$objPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getTop()->setBorderStyle('dashed');
$objPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getRight()->setBorderStyle('dashed');
$objPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getBottom()->setBorderStyle('dashed');
$objPHPExcel->getActiveSheet()->getStyle('A1')->getBorders()->getLeft()->setBorderStyle('dashed');

// $objPHPExcel->getActiveSheet()->getHeaderFooter()->setOddFooter('giathanhphat');
if (!PHPExcel_Settings::setPdfRenderer(
		$rendererName,
		$rendererLibraryPath
	)) {
	die(
		'NOTICE: Please set the $rendererName and $rendererLibraryPath values' .
		'<br />' .
		'at the top of this script as appropriate for your directory structure'
	);
}

header('Content-Type: application/pdf');
header('Content-Disposition: attachment; filename="hahahoho.pdf"');
header('Cache-Control: max-age=0');

	$objWriter = new PHPExcel_Writer_PDF($objPHPExcel);
	$objWriter->save('php://output');
?>