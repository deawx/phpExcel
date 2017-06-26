<?php
require_once __DIR__ . '/vendor/autoload.php';



//	Change these values to select the Rendering library that you wish to use
//		and its directory location on your server
//$rendererName = PHPExcel_Settings::PDF_RENDERER_TCPDF;
$rendererName = PHPExcel_Settings::PDF_RENDERER_MPDF;
// $rendererName = PHPExcel_Settings::PDF_RENDERER_DOMPDF;
//$rendererLibrary = 'tcPDF5.9';
$rendererLibrary = 'mpdf61';
// $rendererLibrary = 'domPDF0.6.0beta3';
// $rendererLibraryPath = '/php/libraries/PDF/' . $rendererLibrary;
$rendererLibraryPath = $rendererLibrary;

// Create new PHPExcel object
$objPHPExcel = new PHPExcel();

// Set document properties
$objPHPExcel->getProperties()->setCreator("Maarten Balliauw")
							 ->setLastModifiedBy("Maarten Balliauw")
							 ->setTitle("PDF Test Document")
							 ->setSubject("PDF Test Document")
							 ->setDescription("Test document for PDF, generated using PHP classes.")
							 ->setKeywords("pdf php")
							 ->setCategory("Test result file");

$objReader = new PHPExcel_Reader_HTML();

$objPHPExcel = $objReader->load("test.php");

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

$objWriter = new PHPExcel_Writer_PDF($objPHPExcel);
$objWriter->save("htmltopdf0.pdf");
?>