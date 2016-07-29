<?php
include_once("xlsxwriter.class.php");
ini_set('display_errors', 0);
ini_set('log_errors', 1);
error_reporting(E_ALL & ~E_NOTICE);

$filename = "example.xlsx";
header('Content-disposition: attachment; filename="'.XLSXWriter::sanitize_filename($filename).'"');
header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
header('Content-Transfer-Encoding: binary');
header('Cache-Control: must-revalidate');
header('Pragma: public');

$header = array(
    'year'=>'string',
    'month'=>'string',
    'amount'=>'money',
    'first_event'=>'datetime',
    'second_event'=>'date',
	'url' => 'string'
);
$data1 = array(
    array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00','Visit Google'),
    array('2003','=B2', '23.5','2010-01-01 00:00:00','2012-12-31 00:00:00','Visit Github Profile'),
);
$data2 = array(
    array('2003','01','343.12'),
    array('2003','02','345.12'),
);

$hyperlinks = array(
	'F2' => array( 'name' => 'Visit Google', 'link' => 'https://www.google.com/' ),
	'F3' => array( 'name' => 'My Profile', 'link' => 'https://mayank.github.io/' )
);
$writer = new XLSXWriter();
$writer->setAuthor('Some Author');
$writer->writeSheet($data1,'Sheet1',$header);
$writer->writeSheet($data2,'Sheet2');
$writer->writeToStdOut();
//$writer->writeToFile('example.xlsx');
//echo $writer->writeToString();
exit(0);


