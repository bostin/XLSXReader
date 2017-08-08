<?php
include_once './src/XLSXWorksheet.php';
include_once './src/XLSXReader.php';

use Bostin\Office\Excel\XLSXReader;

date_default_timezone_set('Asia/Shanghai');

$s =  memory_get_usage(true) / 1024 / 1024;

/**
 * 加载文件
 */
$xlsx = new XLSXReader('./sample.xlsx');

/**
 * 获取sheet名称数组
 */
$sheetNames = $xlsx->getSheetNames();


/**
 * 根据sheet名称来获取每个sheet的数据
 */
foreach($sheetNames as $sheetName) {
    /**
     * 根据sheet名称来获取到sheet对象
     */
	$sheet = $xlsx->getSheet($sheetName);

    /**
     * 获取sheet对象的数据
     */
	var_dump($sheet->getData());
}


$data = array_map(function($row) {
	$converted = XLSXReader::toUnixTimeStamp($row[0]);
	return array($row[0], $converted, date('c', $converted), $row[1]);
}, $xlsx->getSheetData('Dates'));
array_unshift($data, array('Excel Date', 'Unix Timestamp', 'Formatted Date', 'Data'));


$e = memory_get_usage(true) / 1024 / 1024;

var_dump($data, $s, $e);
exit;
