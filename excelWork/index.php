<?php

//ini_set('max_execution_time', 600);

use OpenExcel\ParseExcel;

require_once __DIR__ . '/ParseExcel.php';

$excel_file = 'F:/strateg/Лайно/result_excel/база вся.xlsx';
$parse_excel = new ParseExcel($excel_file);

$parameters = [
    'TDSheet',
    'A',
    'F:/strateg/Лайно/result_excel/',
];

try {
    $parse_excel->createExcelFromString($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
    echo $e->getMessage() . '<br/>';

    try {
        $parse_excel->__destruct();
    } catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
        echo $e->getMessage() . '<br/>';
    }
}


