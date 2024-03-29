<?php
// You can run this script from bash: php export_advanced.php
// This script contains various library features such as:
// - copying sheets
// - renaming sheets
// - copying styles
// - formulas
// - raw cell values
// - custom strings (color, size)
// But no comments

require_once '../library/PclZip.php';
require_once '../library/Workbook.php';
require_once '../library/Worksheet.php';
require_once '../library/SharedStrings.php';
require_once '../library/Formula.php';

use ExcelExport\Workbook;
use ExcelExport\Formula;
use ExcelExport\SharedStrings;

$workbook = new Workbook(__DIR__ . '/templates/template.xlsx');
$workbook->extract();

$workbook->renameWorksheet(1, '03.04.2005');
$workbook->renameWorksheet(2, 'New name 2');
$workbook->cloneWorksheet(1, '01.02.2003');
$workbook->cloneWorksheet(1, '4');
$workbook->cloneWorksheet(1, '5');
$workbook->cloneWorksheet(1, '6');
$workbook->cloneWorksheet(1, '7');
$workbook->cloneWorksheet(1, '8');
$workbook->cloneWorksheet(1, '9');
$workbook->cloneWorksheet(1, '10');
$workbook->cloneWorksheet(1, '11');

$sheet = $workbook->getWorksheet(11);

$style1 = $sheet->getStyleIndex('M4');
$style2 = $sheet->getStyleIndex('N4');
$style3 = $sheet->getStyleIndex('O4');
$style4 = $sheet->getStyleIndex('P4');
$style5 = $sheet->getStyleIndex('Q4');
$style6 = $sheet->getStyleIndex('R4');
$style7 = $sheet->getStyleIndex('S4');

// Delete cells
$sheet->setCellValue('M4', null);
$sheet->setCellValue('N4', null);
$sheet->setCellValue('O4', null);
$sheet->setCellValue('P4', null);
$sheet->setCellValue('Q4', null);
$sheet->setCellValue('R4', null);
$sheet->setCellValue('S4', null);


$sheet->setCellValue('C2', "Тест", $style1);


$sum = new Formula('SUM(H7:H1006)');
$sheet->setCellValue('C3', $sum, $style3);

$sheet->setCellValue('C10', "test", $style4);
$sheet->setCellValue('C11', "test", $style6);

$sheet->saveCells();

$sheet1 = $workbook->getWorksheet(1);
$sheet1->setCellValue('C11', 3.1456, $style4);
$sheet1->saveCells();

$sheet->initRowsInserting(7, ['B', 'C', 'D', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'R']);

$styleMap = [0 => $style1, 1 => $style2, 2 => $style3, 3 => $style4, 4 => $style5];

for ($i = 0; $i < 1000; $i++) {
  $f = new Formula('J' . ($i + 7) . '+M' . ($i + 7));
  $v2 = rand(0, 4);
  $s2 = $styleMap[$v2];
  $sheet->insertRow([
    $v2,
    '    <r><t>test</t></r>
         <r><rPr><sz val="11"/><color rgb="FF0000"/></rPr>
         <t xml:space="preserve"> ' . rand(0, 5000) . '</t></r>',
    3,
    $f,
    "test string",
    6234.232,
    7,
    8,
    rand(0, 5000),
    rand(0, 5000),
    SharedStrings::customString("Common text\n") .
    SharedStrings::customString("Red text ", 'FF0000') .
    SharedStrings::customString("Green text", '22AA00', 8),
    'Sample text ' . rand(0, 5000),
    'Sample text ' . rand(0, 5000)
  ],
    [$s2, $style2, $style3, $style4, $style5, $style7, null, $style6, $style5]
  );
}
$sheet->finishRowsInserting();

// Sheet2
$sheet2 = $workbook->getWorksheet(2);
$sheet2->initRowsInserting(1, ['A', 'B', 'C', 'D', 'E']);
$sheet2->insertRow([new Formula('SUM(B1:D1)'), rand(0, 10), rand(0, 10), rand(0, 10), substr(str_shuffle(str_repeat($x = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ', ceil(20 / strlen($x)))), 1, 20)]);
$sheet2->insertRow(['', '', '', '', ''], [$s2, $style2, $style3, $style4, $style5]);
for ($i = 1; $i < 1000; $i++) {
  $sheet2->insertRow([new Formula('SUM(B' . ($i + 2) . ':D' . ($i + 2) . ')'), rand(0, 10), rand(0, 10), rand(0, 10), substr(str_shuffle(str_repeat($x = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ', ceil(20 / strlen($x)))), 1, 20)]);
}
$sheet2->finishRowsInserting();

$exportPath = __DIR__ . '/exported_files/output.xlsx';

$workbook->createXlsx($exportPath);
$workbook->clean();
echo 'File generated: ' . $exportPath . PHP_EOL;
