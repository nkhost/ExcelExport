<?php
require_once './vendor/exportella/PclZip.php';
require_once './vendor/exportella/Workbook.php';
require_once './vendor/exportella/Worksheet.php';
require_once './vendor/exportella/SharedStrings.php';
require_once './vendor/exportella/Formula.php';

use Exportella\Workbook;
use Exportella\Formula;

//$workbook = new Workbook(__DIR__ . '/templates/template.xlsx', __DIR__ . '/data');
$workbook = new Workbook(__DIR__ . '/templates/template.xlsx');
$workbook->extract();

$sheet = $workbook->getWorksheet(1);

// Сохраняем номера стилей
$style1 = $sheet->getStyleIndex('M4');
$style2 = $sheet->getStyleIndex('N4');
$style3 = $sheet->getStyleIndex('O4');
$style4 = $sheet->getStyleIndex('P4');
$style5 = $sheet->getStyleIndex('Q4');
$style6 = $sheet->getStyleIndex('R4');
$style7 = $sheet->getStyleIndex('S4');

// удаляем ячейки со стилями
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

$sheet->setCellValue('C10', "Тест", $style4);
$sheet->setCellValue('C11', "Тест", $style6);

$sheet->saveCells();


$sheet->initRowsInserting(7, ['B', 'C', 'D', 'F', 'G', 'H', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'R']);

$styleMap = [0 => $style1, 1 => $style2, 2 => $style3, 3 => $style4, 4 => $style5];

for ($i = 0; $i < 10000; $i++) {
  $f = new Formula('J' . ($i + 7) . '+M' . ($i + 7));
  $v2 = rand(0, 4);
  $s2 = $styleMap[$v2];
  $sheet->insertRow([
    $v2,
    'привет' . rand(0, 5000),
    3,
    $f,
    "sdfs",
    6234.232,
    7,
    8,
    rand(0, 5000),
    rand(0, 5000),
    'Что то ещё там ' . rand(0, 5000),
    'Что то ещё там ' . rand(0, 5000),
    'Что то ещё там ' . rand(0, 5000)
  ],
    [$s2, $style2, $style3, $style4, $style5, $style7, null, $style6, $style5]
  );
}
$sheet->finishRowsInserting();

// Лист 2
$sheet2 = $workbook->getWorksheet(2);
$sheet2->initRowsInserting(1, ['A', 'B', 'C', 'D', 'E']);
for ($i = 1; $i < 100000; $i++) {
  $sheet2->insertRow([new Formula('SUM(B' . $i . ':D' . $i . ')'), rand(0, 10), rand(0, 10), rand(0, 10), substr(str_shuffle(str_repeat($x = '0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ', ceil(20 / strlen($x)))), 1, 20)]);
}
$sheet2->finishRowsInserting();

$workbook->createXlsx(__DIR__ . '/data/test.xlsx');
$workbook->clean();
