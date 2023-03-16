<?php
require_once './vendor/exportella/PclZip.php';
require_once './vendor/exportella/Workbook.php';
require_once './vendor/exportella/Worksheet.php';
require_once './vendor/exportella/SharedStrings.php';
require_once './vendor/exportella/Formula.php';

use Exportella\Workbook;
use Exportella\Formula;
use Exportella\SharedStrings;

$linesAmount = 100000;

$dataList = [];

for ($i = 0; $i < $linesAmount; $i++) {
  $dataList[] = [
    'Значение < 1-' . $i,
    'Значение > 2-' . $i,
    'Значение 3-' . $i,
    'Значение 4-' . $i,
    'Значение 5-' . $i,
    'Значение 6-' . $i,
    'Значение 7-' . $i,
    'Значение 8-' . $i,
    'Значение 9-' . $i,
    'Значение 10-' . $i
  ];
}

$workbook = new Workbook(__DIR__ . '/templates/simple_template.xlsx');
$workbook->extract();

$workbook->renameWorksheet(1, 'Новое название листа');

$sheet = $workbook->getWorksheet(1);

// Сохраняем номера стилей
$style1 = $sheet->getStyleIndex('A5');

$sheet->setCellValue('D2', "Новый заголовок");
$sheet->saveCells();

$columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];
$sheet->initRowsInserting(5, $columns);

$styleMap = [];
foreach ($columns as $column) {
  $styleMap[] = $style1;
}

foreach ($dataList as $row) {
  foreach ($row as $key => $item) {
    $row[$key] = htmlspecialchars($item);
  }
  $sheet->insertRow($row, $styleMap);
}
$sheet->finishRowsInserting();

$workbook->createXlsx(__DIR__ . '/data/simple_output.xlsx');
$workbook->clean();
