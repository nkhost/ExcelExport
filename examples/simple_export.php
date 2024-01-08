<?php
require_once '../library/PclZip.php';
require_once '../library/Workbook.php';
require_once '../library/Worksheet.php';
require_once '../library/SharedStrings.php';
require_once '../library/Formula.php';

use ExcelExport\Workbook;
use ExcelExport\Formula;
use ExcelExport\SharedStrings;

$linesAmount = 100000;

$dataList = [];

for ($i = 0; $i < $linesAmount; $i++) {
  $dataList[] = [
    'Value < 1-' . $i,
    'Value > 2-' . $i,
    'Value 3-' . $i,
    'Value 4-' . $i,
    'Value 5-' . $i,
    'Value 6-' . $i,
    'Value 7-' . $i,
    'Value 8-' . $i,
    'Value 9-' . $i,
    'Value 10-' . $i
  ];
}

$workbook = new Workbook(__DIR__ . '/templates/simple_template.xlsx');
$workbook->extract();

$workbook->renameWorksheet(1, 'New name');

$sheet = $workbook->getWorksheet(1);

// Сохраняем номера стилей
$style1 = $sheet->getStyleIndex('A5');

$sheet->setCellValue('D2', "New header");
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

$exportPath = __DIR__ . '/exported_files/simple_output.xlsx';

$workbook->createXlsx($exportPath);
$workbook->clean();

echo 'File generated: ' . $exportPath . PHP_EOL;
