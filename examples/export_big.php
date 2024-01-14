<?php
// You can run this script from bash: php export_big.php

require_once '../library/PclZip.php';
require_once '../library/Workbook.php';
require_once '../library/Worksheet.php';
require_once '../library/SharedStrings.php';
require_once '../library/Formula.php';

use ExcelExport\Workbook;
use ExcelExport\Formula;
use ExcelExport\SharedStrings;

// Generating data
$linesAmount = 100000;
$dataToInsert = [];
for ($i = 0; $i < $linesAmount; $i++) {
  $dataToInsert[] = [
    'Value < 1-' . $i, // special XML char <
    'Value > 2-' . $i, // special XML char >
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

// Inserting data
// Need to calculate export time
$timeStart = microtime(true);

// Workbook initialization
$workbook = new Workbook(__DIR__ . '/templates/template_big.xlsx'); // Xlsx template created in Microsoft Excel

// Extracting xlsx to temp folder
$workbook->extract();

// You can rename worksheets
$workbook->renameWorksheet(1, 'New name');

// Get first sheet from workbook
$sheet = $workbook->getWorksheet(1);

// Getting styles from template cells
$style1 = $sheet->getStyleIndex('A5');

// STEP1: Insert values to header
// IMPORTANT!
// You have to insert cell values to header and footer before inserting rows.
// You need initialize cells (set temporary value in template) you want to set value.
$sheet->setCellValue('D2', "New header");
$sheet->saveCells();


// STEP2: Insert rows
// Selecting columns to insert. You can skip unused columns.
$columns = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J'];

// Initiate inserting rows mode
// IMPORTANT! First inserted row must not be empty. Set temporary value to first cell in template.
$sheet->initRowsInserting(5, $columns);

$stylesList = [];
foreach ($columns as $column) {
  $stylesList[] = $style1;
}

// Inserting rows
foreach ($dataToInsert as $row) {
  // IMPORTANT! Escaping special characters
  foreach ($row as $key => $item) {
    $row[$key] = htmlspecialchars($item);
  }
  $sheet->insertRow($row);
}

// Exiting from inserting rows mode
$sheet->finishRowsInserting();

$exportPath = __DIR__ . '/exported_files/output_big.xlsx';

// Creating xlsx file from temp folder
$workbook->createXlsx($exportPath);

// Removing temp folder
$workbook->clean();

// Calculating execution export time
$exportTime = microtime(true) - $timeStart;

echo 'File generated: ' . $exportPath . PHP_EOL . 'Export time: ' . $exportTime . ' seconds' . PHP_EOL;
