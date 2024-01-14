<?php
// You can run this script from bash: php export_basic.php

require_once '../library/PclZip.php';
require_once '../library/Workbook.php';
require_once '../library/Worksheet.php';
require_once '../library/SharedStrings.php';
require_once '../library/Formula.php';

use ExcelExport\Workbook;


$dataToInsert = [
  ['row 1', 1, 2],
  ['row 2', 3, 4],
  ['row 3', 5, 6]
];

// Workbook initialization
$workbook = new Workbook(__DIR__ . '/templates/template_basic.xlsx'); // Xlsx template created in Microsoft Excel

// Extracting xlsx to temp folder
$workbook->extract();

// Get first sheet from workbook
$sheet = $workbook->getWorksheet(1);

// Selecting columns to insert
$columns = ['A', 'B', 'C'];

// Initiate inserting rows mode
$sheet->initRowsInserting(1, $columns);

// Inserting rows
foreach ($dataToInsert as $row) {
  $sheet->insertRow($row);
}

// Exiting from inserting rows mode
$sheet->finishRowsInserting();

$exportPath = __DIR__ . '/exported_files/output_basic.xlsx';

// Creating xlsx file from temp folder
$workbook->createXlsx($exportPath);

// Removing temp folder
$workbook->clean();

echo 'File generated: ' . $exportPath . PHP_EOL;
