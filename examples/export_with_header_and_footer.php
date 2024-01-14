<?php
// You can run this script from bash: php export_with_header_and_footer.php

require_once '../library/PclZip.php';
require_once '../library/Workbook.php';
require_once '../library/Worksheet.php';
require_once '../library/SharedStrings.php';
require_once '../library/Formula.php';

use ExcelExport\Workbook;


$dataToInsert = [
  ['row 1', 1, 2, 'some text'],
  ['row 2', 3, 4, 'some text 2'],
  ['row 3', 5, 6, 'some text 3']
];

// Workbook initialization
$workbook = new Workbook(__DIR__ . '/templates/template_with_header_and_footer.xlsx'); // Xlsx template created in Microsoft Excel

// Extracting xlsx to temp folder
$workbook->extract();

// Get first sheet from workbook
$sheet = $workbook->getWorksheet(1);


// STEP1: Insert values to header
// IMPORTANT!
// You have to insert cell values to header and footer before inserting rows.
// You need initialize cells (set temporary value in template) you want to set value.
$sheet->setCellValue('E1', 'Some title');
$sheet->setCellValue('E2', 'Some subtitle');
$sheet->setCellValue('E4', date('d.m.Y'));


// STEP2: Insert rows
// Selecting columns to insert. You can skip unused columns.
$columns = ['A', 'B', 'C', 'E'];

// Initiate inserting rows mode
// IMPORTANT! First inserted row must not be empty. Set temporary value to first cell in template.
$sheet->initRowsInserting(3, $columns);

// Inserting rows
foreach ($dataToInsert as $row) {
  $sheet->insertRow($row);
}

// Exiting from inserting rows mode
$sheet->finishRowsInserting();

$exportPath = __DIR__ . '/exported_files/output_with_header_and_footer.xlsx';

// Creating xlsx file from temp folder
$workbook->createXlsx($exportPath);

// Removing temp folder
$workbook->clean();

echo 'File generated: ' . $exportPath . PHP_EOL;
