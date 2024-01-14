<?php
// You can run this script from bash: php template_cells_styles.php

require_once '../library/PclZip.php';
require_once '../library/Workbook.php';
require_once '../library/Worksheet.php';
require_once '../library/SharedStrings.php';
require_once '../library/Formula.php';

use ExcelExport\Formula;
use ExcelExport\Workbook;


$dataToInsert = [
  ['row 1', 1, 2],
  ['row 2', 3, 4],
  ['row 3', 5, 6]
];

// Workbook initialization
$workbook = new Workbook(__DIR__ . '/templates/template_cells_styles.xlsx'); // Xlsx template created in Microsoft Excel

// Extracting xlsx to temp folder
$workbook->extract();

// Get first sheet from workbook
$sheet = $workbook->getWorksheet(1);

// Getting styles from template cells
$titleStyle = $sheet->getStyleIndex('D1');
$style1 = $sheet->getStyleIndex('G1');
$style2 = $sheet->getStyleIndex('H1');

// You can remove cells after getting the styles
$sheet->setCellValue('G1', null);
$sheet->setCellValue('H1', null);

// STEP1: Insert values to header
// IMPORTANT!
// You have to insert cell values to header and footer before inserting rows.
// You need initialize cells (set temporary value in template) you want to set value.
$sheet->setCellValue('E1', 'Some title', $titleStyle);
$sheet->setCellValue('E2', 'Some subtitle');
$sheet->setCellValue('E4', date('d.m.Y'));


// STEP2: Insert rows
// Selecting columns to insert. You can skip unused columns.
$columns = ['A', 'B', 'C', 'E'];

// Initiate inserting rows mode
// IMPORTANT! First inserted row must not be empty. Set temporary value to first cell in template.
$startRow = 3;
$sheet->initRowsInserting($startRow, $columns);

// Inserting rows
$i = $startRow;
foreach ($dataToInsert as $row) {
  // You can use formulas
  $row[] = new Formula('SUM(B' . $i . ',C' . $i . ')');
  $sheet->insertRow($row, [null, $style1, null, $style2]);
  $i++;
}

// Exiting from inserting rows mode
$sheet->finishRowsInserting();

$exportPath = __DIR__ . '/exported_files/output_cells_styles.xlsx';

// Creating xlsx file from temp folder
$workbook->createXlsx($exportPath);

// Removing temp folder
$workbook->clean();

echo 'File generated: ' . $exportPath . PHP_EOL;
