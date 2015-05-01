#!/usr/bin/env php

<?php

require_once(__DIR__.'/../src/SimpleXLSX.php');
require_once(__DIR__.'/../src/SimpleXLSXStyle.php');
require_once(__DIR__.'/../src/SimpleXLSXWorkbook.php');
require_once(__DIR__.'/../src/SimpleXLSXWorksheet.php');


$xlsx = new SimpleXLSX();
$book = $xlsx->createWorkbook();
$sheet = $book->createWorksheet();

for ($row=0; $row<1000; $row++) {
    for ($col=0; $col<32; $col++) {
	$value = ($col % 2)==1 ? 'Hello,' : 'World!';
	$sheet->setCellVal($row, $col, $value);
    }
}

$path = __DIR__.'/sample.xlsx';
$book->save( $path );
echo "Saved to $path\n";