<?php

require dirname( __DIR__ ) . '/vendor/autoload.php';

use kasparsd\MiniSheets\XlsxBuilder;

$zipper = new \ZipArchive;
$filename = __DIR__ . '/basic-xlsx.xlsx';
$zipper->open( $filename, \ZipArchive::CREATE | \ZipArchive::OVERWRITE );

$builder = new XlsxBuilder( $zipper );
$builder->add_rows(
	[
		[
			'row 1, col1',
			'row 1, col2',
		],
		[
			'row 2, col1',
			'row 2, col2',
		],
	]
);

$builder->build();

$zipper->close(); // See basic-xlsx.xlsx for the generated output.
