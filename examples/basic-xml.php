<?php

require dirname( __DIR__ ) . '/vendor/autoload.php';

use kasparsd\MiniSheets\XmlBuilder;

$builder = new XmlBuilder();

$builder->add_props( [
	'Author' => 'Name Surname',
] );

$builder->add_rows( [
	[
		'row 1, col1', 'row 1, col2'
	],
	[
		'row 2, col1', 'row 2, col2'
	]
] );

echo $builder->build(); // See basic-xml.xml for the generated output.
