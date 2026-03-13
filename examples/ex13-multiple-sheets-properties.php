<?php

$header = array(
    'year'=>'string',
    'month'=>'string',
    'amount'=>'price',
    'first_event'=>'datetime',
    'second_event'=>'date',
);
$data1 = array(
    array('2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'),
    array('2003','=B2', '23.5','2010-01-01 00:00:00','2012-12-31 00:00:00'),
    array('2003',"'=B2", '23.5','2010-01-01 00:00:00','2012-12-31 00:00:00'),
);
$data2 = array(
    array('2003','01','343.12','4000000000'),
    array('2003','02','345.12','2000000000'),
);
$writer = new XLSXWriter();
$writer->writeSheetHeader('Sheet1', $header);
foreach($data1 as $row)
	$writer->writeSheetRow('Sheet1', $row);
$writer->finalizeSheet($subsetname, ['orientation' => 'landscape', 'header' => 'Custom printed header for Sheet 1 - '.  date('Y-m-d')]);

foreach($data2 as $row)
	$writer->writeSheetRow('Sheet2', $row);
$writer->finalizeSheet($subsetname, ['orientation' => 'portrait', 'header' => 'Custom printed header for Sheet 2 - '.  date('Y-m-d'), 'printOptions' => ['gridLines' => 'true']]);

/*
writeToFile calls finalizeSheet for all sheets as well, but finalizeSheet checks if it has been applied before, so you're good to go
available options for custom values as strings are
[
    'orientation' => 'portrait',
    'header' => '&C&A',
    'footer' => '- &P -',
    'paperSize' => '1',
    'margins' => [
        'left' => '0.5',
        'right' => '0.5',
        'top' => '1',
        'bottom' => '1',
        'header' => '0.5',
        'footer' => '0.5',
    ],
    'printOptions' => [
        'headings' => 'false',
        'gridLines' => 'false'
    ]
]

 */$writer->writeToFile('xlsx-sheets.xlsx');
//$writer->writeToStdOut();
//echo $writer->writeToString();

exit(0);


