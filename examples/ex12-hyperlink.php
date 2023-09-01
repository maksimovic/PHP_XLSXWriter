<?php


$rows = array(
  array("not clickable","also not clickable",'=HYPERLINK("http://github.com", "Click for Github.com")'),
);
$writer = new XLSXWriter();
foreach($rows as $row)
	$writer->writeSheetRow('Sheet1', $row);

$writer->writeToFile('xlsx-hyperlink.xlsx');
//$writer->writeToStdOut();
//echo $writer->writeToString();

