<?php
	// include "src/XLSXWriter_BuffererWriter.php";
	// include "src/XLSXWriter.php";

	$rows = array(
		array(
			"IMAGE(URL,ALT_TEXT,0)",
			"Fit to cell is the most common selection as that preserves the image aspect ratio and grows/shrinks with cell size",
			"0 = Fit to cell",
			'=IMAGE("https://images.unsplash.com/photo-1658203897339-0b8c64a42fba?ixid=M3w4MzU4MTJ8MHwxfHNlYXJjaHwxfHxleGNlbHxlbnwwfHx8fDE3ODMwMDMyOTN8MA&ixlib=rb-4.1.0&w=512&h=512&q=100&fm=png&fit=crop&dpr=1", "Excel Splash Image",0)',
		),
		array(
			"IMAGE(URL,ALT_TEXT,1)",
			"Fill cell is the hardest option to use, as it distorts the image and is nearly always difficult to work with.",
			"1 = Fill cell",
			'=IMAGE("https://images.unsplash.com/photo-1658203897339-0b8c64a42fba?ixid=M3w4MzU4MTJ8MHwxfHNlYXJjaHwxfHxleGNlbHxlbnwwfHx8fDE3ODMwMDMyOTN8MA&ixlib=rb-4.1.0&w=512&h=512&q=100&fm=png&fit=crop&dpr=1", "Excel Splash Image",1)',
		),
		array(
			"IMAGE(URL,ALT_TEXT,2)",
			"Original Size renders the image at whatever the file resolution specifies.",
			"2 = Original Size",
			'=IMAGE("https://images.unsplash.com/photo-1658203897339-0b8c64a42fba?ixid=M3w4MzU4MTJ8MHwxfHNlYXJjaHwxfHxleGNlbHxlbnwwfHx8fDE3ODMwMDMyOTN8MA&ixlib=rb-4.1.0&w=512&h=512&q=100&fm=png&fit=crop&dpr=1", "Excel Splash Image",2)',
		),
		array(
			"IMAGE(URL,ALT_TEXT,3,128,128)",
			"Custom size allows you to specify the height/width in pixels as the 4th and 5th arguments.",
			"3 = Custom Size",
			'=IMAGE("https://images.unsplash.com/photo-1658203897339-0b8c64a42fba?ixid=M3w4MzU4MTJ8MHwxfHNlYXJjaHwxfHxleGNlbHxlbnwwfHx8fDE3ODMwMDMyOTN8MA&ixlib=rb-4.1.0&w=512&h=512&q=100&fm=png&fit=crop&dpr=1", "Excel Splash Image",3,128,128)',
		),
		array(
			"IMAGE(C6,B6,0)",
			"URL can be selected from separate cell by cell address.",
			"https://images.unsplash.com/photo-1529078155058-5d716f45d604?ixid=M3w4MzU4MTJ8MHwxfHNlYXJjaHwzfHxleGNlbHxlbnwwfHx8fDE3ODMwMDMyOTN8MA&ixlib=rb-4.1.0&w=512&h=512&q=80&fm=jpg&fit=crop&dpr=1",
			'=IMAGE(C6,B6,0)',
		),
	);

	$writer = new XLSXWriter();
	
	$header_opts = array(
		'widths'         => array(48,64,64,64),
		'font-size'      => 10,
		'font-style'     => 'bold',
		'color'          => '#ffffff',
		'fill'           => '#5800e8',
		'supress_row'    => true,
		'halign'         => 'center',
	);
	$header = array(
		'Formula'        => '@',
		'Description'    => '@',
		'Setting/Option' => '@',
		'Rendered Image' => 'General',
	);
	$writer->writeSheetHeader('ImageSheet', $header, $header_opts);

	$row_opts = array(
		array(
			'wrap_text' => true,
			'font'      => 'Courier New',
			'color'     => '#5800e8',
			'halign'    => 'top',
			'height'    => 48,
		), 
		array(
			'wrap_text' => true,
			'halign'    => 'top',
			'height'    => 48,
		), 
		array(
			'wrap_text' => true,
			'halign'    => 'top',
			'height'    => 48,
		), 
		array(
			'wrap_text' => true,
			'halign'    => 'top',
			'height'    => 48,
		),
		'height' => 128,
	);
	
	foreach($rows as $row) {
		$writer->writeSheetRow('ImageSheet', $row, $row_opts);
	}

	$writer->writeToFile('xlsx-images.xlsx');
	//$writer->writeToStdOut();
	//echo $writer->writeToString();

