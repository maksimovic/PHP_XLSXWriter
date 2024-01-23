<?php

class WriterTest extends \PHPUnit\Framework\TestCase
{
    /** @var XLSXWriter */
    private $writer;

    public function setUp(): void
    {
        $this->writer = new TestXLSWriter();
    }

    public function testCreate(): void
    {
        $header = array(
            'year'=>'string',
            'month'=>'string',
            'amount'=>'money',
            'first_event'=>'datetime',
            'second_event'=>'date',
        );

        $data2 = [
            ['2003','01','343.12'],
            ['2003','02','345.12'],
        ];

        $this->writer->setTempDir(sys_get_temp_dir());
        $this->writer->setRightToLeft(false);

        $this->writer->setTitle('Some Title');
        $this->writer->setTabRatio(2000);

        // check that 1000 can't be exceeded
        self::assertEquals(1000, $this->writer->getTabRatio());

        $this->writer->setTabRatio(500);
        self::assertEquals(500, $this->writer->getTabRatio());

        $this->writer->setSubject('Some Subject');
        $this->writer->setAuthor('Some Author');
        $this->writer->setCompany('Some Company');
        $this->writer->setKeywords(['some','interesting','keywords']);
        $this->writer->setDescription('Some interesting description');

        $this->writer->writeSheetHeader("Test sheet", ['c1-text'=>'string'], ['widths'=>[null,null,null,40]]);

        $this->writer->writeSheetRow("Sheet1", array('2003','3', '33.3','2023-02-01 00:00:00','2012-12-31 00:00:00'), ['hidden' => false]);
        $this->writer->markMergedCell("Sheet1", "A", 2, "B", 3);

        $this->writer->writeSheet($this->getTestRowsData(),'Sheet1', $header);
        $this->writer->writeSheet($data2, 'Sheet2');

        $this->writer->writeSheetRow("", []);

        $this->assertEquals(4, $this->writer->countSheetRows("Sheet1"));

        $this->assertEquals("test", XLSXWriter::sanitize_filename("t<>e?\"s:/\\|*&t".chr(0)));

        $fileName = $this->getTmpFileName();
        $this->writer->writeToFile($fileName);
        $this->assertFileExists($fileName);

        $this->assertEquals($this->writer->writeToString(), file_get_contents($fileName));
    }

    public function testNoFileWritten(): void
    {
        $this->writer->setKeywords(['some','interesting','keywords']);

        $fileName = $this->getTmpFileName();
        file_put_contents($fileName, "test");

        chmod($fileName, 600);

        // don't print the expected error log to stdout
        $capture = tmpfile();
        $saved = ini_set('error_log', stream_get_meta_data($capture)['uri']);

        $this->writer->writeToFile($fileName);

        // revert error logging
        ini_set('error_log', $saved);

        chmod($fileName, 777);
        $this->assertSame("test", file_get_contents($fileName));
    }

    public function testWriteToStdout(): void
    {
        $this->writer->setKeywords(['some','interesting','keywords']);
        $this->writer->writeSheetRow("Sheet1", array('2003','3', '33.3','2023-02-01 00:00:00','2012-12-31 00:00:00'), ['hidden' => false]);

        ob_start();
        $this->writer->writeToStdOut();

        $contents = ob_get_clean();

        $this->assertSame($this->writer->writeToString(), $contents);
    }

    /**
     * @doesNotPerformAssertions
     */
    public function testFreezeRowsAndColumns(): void
    {
        $this->writer->writeSheetHeader("Sheet1", ["money"], ['freeze_rows' => 1, 'freeze_columns' => 1, 'widths' => [10, 10, 10, 10]]);

        $this->writeToFile();
    }

    /**
     * @doesNotPerformAssertions
     */
    public function testFreezeRows(): void
    {
        $this->writer->writeSheetHeader("Sheet1", ["money"], ['freeze_rows' => 1]);
        $this->writeToFile();
    }

    /**
     * @doesNotPerformAssertions
     */
    public function testFreezeColumns(): void
    {
        $this->writer->writeSheetHeader("Sheet1", ["money"], ['freeze_columns' => 1]);
        $this->writeToFile();
    }

    public function testFormars(): void
    {
        $sheet1header = array(
            'c1-string'=>'string',
            'c2-integer'=>'integer',
            'c3-custom-integer'=>'0',
            'c4-custom-1decimal'=>'0.0',
            'c5-custom-2decimal'=>'0.00',
            'c6-custom-percent'=>'0%',
            'c7-custom-percent1'=>'0.0%',
            'c8-custom-percent2'=>'0.00%',
            'c9-custom-text'=>'@',//text
        );
        $sheet2header = array(
            'col1-date'=>'date',
            'col2-datetime'=>'datetime',
            'col3-time'=>'time',
            'custom-date1'=>'YYYY-MM-DD',
            'custom-date2'=>'MM/DD/YYYY',
            'custom-date3'=>'DD-MMM-YYYY HH:MM AM/PM',
            'custom-date4'=>'MM/DD/YYYY HH:MM:SS',
            'custom-date5'=>'YYYY-MM-DD HH:MM:SS',
            'custom-date6'=>'YY MMMM',
            'custom-date7'=>'QQ YYYY',
            'custom-time1'=>'HH:MM',
            'custom-time2'=>'HH:MM:SS',
        );
        $sheet3header = array(
            'col1-dollar'=>'dollar',
            'col2-euro'=>'euro',
            'custom-amount1'=>'0',
            'custom-amount2'=>'0.0',//1 decimal place
            'custom-amount3'=>'0.00',//2 decimal places
            'custom-currency1'=>'#,##0.00',//currency 2 decimal places, no currency/dollar sign
            'custom-currency2'=>'[$$-1009]#,##0.00;[RED]-[$$-1009]#,##0.00',//w/dollar sign
            'custom-currency3'=>'#,##0.00 [$€-407];[RED]-#,##0.00 [$€-407]',//w/euro sign
            'custom-currency4'=>'[$￥-411]#,##0;[RED]-[$￥-411]#,##0', //japanese yen
            'custom-scientific'=>'0.00E+000',//-1.23E+003 scientific notation
        );
        $pi = 3.14159;
        $date = '2018-12-31 23:59:59';
        $time = '23:59:59';
        $amount = '5120.5';
        
        $this->writer->setAuthor('Some Author');
        $this->writer->writeSheetHeader('BasicFormats',$sheet1header);
        $this->writer->writeSheetRow('BasicFormats',array($pi,$pi,$pi,$pi,$pi,$pi,$pi,$pi,$pi) );
        $this->writer->writeSheetHeader('Dates',$sheet2header);
        $this->writer->writeSheetRow('Dates',array($date,$date,$date,$date,$date,$date,$date,$date,$date,$date,$time,$time) );
        $this->writer->writeSheetHeader('Currencies',$sheet3header);
        $this->writer->writeSheetRow('Currencies',array($amount,$amount,$amount,$amount,$amount,$amount,$amount,$amount,$amount) );

        $fileName = $this->getTmpFileName();
        $this->writer->writeToFile($fileName);
        $this->assertFileExists($fileName);
    }
    
    public function testStyles(): void
    {
        $styles1 = array( 'font'=>'Arial','font-size'=>10,'font-style'=>'bold', 'fill'=>'#eee', 'halign'=>'center', 'border'=>'left,right,top,bottom');
        $styles2 = array( ['font-size'=>6],['font-size'=>8],['font-size'=>10],['font-size'=>16] );
        $styles3 = array( ['font'=>'Arial'],['font'=>'Courier New'],['font'=>'Times New Roman'],['font'=>'Comic Sans MS']);
        $styles4 = array( ['font-style'=>'bold'],['font-style'=>'italic'],['font-style'=>'underline'],['font-style'=>'strikethrough']);
        $styles5 = array( ['color'=>'#f00'],['color'=>'#0f0'],['color'=>'#00f'],['color'=>'#666']);
        $styles6 = array( ['fill'=>'#ffc'],['fill'=>'#fcf'],['fill'=>'#ccf'],['fill'=>'#cff']);
        $styles7 = array( 'border'=>'left,right,top,bottom');
        $styles8 = array( ['halign'=>'left'],['halign'=>'right'],['halign'=>'center'],['halign'=>'none']);
        $styles9 = array( array(),['border'=>'left,top,bottom'],['border'=>'top,bottom'],['border'=>'top,bottom,right']);


        $this->writer->writeSheetRow('Sheet1', array(300,234,456,789), $styles1 );
        $this->writer->writeSheetRow('Sheet1', array(300,234,456,789), $styles2 );
        $this->writer->writeSheetRow('Sheet1', array(300,234,456,789), $styles3 );
        $this->writer->writeSheetRow('Sheet1', array(300,234,456,789), $styles4 );
        $this->writer->writeSheetRow('Sheet1', array(300,234,456,789), $styles5 );
        $this->writer->writeSheetRow('Sheet1', array(300,234,456,789), $styles6 );
        $this->writer->writeSheetRow('Sheet1', array(300,234,456,789), $styles7 );
        $this->writer->writeSheetRow('Sheet1', array(300,234,456,789), $styles8 );
        $this->writer->writeSheetRow('Sheet1', array(300,234,456,789), $styles9 );

        $fileName = $this->getTmpFileName();
        $this->writer->writeToFile($fileName);
        $this->assertFileExists($fileName);
    }

    public function testColors()
    {
        $colors = ['ff','cc','99','66','33','00'];
        foreach($colors as $b) {
            foreach($colors as $g) {
                $rowdata = [];
                $rowstyle = [];
                foreach($colors as $r) {
                    $rowdata[] = "#$r$g$b";
                    $rowstyle[] = ['fill'=>"#$r$g$b"];
                }

                $this->writer->writeSheetRow('Sheet1', $rowdata, $rowstyle );
            }
        }

        $fileName = $this->getTmpFileName();
        $this->writer->writeToFile($fileName);
        $this->assertFileExists($fileName);
    }
    
    public function testWidths(): void
    {
        $header = array(
            "col1"=>"string",
            "col2"=>"string",
            "col3"=>"string",
            "col4"=>"string",
        );

        $this->writer->writeSheetHeader('Sheet1', $header, $col_options = ['widths'=>[10,20,30,40]] );
        $this->writer->writeSheetRow('Sheet1', $rowdata = array(300,234,456,789), $row_options = ['height'=>20] );
        $this->writer->writeSheetRow('Sheet1', $rowdata = array(300,234,456,789), $row_options = ['height'=>30] );
        $this->writer->writeSheetRow('Sheet1', $rowdata = array(300,234,456,789), $row_options = ['height'=>40] );

        $fileName = $this->getTmpFileName();
        $this->writer->writeToFile($fileName);
        $this->assertFileExists($fileName);
    }
    
    public function testAdvanced(): void
    {
        $keywords = array('some','interesting','keywords');

        $this->writer->setTitle('Some Title');
        $this->writer->setSubject('Some Subject');
        $this->writer->setAuthor('Some Author');
        $this->writer->setCompany('Some Company');
        $this->writer->setKeywords($keywords);
        $this->writer->setDescription('Some interesting description');
        $this->writer->setTempDir(sys_get_temp_dir());//set custom tempdir
        
        $sheet1 = 'merged_cells';
        $header = array("string","string","string","string","string");
        $rows = array(
            array("Merge Cells Example"),
            array(100, 200, 300, 400, 500),
            array(110, 210, 310, 410, 510),
        );
        
        $this->writer->writeSheetHeader($sheet1, $header, $col_options = ['suppress_row'=>true] );
        foreach($rows as $row)
            $this->writer->writeSheetRow($sheet1, $row);
        $this->writer->markMergedCell($sheet1, $start_row=0, $start_col=0, $end_row=0, $end_col=4);


        $sheet2 = 'utf8';
        $rows = array(
            array('Spreadsheet','_'),
            array("Hoja de cálculo", "Hoja de c\xc3\xa1lculo"),
            array("Електронна таблица", "\xd0\x95\xd0\xbb\xd0\xb5\xd0\xba\xd1\x82\xd1\x80\xd0\xbe\xd0\xbd\xd0\xbd\xd0\xb0 \xd1\x82\xd0\xb0\xd0\xb1\xd0\xbb\xd0\xb8\xd1\x86\xd0\xb0"),//utf8 encoded
            array("電子試算表", "\xe9\x9b\xbb\xe5\xad\x90\xe8\xa9\xa6\xe7\xae\x97\xe8\xa1\xa8"),//utf8 encoded
        );

        $this->writer->writeSheet($rows, $sheet2);

        $sheet3 = 'fonts';
        $format = array('font'=>'Arial','font-size'=>10,'font-style'=>'bold,italic', 'fill'=>'#eee','color'=>'#f00','fill'=>'#ffc', 'border'=>'top,bottom', 'halign'=>'center');
        $this->writer->writeSheetRow($sheet3, $row=array(101,102,103,104,105,106,107,108,109,110), $format);
        $this->writer->writeSheetRow($sheet3, $row=array(201,202,203,204,205,206,207,208,209,210), $format);

        $sheet4 = 'row_options';
        $this->writer->writeSheetHeader($sheet4, ["col1"=>"string", "col2"=>"string"], $col_options = array('widths'=>[10,10]) );
        $this->writer->writeSheetRow($sheet4, array(101,'this text will wrap'    ), $row_options = array('height'=>30,'wrap_text'=>true));
        $this->writer->writeSheetRow($sheet4, array(201,'this text is hidden'    ), $row_options = array('height'=>30,'hidden'=>true));
        $this->writer->writeSheetRow($sheet4, array(301,'this text will not wrap'), $row_options = array('height'=>30,'collapsed'=>true));

        $fileName = $this->getTmpFileName();
        $this->writer->writeToFile($fileName);
        $this->assertFileExists($fileName);
    }

    public function testAutoFilter()
    {
        $chars = 'abcdefgh';

        $header = array('col-string'=>'string','col-numbers'=>'integer','col-timestamps'=>'datetime');

        $this->writer->writeSheetHeader('Sheet1', $header, ['auto_filter'=>true, 'widths'=>[15,15,30]] );
        for($i=0; $i<1000; $i++)
        {
            $this->writer->writeSheetRow('Sheet1', array(
                str_shuffle($chars),
                mt_rand()%10000,
                date('Y-m-d H:i:s',time()-(mt_rand()%31536000))
            ));
        }

        $fileName = $this->getTmpFileName();
        $this->writer->writeToFile($fileName);
        $this->assertFileExists($fileName);
    }

    private function getTmpFileName(): string
    {
        return tempnam(sys_get_temp_dir(), 'xlsxWriter');
    }

    private function getTestRowsData(): array
    {
        return [
            ['2003','1','-50.5','2010-01-01 23:00:00','2012-12-31 23:00:00'],
            ['2003','=B2', '23.5','2010-01-01 00:00:00','2012-12-31 00:00:00'],
        ];
    }

    private function writeToFile(): void
    {
        $this->writer->writeToFile($this->getTmpFileName());
    }

    /**
     * @phan-suppress PhanTypeObjectUnsetDeclaredProperty
     *
     * @return void
     */
    public function tearDown(): void
    {
        unset($this->writer);
    }
}



