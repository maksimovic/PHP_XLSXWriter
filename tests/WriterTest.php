<?php

class WriterTest extends \PHPUnit\Framework\TestCase
{
    /** @var XLSXWriter */
    private $writer;

    public function setUp(): void
    {
        $this->writer = new XLSXWriter();
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
        $this->writer->setSubject('Some Subject');
        $this->writer->setAuthor('Some Author');
        $this->writer->setCompany('Some Company');
        $this->writer->setKeywords(['some','interesting','keywords']);
        $this->writer->setDescription('Some interesting description');

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



