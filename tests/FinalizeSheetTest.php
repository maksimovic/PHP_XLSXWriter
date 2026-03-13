<?php

class FinalizeSheetTest extends \PHPUnit\Framework\TestCase
{
    /** @var XLSXWriter */
    private $writer;

    public function setUp(): void
    {
        $this->writer = new TestXLSWriter();
    }

    public function testFinalizeSheetWithCustomOptions(): void
    {
        $this->writer->writeSheetHeader('Sheet1', ['col1' => 'string', 'col2' => 'integer']);
        $this->writer->writeSheetRow('Sheet1', ['hello', 42]);

        $this->writer->finalizeSheet('Sheet1', [
            'orientation' => 'landscape',
            'header' => '&LCustom Header',
            'footer' => '&RPage &P',
            'paperSize' => '9',
            'margins' => [
                'left' => '0.75',
                'right' => '0.75',
                'top' => '1.25',
                'bottom' => '1.25',
                'header' => '0.3',
                'footer' => '0.3',
            ],
            'printOptions' => [
                'headings' => 'true',
                'gridLines' => 'true',
            ],
        ]);

        // Sheet should be finalized; writing to file should succeed
        $fileName = tempnam(sys_get_temp_dir(), 'xlsxFinalize');
        $this->writer->writeToFile($fileName);
        $this->assertFileExists($fileName);
        $this->assertGreaterThan(0, filesize($fileName));
        unlink($fileName);
    }

    public function testFinalizeSheetEmptyNameIsNoop(): void
    {
        $this->writer->writeSheetRow('Sheet1', ['test']);

        // Finalize with empty name should not error
        $this->writer->finalizeSheet('');

        $fileName = tempnam(sys_get_temp_dir(), 'xlsxFinalize');
        $this->writer->writeToFile($fileName);
        $this->assertFileExists($fileName);
        unlink($fileName);
    }

    public function testFinalizeSheetCalledTwiceIsIdempotent(): void
    {
        $this->writer->writeSheetHeader('Sheet1', ['col1' => 'string']);
        $this->writer->writeSheetRow('Sheet1', ['data']);

        $this->writer->finalizeSheet('Sheet1');
        $this->writer->finalizeSheet('Sheet1'); // second call should be no-op

        $fileName = tempnam(sys_get_temp_dir(), 'xlsxFinalize');
        $this->writer->writeToFile($fileName);
        $this->assertFileExists($fileName);
        unlink($fileName);
    }

    public function testFinalizeSheetPartialOptionsUsesDefaults(): void
    {
        $this->writer->writeSheetHeader('Sheet1', ['col1' => 'string']);
        $this->writer->writeSheetRow('Sheet1', ['value']);

        // Only override orientation, rest should use defaults
        $this->writer->finalizeSheet('Sheet1', [
            'orientation' => 'landscape',
        ]);

        $fileName = tempnam(sys_get_temp_dir(), 'xlsxFinalize');
        $this->writer->writeToFile($fileName);
        $this->assertFileExists($fileName);
        unlink($fileName);
    }

    /**
     * @phan-suppress PhanTypeObjectUnsetDeclaredProperty
     */
    public function tearDown(): void
    {
        unset($this->writer);
    }
}
