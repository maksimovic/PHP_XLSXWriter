<?php

class BuffererWriterTest extends \PHPUnit\Framework\TestCase
{
    private function getTmpFileName(): string
    {
        return tempnam(sys_get_temp_dir(), 'xlsxBufWriter');
    }

    public function testWriteAndClose(): void
    {
        $file = $this->getTmpFileName();
        $writer = new XLSXWriter_BuffererWriter($file);

        $writer->write('hello world');
        $writer->close();

        $this->assertSame('hello world', file_get_contents($file));
        unlink($file);
    }

    public function testFtellReturnsPosition(): void
    {
        $file = $this->getTmpFileName();
        $writer = new XLSXWriter_BuffererWriter($file);

        $writer->write('hello');
        $pos = $writer->ftell();

        $this->assertSame(5, $pos);

        $writer->close();
        unlink($file);
    }

    public function testFtellReturnsNegativeOneWhenClosed(): void
    {
        $file = $this->getTmpFileName();
        $writer = new XLSXWriter_BuffererWriter($file);

        $writer->close();

        $this->assertSame(-1, $writer->ftell());
        unlink($file);
    }

    public function testFseekMovesPosition(): void
    {
        $file = $this->getTmpFileName();
        $writer = new XLSXWriter_BuffererWriter($file);

        $writer->write('hello world');
        $result = $writer->fseek(0);

        $this->assertSame(0, $result);

        // Overwrite from beginning
        $writer->write('HELLO');
        $writer->close();

        $this->assertSame('HELLO world', file_get_contents($file));
        unlink($file);
    }

    public function testFseekReturnsNegativeOneWhenClosed(): void
    {
        $file = $this->getTmpFileName();
        $writer = new XLSXWriter_BuffererWriter($file);

        $writer->close();

        $this->assertSame(-1, $writer->fseek(0));
        unlink($file);
    }

    public function testPurgeWithUtf8CheckValid(): void
    {
        $file = $this->getTmpFileName();
        $writer = new XLSXWriter_BuffererWriter($file, 'w', true);

        $writer->write('valid UTF-8 string');
        $writer->close();

        $this->assertSame('valid UTF-8 string', file_get_contents($file));
        unlink($file);
    }

    public function testPurgeWithUtf8CheckInvalid(): void
    {
        $file = $this->getTmpFileName();
        $writer = new XLSXWriter_BuffererWriter($file, 'w', true);

        // Write invalid UTF-8 bytes - this should trigger the error log
        $capture = tmpfile();
        $saved = ini_set('error_log', stream_get_meta_data($capture)['uri']);

        $writer->write("\xC0\xAF invalid utf8");
        $writer->close();

        ini_set('error_log', $saved);

        // The data should still be written despite the UTF-8 error
        $this->assertSame("\xC0\xAF invalid utf8", file_get_contents($file));
        unlink($file);
    }

    public function testPurgeLargeBuffer(): void
    {
        $file = $this->getTmpFileName();
        $writer = new XLSXWriter_BuffererWriter($file);

        // Write more than 8192 bytes to trigger automatic purge
        $largeString = str_repeat('a', 9000);
        $writer->write($largeString);
        $writer->close();

        $this->assertSame($largeString, file_get_contents($file));
        unlink($file);
    }

    public function testDestructorClosesFile(): void
    {
        $file = $this->getTmpFileName();
        $writer = new XLSXWriter_BuffererWriter($file);

        $writer->write('destruct test');
        unset($writer); // triggers __destruct

        $this->assertSame('destruct test', file_get_contents($file));
        unlink($file);
    }
}
