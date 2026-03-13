<?php

class XLSXWriterStaticMethodsTest extends \PHPUnit\Framework\TestCase
{
    // --- xlsCell ---

    public function testXlsCellRelative(): void
    {
        $this->assertSame('A1', XLSXWriter::xlsCell(0, 0));
        $this->assertSame('B2', XLSXWriter::xlsCell(1, 1));
        $this->assertSame('Z1', XLSXWriter::xlsCell(0, 25));
        $this->assertSame('AA1', XLSXWriter::xlsCell(0, 26));
    }

    public function testXlsCellAbsolute(): void
    {
        $this->assertSame('$A$1', XLSXWriter::xlsCell(0, 0, true));
        $this->assertSame('$C$5', XLSXWriter::xlsCell(4, 2, true));
        $this->assertSame('$AA$1', XLSXWriter::xlsCell(0, 26, true));
    }

    // --- sanitize_sheetname ---

    public function testSanitizeSheetnameRemovesBadChars(): void
    {
        $this->assertSame('clean', XLSXWriter::sanitize_sheetname('clean'));
        $this->assertSame('a b c d e f g', XLSXWriter::sanitize_sheetname('a\\b/c?d*e:f[g]'));
    }

    public function testSanitizeSheetnameTruncatesTo31Chars(): void
    {
        $long = str_repeat('a', 50);
        $result = XLSXWriter::sanitize_sheetname($long);
        $this->assertSame(31, strlen($result));
    }

    public function testSanitizeSheetnameTrimsQuotes(): void
    {
        $this->assertSame('Sheet', XLSXWriter::sanitize_sheetname("'Sheet'"));
    }

    public function testSanitizeSheetnameEmptyReturnsFallback(): void
    {
        $result = XLSXWriter::sanitize_sheetname('');
        $this->assertMatchesRegularExpression('/^Sheet\d{3}$/', $result);
    }

    // --- log ---

    public function testLogStringMessage(): void
    {
        $capture = tmpfile();
        $saved = ini_set('error_log', stream_get_meta_data($capture)['uri']);

        XLSXWriter::log('Test log message');

        ini_set('error_log', $saved);

        rewind($capture);
        $output = stream_get_contents($capture);
        $this->assertStringContainsString('Test log message', $output);
    }

    public function testLogArrayMessage(): void
    {
        $capture = tmpfile();
        $saved = ini_set('error_log', stream_get_meta_data($capture)['uri']);

        XLSXWriter::log(['key' => 'value']);

        ini_set('error_log', $saved);

        rewind($capture);
        $output = stream_get_contents($capture);
        $this->assertStringContainsString('"key":"value"', $output);
    }

    // --- add_to_list_get_index ---

    public function testAddToListGetIndexNewItem(): void
    {
        $list = ['a', 'b', 'c'];
        $idx = XLSXWriter::add_to_list_get_index($list, 'd');
        $this->assertSame(3, $idx);
        $this->assertCount(4, $list);
        $this->assertSame('d', $list[3]);
    }

    public function testAddToListGetIndexExistingItem(): void
    {
        $list = ['a', 'b', 'c'];
        $idx = XLSXWriter::add_to_list_get_index($list, 'b');
        $this->assertSame(1, $idx);
        $this->assertCount(3, $list); // no duplicate added
    }

    // --- convert_date_time ---

    public function testConvertDateTimeStandardDate(): void
    {
        // 2018-12-31 should give a known Excel serial number
        $result = XLSXWriter::convert_date_time('2018-12-31');
        $this->assertSame(43465.0, (float)$result);
    }

    public function testConvertDateTimeWithTime(): void
    {
        $result = XLSXWriter::convert_date_time('2018-12-31 12:00:00');
        $this->assertEqualsWithDelta(43465.5, $result, 0.0001);
    }

    public function testConvertDateTimeOnlyTime(): void
    {
        // Time-only: no date regex match, year/month/day stay 0, boundary check is skipped
        $result = XLSXWriter::convert_date_time('12:00:00');
        // When no date part, days calculation gives a negative result; just verify it runs
        $this->assertIsNumeric($result);

        // Use the 1899-12-31 epoch date to get pure time
        $result2 = XLSXWriter::convert_date_time('1899-12-31 12:00:00');
        $this->assertEqualsWithDelta(0.5, $result2, 0.0001);
    }

    public function testConvertDateTimeExcelEpoch1900(): void
    {
        // 1899-12-31 => seconds only (epoch)
        $this->assertSame(0.0, (float)XLSXWriter::convert_date_time('1899-12-31'));
    }

    public function testConvertDateTimeExcelFalseLeapday(): void
    {
        // 1900-02-29 => 60
        $result = XLSXWriter::convert_date_time('1900-02-29');
        $this->assertSame(60.0, (float)$result);
    }

    public function testConvertDateTimeInvalidYear(): void
    {
        // Year before 1900 epoch
        $this->assertSame(0, XLSXWriter::convert_date_time('1800-01-01'));
    }

    public function testConvertDateTimeInvalidMonth(): void
    {
        $this->assertSame(0, XLSXWriter::convert_date_time('2020-13-01'));
    }

    public function testConvertDateTimeInvalidDay(): void
    {
        $this->assertSame(0, XLSXWriter::convert_date_time('2020-02-30'));
    }

    public function testConvertDateTimeLeapYear(): void
    {
        // 2000 is a leap year (divisible by 400)
        $result = XLSXWriter::convert_date_time('2000-02-29');
        $this->assertGreaterThan(0, $result);
    }

    public function testConvertDateTimeNonLeapCentury(): void
    {
        // 1900 is NOT a leap year (divisible by 100 but not 400)
        // But Excel treats it as one, so 1900-02-29 is the special case
        // 2100-02-29 should return 0 (invalid day for non-leap century)
        $this->assertSame(0, XLSXWriter::convert_date_time('2100-02-29'));
    }
}
