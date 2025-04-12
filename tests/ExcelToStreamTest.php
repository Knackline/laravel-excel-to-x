<?php

namespace Knackline\ExcelTo\Tests;

use Knackline\ExcelTo\ExcelTo;
use PHPUnit\Framework\TestCase;

class ExcelToStreamTest extends TestCase
{
    protected function setUp(): void
    {
        parent::setUp();
        // Create a test Excel file with multiple sheets
        $this->createTestExcelFile();
    }

    protected function tearDown(): void
    {
        parent::tearDown();
        // Clean up test files
        if (file_exists($this->getTestFilePath())) {
            unlink($this->getTestFilePath());
        }
    }

    private function getTestFilePath(): string
    {
        return __DIR__ . '/test_files/test_stream.xlsx';
    }

    private function createTestExcelFile(): void
    {
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();

        // Sheet 1
        $sheet1 = $spreadsheet->getActiveSheet();
        $sheet1->setTitle('Sheet1');
        $sheet1->setCellValue('A1', 'First Name');
        $sheet1->setCellValue('B1', 'Last Name');
        $sheet1->setCellValue('C1', 'Age');
        $sheet1->setCellValue('A2', 'John');
        $sheet1->setCellValue('B2', 'Doe');
        $sheet1->setCellValue('C2', '30');
        $sheet1->setCellValue('A3', 'Jane');
        $sheet1->setCellValue('B3', 'Smith');
        $sheet1->setCellValue('C3', '25');

        // Sheet 2
        $sheet2 = $spreadsheet->createSheet();
        $sheet2->setTitle('Sheet2');
        $sheet2->setCellValue('A1', 'Product');
        $sheet2->setCellValue('B1', 'Price');
        $sheet2->setCellValue('A2', 'Laptop');
        $sheet2->setCellValue('B2', '1000');
        $sheet2->setCellValue('A3', 'Phone');
        $sheet2->setCellValue('B3', '500');

        // Create directory if it doesn't exist
        $dir = dirname($this->getTestFilePath());
        if (!is_dir($dir)) {
            mkdir($dir, 0755, true);
        }

        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($this->getTestFilePath());
    }

    public function test_stream_returns_correct_format_for_single_sheet()
    {
        // Create a single sheet Excel file
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Name');
        $sheet->setCellValue('B1', 'Age');
        $sheet->setCellValue('A2', 'John');
        $sheet->setCellValue('B2', '30');

        $singleSheetPath = __DIR__ . '/test_files/test_single_sheet.xlsx';
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($singleSheetPath);

        $result = ExcelTo::stream($singleSheetPath, 1);

        $this->assertIsArray($result);
        $this->assertCount(1, $result); // One row of data
        $this->assertEquals([
            'Name' => 'John',
            'Age' => '30'
        ], $result[0]);

        unlink($singleSheetPath);
    }

    public function test_stream_returns_correct_format_for_multiple_sheets()
    {
        $result = ExcelTo::stream($this->getTestFilePath(), 1);

        $this->assertIsArray($result);
        $this->assertArrayHasKey('Sheet1', $result);
        $this->assertArrayHasKey('Sheet2', $result);

        // Verify Sheet1 data
        $this->assertCount(2, $result['Sheet1']);
        $this->assertEquals([
            'First Name' => 'John',
            'Last Name' => 'Doe',
            'Age' => '30'
        ], $result['Sheet1'][0]);

        // Verify Sheet2 data
        $this->assertCount(2, $result['Sheet2']);
        $this->assertEquals([
            'Product' => 'Laptop',
            'Price' => '1000'
        ], $result['Sheet2'][0]);
    }

    public function test_stream_handles_large_files_in_chunks()
    {
        // Create a large Excel file with 1000 rows
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'ID');
        $sheet->setCellValue('B1', 'Value');

        for ($i = 1; $i <= 1000; $i++) {
            $sheet->setCellValue('A' . ($i + 1), $i);
            $sheet->setCellValue('B' . ($i + 1), 'Value ' . $i);
        }

        $largeFilePath = __DIR__ . '/test_files/test_large.xlsx';
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($largeFilePath);

        $result = ExcelTo::stream($largeFilePath, 100);

        $this->assertIsArray($result);
        $this->assertCount(1000, $result);
        $this->assertEquals([
            'ID' => '1',
            'Value' => 'Value 1'
        ], $result[0]);
        $this->assertEquals([
            'ID' => '1000',
            'Value' => 'Value 1000'
        ], $result[999]);

        unlink($largeFilePath);
    }
}
