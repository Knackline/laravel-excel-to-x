<?php

namespace Knackline\ExcelTo\Tests\Feature;

use Knackline\ExcelTo\ExcelTo;
use PHPUnit\Framework\TestCase;

class ExcelToStreamingFeatureTest extends TestCase
{
    protected function setUp(): void
    {
        parent::setUp();
        $this->createTestExcelFile();
    }

    protected function tearDown(): void
    {
        parent::tearDown();
        if (file_exists($this->getTestFilePath())) {
            unlink($this->getTestFilePath());
        }
    }

    private function getTestFilePath(): string
    {
        return __DIR__ . '/../test_files/test_stream.xlsx';
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
        $sheet1->setCellValue('D1', 'Date');
        $sheet1->setCellValue('A2', 'John');
        $sheet1->setCellValue('B2', 'Doe');
        $sheet1->setCellValue('C2', '30');
        $sheet1->setCellValue('D2', '2023-01-01');

        // Create directory if it doesn't exist
        $dir = dirname($this->getTestFilePath());
        if (!is_dir($dir)) {
            mkdir($dir, 0755, true);
        }

        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($this->getTestFilePath());
    }

    public function test_stream_processes_large_files_in_chunks()
    {
        $result = ExcelTo::stream($this->getTestFilePath(), 1);

        $this->assertIsArray($result);
        $this->assertEquals([
            'First Name' => 'John',
            'Last Name' => 'Doe',
            'Age' => '30',
            'Date' => '2023-01-01'
        ], $result[0]);
    }

    public function test_stream_handles_custom_chunk_size()
    {
        $result = ExcelTo::stream($this->getTestFilePath(), 2);

        $this->assertIsArray($result);
        $this->assertCount(1, $result);
    }

    public function test_stream_handles_empty_sheets()
    {
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $emptyFilePath = __DIR__ . '/../test_files/test_empty.xlsx';
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($emptyFilePath);

        $result = ExcelTo::stream($emptyFilePath, 1);

        $this->assertIsArray($result);
        $this->assertEmpty($result);

        unlink($emptyFilePath);
    }
}
