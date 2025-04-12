<?php

namespace Knackline\ExcelTo\Tests\Feature;

use Knackline\ExcelTo\ExcelTo;
use PHPUnit\Framework\TestCase;

class ExcelToJsonTest extends TestCase
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
        return __DIR__ . '/../test_files/test.xlsx';
    }

    private function createTestExcelFile(): void
    {
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setCellValue('A1', 'Name');
        $sheet->setCellValue('B1', 'Age');
        $sheet->setCellValue('A2', 'John Doe');
        $sheet->setCellValue('B2', '30');

        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($this->getTestFilePath());
    }

    /** @test */
    public function it_can_convert_excel_to_json()
    {
        $json = ExcelTo::json($this->getTestFilePath());
        $this->assertIsString($json);
        $data = json_decode($json, true);
        $this->assertIsArray($data);
        $this->assertCount(1, $data);
        $this->assertEquals([
            'Name' => 'John Doe',
            'Age' => '30'
        ], $data[0]);
    }
}
