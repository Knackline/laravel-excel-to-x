<?php

namespace Knackline\ExcelTo\Tests\Feature;

use Knackline\ExcelTo\ExcelTo;
use PHPUnit\Framework\TestCase;
use Illuminate\Support\Collection;

class ExcelToBasicFeatureTest extends TestCase
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
        return __DIR__ . '/../test_files/test_excel.xlsx';
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
        $sheet1->setCellValue('A3', 'Jane');
        $sheet1->setCellValue('B3', 'Smith');
        $sheet1->setCellValue('C3', '25');
        $sheet1->setCellValue('D3', '2023-01-02');

        // Sheet 2
        $sheet2 = $spreadsheet->createSheet();
        $sheet2->setTitle('Sheet2');
        $sheet2->setCellValue('A1', 'Product');
        $sheet2->setCellValue('B1', 'Price');
        $sheet2->setCellValue('C1', 'Stock');
        $sheet2->setCellValue('A2', 'Laptop');
        $sheet2->setCellValue('B2', '1000');
        $sheet2->setCellValue('C2', '10');
        $sheet2->setCellValue('A3', 'Phone');
        $sheet2->setCellValue('B3', '500');
        $sheet2->setCellValue('C3', '20');

        // Create directory if it doesn't exist
        $dir = dirname($this->getTestFilePath());
        if (!is_dir($dir)) {
            mkdir($dir, 0755, true);
        }

        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($this->getTestFilePath());
    }

    public function test_json_conversion_single_sheet()
    {
        $result = ExcelTo::json($this->getTestFilePath());
        $data = json_decode($result, true);

        $this->assertIsArray($data);
        $this->assertCount(2, $data);
        $this->assertEquals([
            'First Name' => 'John',
            'Last Name' => 'Doe',
            'Age' => '30',
            'Date' => '2023-01-01'
        ], $data[0]);
    }

    public function test_json_conversion_multiple_sheets()
    {
        $result = ExcelTo::json($this->getTestFilePath());
        $data = json_decode($result, true);

        $this->assertIsArray($data);
        $this->assertArrayHasKey('Sheet1', $data);
        $this->assertArrayHasKey('Sheet2', $data);

        $this->assertCount(2, $data['Sheet1']);
        $this->assertCount(2, $data['Sheet2']);
    }

    public function test_array_conversion()
    {
        $result = ExcelTo::array($this->getTestFilePath());

        $this->assertIsArray($result);
        $this->assertArrayHasKey('Sheet1', $result);
        $this->assertArrayHasKey('Sheet2', $result);

        $this->assertCount(2, $result['Sheet1']);
        $this->assertCount(2, $result['Sheet2']);
    }

    public function test_collection_conversion()
    {
        $result = ExcelTo::collection($this->getTestFilePath());

        $this->assertInstanceOf(Collection::class, $result);
        $this->assertTrue($result->has('Sheet1'));
        $this->assertTrue($result->has('Sheet2'));

        $this->assertCount(2, $result['Sheet1']);
        $this->assertCount(2, $result['Sheet2']);
    }

    public function test_handles_empty_sheets()
    {
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $emptyFilePath = __DIR__ . '/../test_files/test_empty.xlsx';
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($emptyFilePath);

        $result = ExcelTo::json($emptyFilePath);
        $data = json_decode($result, true);

        $this->assertIsArray($data);
        $this->assertEmpty($data);

        unlink($emptyFilePath);
    }

    public function test_handles_merged_cells()
    {
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        // Set up merged cells with header row
        $sheet->setCellValue('A1', 'Header1');
        $sheet->setCellValue('B1', 'Header2');
        $sheet->setCellValue('A2', 'Value 1');
        $sheet->setCellValue('B2', 'Value 2');
        $sheet->mergeCells('A1:B1');

        $mergedFilePath = __DIR__ . '/../test_files/test_merged.xlsx';
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($mergedFilePath);

        $result = ExcelTo::json($mergedFilePath);
        $data = json_decode($result, true);

        $this->assertIsArray($data);
        $this->assertCount(1, $data);
        $this->assertEquals('Value 1', $data[0]['Header1']);
        $this->assertEquals('Value 2', $data[0]['Header2']);

        unlink($mergedFilePath);
    }

    public function test_handles_date_formats()
    {
        $spreadsheet = new \PhpOffice\PhpSpreadsheet\Spreadsheet();
        $sheet = $spreadsheet->getActiveSheet();

        $sheet->setCellValue('A1', 'Date');
        $sheet->setCellValue('A2', '2023-01-01');
        $sheet->getStyle('A2')->getNumberFormat()->setFormatCode('yyyy-mm-dd');

        $dateFilePath = __DIR__ . '/../test_files/test_date.xlsx';
        $writer = new \PhpOffice\PhpSpreadsheet\Writer\Xlsx($spreadsheet);
        $writer->save($dateFilePath);

        $result = ExcelTo::json($dateFilePath);
        $data = json_decode($result, true);

        $this->assertIsArray($data);
        $this->assertCount(1, $data);
        $this->assertEquals('2023-01-01', $data[0]['Date']);

        unlink($dateFilePath);
    }

    public function test_handles_invalid_file()
    {
        $this->expectException(\InvalidArgumentException::class);
        ExcelTo::json('non_existent_file.xlsx');
    }

    public function test_handles_invalid_file_type()
    {
        $invalidFilePath = __DIR__ . '/../test_files/test_invalid.txt';
        file_put_contents($invalidFilePath, 'This is not an Excel file');

        $this->expectException(\InvalidArgumentException::class);
        ExcelTo::json($invalidFilePath);

        unlink($invalidFilePath);
    }
}
