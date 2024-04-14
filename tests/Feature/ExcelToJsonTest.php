<?php

namespace Knackline\ExcelTo\Tests\Feature;

use Knackline\ExcelTo\ExcelTo;
use PHPUnit\Framework\TestCase;

class ExcelToTest extends TestCase
{
    /** @test */
    public function it_can_convert_excel_to_json()
    {
        $excelFile = __DIR__ . '/../../public/test.xlsx'; // Adjust the path accordingly
        $json = ExcelTo::json($excelFile);

        $this->assertIsString($json);
    }
}
