<?php

namespace Knackline\ExcelTo\Jobs;

use Illuminate\Bus\Queueable;
use Illuminate\Contracts\Queue\ShouldQueue;
use Illuminate\Foundation\Bus\Dispatchable;
use Illuminate\Queue\InteractsWithQueue;
use Illuminate\Queue\SerializesModels;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;

class ProcessExcelChunk implements ShouldQueue
{
    use Dispatchable, InteractsWithQueue, Queueable, SerializesModels;

    protected string $filePath;
    protected int $startRow;
    protected int $chunkSize;
    protected string $sheetName;

    public function __construct(string $filePath, int $startRow, int $chunkSize, string $sheetName)
    {
        $this->filePath = $filePath;
        $this->startRow = $startRow;
        $this->chunkSize = $chunkSize;
        $this->sheetName = $sheetName;
    }

    public function handle()
    {
        $spreadsheet = IOFactory::load($this->filePath);
        $worksheet = $spreadsheet->getSheetByName($this->sheetName);

        if (!$worksheet) {
            return [];
        }

        $data = [];
        $endRow = min($this->startRow + $this->chunkSize - 1, $worksheet->getHighestRow());
        $highestColumn = $worksheet->getHighestColumn();

        // Process data rows (excluding header)
        for ($row = $this->startRow; $row <= $endRow; $row++) {
            $rowData = [];
            for ($col = 'A'; $col <= $highestColumn; $col++) {
                $cellValue = $worksheet->getCell($col . $row)->getCalculatedValue();
                $rowData[] = $cellValue;
            }
            $data[] = $rowData;
        }

        return $data;
    }
}
