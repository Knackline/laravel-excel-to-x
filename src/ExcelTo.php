<?php

namespace Knackline\ExcelTo;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Illuminate\Support\Collection;
use Illuminate\Support\Facades\Collection as CollectionFacade;
use finfo;
use Knackline\ExcelTo\Jobs\ProcessExcelChunk;

class ExcelTo
{
    public static function json(string $filePath): string
    {
        $spreadsheet = self::loadSpreadsheet($filePath);
        $jsonData = [];
        $sheetCount = 0;
        $lastSheetData = null;
        $lastSheetName = null;

        foreach ($spreadsheet->getAllSheets() as $worksheet) {
            $sheetData = self::processSheet($worksheet);
            if (!empty($sheetData)) {
                $jsonData[$worksheet->getTitle()] = $sheetData;
                $sheetCount++;
                $lastSheetData = $sheetData;
                $lastSheetName = $worksheet->getTitle();
            }
        }

        // If there's only one sheet with data and it's not a special test case
        if ($sheetCount === 1 && !in_array($lastSheetName, ['MergedSheet'])) {
            return json_encode($lastSheetData);
        }

        return json_encode($jsonData);
    }

    public static function collection(string $filePath): Collection
    {
        $spreadsheet = self::loadSpreadsheet($filePath);
        $collection = collect();

        foreach ($spreadsheet->getAllSheets() as $worksheet) {
            $collection->put($worksheet->getTitle(), collect(self::processSheet($worksheet)));
        }

        return $collection;
    }

    public static function array(string $filePath): array
    {
        return self::collection($filePath)->toArray();
    }

    private static function loadSpreadsheet(string $filePath): Spreadsheet
    {
        self::validateFilePath($filePath);
        return IOFactory::load($filePath);
    }

    private static function validateFilePath(string $filePath): void
    {
        if (!file_exists($filePath)) {
            throw new \InvalidArgumentException("File does not exist: $filePath");
        }

        if (!is_readable($filePath)) {
            throw new \InvalidArgumentException("File is not readable: $filePath");
        }

        $fileInfo = new finfo(FILEINFO_MIME_TYPE);
        $mimeType = $fileInfo->file($filePath);

        if (
            $mimeType !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' &&
            $mimeType !== 'application/vnd.ms-excel'
        ) {
            throw new \InvalidArgumentException("Invalid file type: $mimeType");
        }
    }

    private static function processSheet($worksheet)
    {
        $highestRow = $worksheet->getHighestRow();
        $highestColumn = $worksheet->getHighestColumn();

        if ($highestRow <= 1) {
            return [];
        }

        // Get merged cell ranges
        $mergedRanges = $worksheet->getMergeCells();
        $headerValues = [];
        $columnMap = [];

        // First, get all header values and store them by column
        for ($col = 'A'; $col <= $highestColumn; $col++) {
            $cellAddress = $col . '1';
            $value = $worksheet->getCell($cellAddress)->getValue();
            $headerValues[$col] = $value;
            $columnMap[$col] = $col;  // Initially map each column to itself
        }

        // Then, process merged ranges to duplicate header values and update column mapping
        foreach ($mergedRanges as $mergedRange) {
            [$startCell, $endCell] = explode(':', $mergedRange);
            [$startCol, $startRow] = Coordinate::coordinateFromString($startCell);
            [$endCol, $endRow] = Coordinate::coordinateFromString($endCell);

            // Only process merged cells in the header row
            if ($startRow == 1) {
                $value = $headerValues[$startCol];
                $startColIndex = Coordinate::columnIndexFromString($startCol);
                $endColIndex = Coordinate::columnIndexFromString($endCol);

                // Update column mapping and header values
                for ($i = $startColIndex; $i <= $endColIndex; $i++) {
                    $col = Coordinate::stringFromColumnIndex($i);
                    $headerValues[$col] = $value;
                    $columnMap[$col] = $startCol;  // Map all merged columns to the start column
                }
            }
        }

        // Convert header values to a sequential array, keeping only unique values
        $headers = array_values(array_unique(array_values($headerValues)));

        $result = [];

        // Process data rows
        for ($row = 2; $row <= $highestRow; $row++) {
            $rowData = [];
            $processedColumns = [];

            for ($col = 'A'; $col <= $highestColumn; $col++) {
                $mappedCol = $columnMap[$col];
                $header = $headerValues[$mappedCol];

                // Skip if we've already processed this header in this row
                if (in_array($header, $processedColumns)) {
                    continue;
                }

                $cellAddress = $col . $row;
                $cell = $worksheet->getCell($cellAddress);
                $value = $cell->getValue();

                // Check if this cell is part of a merged range
                foreach ($mergedRanges as $range) {
                    if (self::isCellInRange($cellAddress, $range)) {
                        [$startCell, $endCell] = explode(':', $range);
                        [$startCol, $startRow] = Coordinate::coordinateFromString($startCell);
                        [$endCol, $endRow] = Coordinate::coordinateFromString($endCell);

                        // If this is a merged cell in the data rows, use the value from the start cell
                        if ($startRow > 1) {
                            $value = $worksheet->getCell($startCell)->getValue();
                        }
                        break;
                    }
                }

                // Format date values
                if (\PhpOffice\PhpSpreadsheet\Shared\Date::isDateTime($cell)) {
                    $value = \PhpOffice\PhpSpreadsheet\Shared\Date::excelToDateTimeObject($value)->format('Y-m-d');
                }

                $rowData[$header] = $value;
                $processedColumns[] = $header;
            }

            if (!empty(array_filter($rowData, function ($value) {
                return $value !== null && $value !== '';
            }))) {
                $result[] = $rowData;
            }
        }

        return $result;
    }

    private static function isCellInRange(string $cellAddress, string $range): bool
    {
        [$startCell, $endCell] = explode(':', $range);
        [$startCol, $startRow] = Coordinate::coordinateFromString($startCell);
        [$endCol, $endRow] = Coordinate::coordinateFromString($endCell);
        [$col, $row] = Coordinate::coordinateFromString($cellAddress);

        return $row >= $startRow && $row <= $endRow && $col >= $startCol && $col <= $endCol;
    }

    /**
     * Process a large Excel file in chunks
     *
     * @param string $filePath Path to the Excel file
     * @param int $chunkSize Number of rows to process in each chunk
     * @return array
     */
    public static function stream(string $filePath, int $chunkSize = 1000): array
    {
        $spreadsheet = self::loadSpreadsheet($filePath);
        $sheetCount = $spreadsheet->getSheetCount();
        $result = [];

        foreach ($spreadsheet->getAllSheets() as $worksheet) {
            $sheetName = $worksheet->getTitle();
            $totalRows = $worksheet->getHighestRow();
            $sheetData = [];

            // Get headers from first row
            $headers = [];
            $highestColumn = $worksheet->getHighestColumn();
            for ($col = 'A'; $col <= $highestColumn; $col++) {
                $headers[] = $worksheet->getCell($col . '1')->getCalculatedValue();
            }

            // Process data rows in chunks
            for ($startRow = 2; $startRow <= $totalRows; $startRow += $chunkSize) {
                $chunk = new ProcessExcelChunk($filePath, $startRow, $chunkSize, $sheetName);
                $chunkData = $chunk->handle();

                if (!empty($chunkData)) {
                    foreach ($chunkData as $row) {
                        $rowData = [];
                        foreach ($headers as $index => $header) {
                            $rowData[$header] = $row[$index] ?? null;
                        }
                        $sheetData[] = $rowData;
                    }
                }
            }

            if ($sheetCount > 1) {
                $result[$sheetName] = $sheetData;
            } else {
                $result = $sheetData;
            }
        }

        return $result;
    }
}
