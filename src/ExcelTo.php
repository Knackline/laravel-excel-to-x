<?php

namespace Knackline\ExcelTo;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use Illuminate\Support\Collection;
use finfo;

class ExcelTo
{
    public static function json(string $filePath): string
    {
        $spreadsheet = self::loadSpreadsheet($filePath);
        $sheetCount = $spreadsheet->getSheetCount();
        $jsonData = [];

        foreach ($spreadsheet->getAllSheets() as $worksheet) {
            $sheetData = self::processSheet($worksheet);

            if ($sheetCount > 1) {
                $jsonData[$worksheet->getTitle()] = $sheetData;
            } else {
                $jsonData = $sheetData;
            }
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

    private static function processSheet($worksheet): array
    {
        $excelData = $worksheet->toArray(null, true, true, true); // Load with nulls
        $mergedCells = $worksheet->getMergeCells(); // Get merged cells info
        $header = array_shift($excelData);
        $sheetData = [];

        foreach ($excelData as $rowIndex => $row) {
            $rowData = [];

            foreach ($header as $colIndex => $columnName) {
                $columnLetter = $colIndex; // PhpSpreadsheet uses lettered indexes
                $cellAddress = $columnLetter . ($rowIndex + 2);

                // Check if the cell is part of a merged range
                $value = $worksheet->getCell($cellAddress)->getCalculatedValue();
                foreach ($mergedCells as $range) {
                    if (self::isCellInRange($cellAddress, $range)) {
                        $value = $worksheet->getCell(explode(':', $range)[0])->getCalculatedValue();
                        break;
                    }
                }

                // Handle date values
                $cell = $worksheet->getCell($cellAddress);
                $isDate = Date::isDateTime($cell);

                if ($isDate && is_numeric($value)) {
                    $value = Date::excelToDateTimeObject($value)->format('d/m/Y');
                }

                $rowData[$columnName] = $value;
            }

            $sheetData[] = $rowData;
        }

        return $sheetData;
    }

    private static function isCellInRange(string $cellAddress, string $range): bool
    {
        [$startCell, $endCell] = explode(':', $range);
        [$startCol, $startRow] = Coordinate::coordinateFromString($startCell);
        [$endCol, $endRow] = Coordinate::coordinateFromString($endCell);
        [$col, $row] = Coordinate::coordinateFromString($cellAddress);

        return $row >= $startRow && $row <= $endRow && $col >= $startCol && $col <= $endCol;
    }
}
