<?php

namespace Knackline\ExcelTo;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

class ExcelTo
{

    public static function json($filePath)
    {
        $spreadsheet = IOFactory::load($filePath);
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

        $json = json_encode($jsonData);

        return $json;
    }

    public static function collection($filePath)
    {
        $spreadsheet = IOFactory::load($filePath);
        $sheetCount = $spreadsheet->getSheetCount();
        $collection = collect();

        foreach ($spreadsheet->getAllSheets() as $worksheet) {
            $sheetData = self::processSheet($worksheet);

            if ($sheetCount > 1) {
                $collection->put($worksheet->getTitle(), collect($sheetData));
            } else {
                $collection = collect($sheetData);
            }
        }

        return $collection;
    }

    private static function processSheet($worksheet)
    {
        $excelData = $worksheet->toArray();
        $header = array_shift($excelData);
        $sheetData = [];

        foreach ($excelData as $rowIndex => $row) {
            $rowData = [];
            foreach ($header as $colIndex => $columnName) {
                // Convert column index to letter (A, B, C, ...)
                $columnLetter = Coordinate::stringFromColumnIndex($colIndex + 1);
                $cellAddress = $columnLetter . ($rowIndex + 2);
                $cell = $worksheet->getCell($cellAddress);
                $value = $cell->getValue();
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
}
