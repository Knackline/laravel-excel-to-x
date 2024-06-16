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
        $worksheet = $spreadsheet->getActiveSheet();
        $excelData = $worksheet->toArray();
        $header = array_shift($excelData);

        $jsonData = [];
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
            $jsonData[] = $rowData;
        }

        $json = json_encode($jsonData);
        $decodedJson = json_decode($json, true);

        return $decodedJson;
    }

    public static function collection($filePath)
    {
        $spreadsheet = IOFactory::load($filePath);
        $worksheet = $spreadsheet->getActiveSheet();
        $excelData = $worksheet->toArray();
        $header = array_shift($excelData);

        $collection = collect();
        foreach ($excelData as $rowIndex => $row) {
            $rowData = [];
            foreach ($header as $colIndex => $columnName) {
                // Convert column index to letter (A, B, C, ...)
                $columnLetter = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::stringFromColumnIndex($colIndex + 1);
                $cellAddress = $columnLetter . ($rowIndex + 2);
                $cell = $worksheet->getCell($cellAddress);
                $value = $cell->getValue();
                $isDate = Date::isDateTime($cell);

                if ($isDate && is_numeric($value)) {
                    $value = Date::excelToDateTimeObject($value)->format('d/m/Y');
                }

                $rowData[$columnName] = $value;
            }
            $collection->push($rowData);
        }

        return $collection;
    }
}
