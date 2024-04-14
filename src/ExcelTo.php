<?php

namespace Knackline\ExcelTo;

use Maatwebsite\Excel\Facades\Excel;

class ExcelTo
{

    public static function json($filePath)
    {
        $excelData = Excel::toArray([], $filePath)[0];
        $header = array_shift($excelData);

        $jsonData = [];
        foreach ($excelData as $row) {
            $rowData = [];
            foreach ($header as $key => $columnName) {
                $rowData[$columnName] = $row[$key];
            }
            $jsonData[] = $rowData;
        }

        $json = json_encode($jsonData);

        $decodedJson = json_decode($json, true);

        return $decodedJson;
    }

    public static function collection($filePath)
    {
        $excelData = Excel::toArray([], $filePath)[0];
        $header = array_shift($excelData);

        $collection = collect();
        foreach ($excelData as $row) {
            $rowData = [];
            foreach ($header as $key => $columnName) {
                $rowData[$columnName] = $row[$key];
            }
            $collection->push($rowData);
        }

        return $collection;
    }
}
