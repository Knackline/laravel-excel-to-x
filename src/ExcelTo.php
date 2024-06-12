<?php

namespace Knackline\ExcelTo;

use Maatwebsite\Excel\Facades\Excel;
use PhpOffice\PhpSpreadsheet\Shared\Date;

class ExcelTo
{

    public static function json($filePath)
    {
        // Read the Excel data into an array
        $excelData = Excel::toArray([], $filePath)[0];

        // Extract the header from the data
        $header = array_shift($excelData);

        // Initialize an empty array to hold the processed data
        $jsonData = [];

        // Iterate through each row of data
        foreach ($excelData as $row) {
            $rowData = [];
            foreach ($header as $key => $columnName) {
                // Check if the column name is "Date" and the value is numeric
                if (($columnName === 'date' || $columnName === 'Date') && is_numeric($row[$key])) {
                    // Convert Excel date to a formatted date string
                    $rowData[$columnName] = Date::excelToDateTimeObject($row[$key])->format('Y-m-d'); // Format as needed
                } else {
                    $rowData[$columnName] = $row[$key];
                }
            }
            $jsonData[] = $rowData;
        }

        // Encode the array into JSON format
        $json = json_encode($jsonData);

        // Decode the JSON back to an array for returning
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
                // Check if the column name is "Date" and the value is numeric
                if (($columnName === 'date' || $columnName === 'Date') && is_numeric($row[$key])) {
                    $rowData[$columnName] = Date::excelToDateTimeObject($row[$key])->format('Y-m-d'); // Format as needed
                } else {
                    $rowData[$columnName] = $row[$key];
                }
            }
            $collection->push($rowData);
        }

        return $collection;
    }
}
