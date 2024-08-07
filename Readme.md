# Laravel Excel to JSON / Collection / Array

This Laravel package provides utilities for converting Excel files to JSON format, Laravel Collections, or PHP Arrays. It also supports reading data from multiple sheets within an Excel file.

## Installation

You can install the package via Composer:

```bash
composer require knackline/excel-to-x
```

## Usage

### JSON Conversion

To convert an Excel file to JSON format, use the `json` method of the `ExcelTo` class:

```php
use Knackline\ExcelTo\ExcelTo;

$jsonData = ExcelTo::json('path/to/your/excel_file.xlsx');
```

This will return a JSON-encoded string representing the Excel data. If the Excel file contains multiple sheets, the data will be organized by sheet names.

### Collection Conversion

To convert an Excel file to a Laravel Collection, use the `collection` method of the `ExcelTo` class:

```php
use Knackline\ExcelTo\ExcelTo;

$collection = ExcelTo::collection('path/to/your/excel_file.xlsx');
```

This will return a Laravel Collection containing the Excel data. When multiple sheets are present, each sheet's data will be a collection keyed by the sheet name.

### Array Conversion

To convert an Excel file to a PHP Array, use the `array` method of the `ExcelTo` class:

```php
use Knackline\ExcelTo\ExcelTo;

$arrayData = ExcelTo::array('path/to/your/excel_file.xlsx');
```

This will return a PHP array containing the Excel data. Similar to JSON and Collection, multiple sheets will be keyed by their names.

## Example

```php
use Knackline\ExcelTo\ExcelTo;

// Convert Excel to JSON
$jsonData = ExcelTo::json('path/to/your/excel_file.xlsx');

// Convert Excel to Collection
$collection = ExcelTo::collection('path/to/your/excel_file.xlsx');

// Convert Excel to Array
$arrayData = ExcelTo::array('path/to/your/excel_file.xlsx');
```

## Requirements

- PHP >= 8.2
- Laravel >= 8.x
- PhpSpreadsheet >= 1.20

## Author

- **RAJKUMAR SAMRA** - [rajkumarsamra@gmail.com](mailto:rajkumarsamra@gmail.com) ([Github](https://github.com/rjsamra))

## Contributing

Contributions are welcome! Feel free to submit pull requests or open an issue if you find any bugs or have any suggestions for improvements.

## License

This package is open-source software licensed under the [MIT license](https://opensource.org/licenses/MIT).


### Key Updates:
1. **Support for Multiple Sheets:** 
   - Described how the package handles multiple sheets, with data organized by sheet names.

2. **Array Conversion:**
   - Added a new section for array conversion, including an example of how to use the new `array` method.

3. **Clarified Output Format:**
   - Explained the structure of the data returned by each method, emphasizing the handling of single vs. multiple sheets.

Feel free to modify any section further if you have additional details or preferences for the README content.