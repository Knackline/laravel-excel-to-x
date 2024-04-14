# Laravel Excel to JSON / Collection

This Laravel package provides utilities for converting Excel files to JSON format or Laravel Collections.

## Installation

You can install the package via Composer:

```bash
composer require knackline/excel-to
```

## Usage

### JSON Conversion

To convert an Excel file to JSON format, use the `json` method of the `ExcelTo` class:

```php
use Knackline\ExcelTo\ExcelTo;

$jsonData = ExcelTo::json('path/to/your/excel_file.xlsx');
```

This will return an associative array representing the Excel data in JSON format.

### Collection Conversion

To convert an Excel file to a Laravel Collection, use the `collection` method of the `ExcelTo` class:

```php
use Knackline\ExcelTo\ExcelTo;

$collection = ExcelTo::collection('path/to/your/excel_file.xlsx');
```

This will return a Laravel Collection containing the Excel data.

## Example

```php
use Knackline\ExcelTo\ExcelTo;

// Convert Excel to JSON
$jsonData = ExcelTo::json('path/to/your/excel_file.xlsx');

// Convert Excel to Collection
$collection = ExcelTo::collection('path/to/your/excel_file.xlsx');
```

## Requirements

- PHP >= 8.2

## Author

- **RAJKUMAR SAMRA** - [rajkumarsamra@gmail.com](mailto:rajkumarsamra@gmail.com) ([Github](https://github.com/rjsamra))

## Contributing

Contributions are welcome! Feel free to submit pull requests or open an issue if you find any bugs or have any suggestions for improvements.

## License

This package is open-source software licensed under the [MIT license](https://opensource.org/licenses/MIT).

---

Feel free to customize the content further if needed. Let me know if there's anything else I can assist you with!