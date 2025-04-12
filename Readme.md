# Laravel Excel To X

A Laravel package for converting Excel files to various formats (JSON, Array, Collection) with support for streaming large files.

## Features

- Convert Excel files to JSON
- Convert Excel files to PHP Arrays
- Convert Excel files to Laravel Collections
- Support for multiple sheets
- Handle merged cells
- Process date formats
- Stream large Excel files efficiently
- Memory-efficient processing of large datasets

## Installation

You can install the package via Composer:

```bash
composer require knackline/excel-to-x
```

## Usage

### Basic Conversion

```php
use Knackline\ExcelTo\ExcelTo;

// Convert to JSON
$json = ExcelTo::json('path/to/file.xlsx');

// Convert to Array
$array = ExcelTo::array('path/to/file.xlsx');

// Convert to Collection
$collection = ExcelTo::collection('path/to/file.xlsx');
```

### Streaming Large Files

For processing large Excel files efficiently:

```php
use Knackline\ExcelTo\ExcelTo;

// Process in chunks of 1000 rows (default)
$data = ExcelTo::stream('path/to/large_file.xlsx');

// Specify custom chunk size
$data = ExcelTo::stream('path/to/large_file.xlsx', 500);
```

The streaming method returns data in the same format as other conversion methods:

- For single sheet files: Array of rows
- For multiple sheet files: Associative array with sheet names as keys

## Response Format

### Single Sheet

```json
[
  {
    "First Name": "John",
    "Last Name": "Doe",
    "Age": "30",
    "Date": "2023-01-01"
  },
  {
    "First Name": "Jane",
    "Last Name": "Smith",
    "Age": "25",
    "Date": "2023-01-02"
  }
]
```

### Multiple Sheets

```json
{
  "Sheet1": [
    {
      "First Name": "John",
      "Last Name": "Doe",
      "Age": "30",
      "Date": "2023-01-01"
    }
  ],
  "Sheet2": [
    {
      "Product": "Laptop",
      "Price": "1000",
      "Stock": "10"
    }
  ]
}
```

## Features in Detail

### Memory Efficiency

The package uses chunking to process large files efficiently:

- Processes data in configurable chunks
- Prevents memory exhaustion with large datasets
- Maintains consistent output format

### Date Handling

- Automatically detects and formats date values
- Preserves date formats from Excel
- Converts Excel date values to standard formats

### Merged Cells

- Properly handles merged cells in Excel
- Preserves data integrity
- Maintains cell relationships

### Error Handling

- Validates file existence and readability
- Checks for valid Excel file types
- Provides clear error messages

## Testing

Run the test suite:

```bash
./vendor/bin/phpunit
```

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
