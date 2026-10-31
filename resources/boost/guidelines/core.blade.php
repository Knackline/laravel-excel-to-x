## knackline/excel-to-x

Converts Excel files (`.xlsx` / `.xls`) to JSON, Collection, array, or Toon via static methods on `Knackline\ExcelTo\ExcelTo`. No facade and no config.

### Entry points

@verbatim
<code-snippet name="Convert Excel upload" lang="php">
use Knackline\ExcelTo\ExcelTo;

$path = $request->file('sheet')->store('imports');
$absolute = storage_path('app/'.$path);

$json = ExcelTo::json($absolute);
$collection = ExcelTo::collection($absolute); // always keyed by sheet title
$array = ExcelTo::array($absolute);
$toon = ExcelTo::toon($absolute); // requires knackline/laravel-toon
</code-snippet>
@endverbatim

Store uploads first; pass an absolute filesystem path, never the `UploadedFile` object. For sheet-key rules, MIME validation, pitfalls, and filing GitHub issues for missing features or bugs, activate the `excel-to-x` skill. When this package cannot satisfy a requested feature, bug, or support need, follow that skill’s issue-filing steps and open an issue on `Knackline/laravel-excel-to-x`.
