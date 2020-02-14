# Spreadsheet

A simple Laravel abstraction to the PHPSpreadsheet (previously PHPExcel) library, ideal for writing XLSX (Excel) files. 

## Installing

```composer require osmaviation/spreadsheet```

## Using

### Via resolve
```php
$filename = 'my-filename.xlsx';

resolve('spreadsheet')->create($filename, function ($excel) {
    $excel->sheet('Worksheet', function ($sheet) {
        $sheet->fromArray([
            'Foo',
            'Bar',
        ], null, 'A1', true, false);
    });
});
```

### Via facade
```php
use OSMAviation\Spreadsheet\Facades\Spreadsheet;

$filename = 'my-filename.xlsx';

Spreadsheet::create($filename, function ($excel) {
    $excel->sheet('Worksheet', function ($sheet) {
        $sheet->fromArray([
            'Foo',
            'Bar',
        ], null, 'A1', true, false);
    });
});
```

### Via injection
```php
use OSMAviation\Spreadsheet\PhpSpreadsheet as Spreadsheet;

class MyController 
{
    public function store(Spreadsheet $spreadsheet)
    {
        $filename = 'my-filename.xlsx';
        
        $spreadsheet->create($filename, function ($excel) {
            $excel->sheet('Worksheet', function ($sheet) {
                $sheet->fromArray([
                    'Foo',
                    'Bar',
                ], null, 'A1', true, false);
            });
        });
    }
}
```

### Saving the spreadsheet

```php
$filename = 'some-folder/my-filename.xlsx';

Spreadsheet::create($filename, function ($excel) {
    $excel->sheet('Worksheet', function ($sheet) { 
        // $sheet will be a PhpOffice\PhpSpreadsheet\Worksheet\Worksheet instance
        $sheet->fromArray([
            'Foo',
            'Bar',
        ], null, 'A1', true, false);
    });
})->store('local');
```

### Loading a file

```php
Spreadsheet::load($filename, function ($excel) {
    $excel->sheet('Some existing sheet', function($sheet) {
        //
    });
});
```

You can also pass the disk name as the second argument to the `load` method to load files from a different file system.

```php
Spreadsheet::load($filename, 's3', function ($excel) {
    $excel->sheet('Some sheet', function($sheet) {
        //
    });
});
```

### Accessing a PHPSpreadsheet spreadsheet

The callback for the create method will provide an instance of `OSMAviation\Spreadsheet\Spreadsheet` which is a 
convenience layer for creating worksheets. You can access the vendor spreadsheet by using the `getSpreadsheet` method.

```php
Spreadsheet::create($filename, function ($excel) {
    $vendorSheet = $excel->getSpreadsheet(); // returns a PhpOffice\PhpSpreadsheet\Spreadsheet instance
})->store('local');
```