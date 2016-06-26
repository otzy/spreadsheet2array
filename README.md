# spreadsheet2array

Easy to use PHPExcel wrapper. If you are developing any sort of import from spreadsheets, this library will help you to import spreadsheet into php array.
You can import from xls, xlsx, ods or whatever else PHPExcel supports or will support.

## Installation

composer require otzy/spreadsheet2array

## How to use

### Import table with column names

```

\Otzy\Spreadsheet2Array::readTable($file_name, $type = 'auto', $sheet = false, $first_row = 0, $first_col = 0,
                                                                         $col_names = false, $check_col_names = false)
```
                                                                        
yes, many parameters, but usage is really simple.
Lets say you have a spreadsheet like the following one (column and row labels are not shown):

| one | two | three  |
|---|---|---|
| 1 | 2 | 3  |
| 4 | 5 | 6  |

your code will be:

```

$my_table = \Otzy\Spreadsheet2Array::readTable('path_to_your_file', 'auto', false, 0, 0, ['one', 'two', 'three']);

//$my_table will contain two dimensional array:
[
 ['one'=>1, 'two'=>2, 'three'=>3],
 ['one'=>1, 'two'=>2, 'three'=>3]
]

```


If you don't need all columns, list only those that you need:

```

$my_table = \Otzy\Spreadsheet2Array::readTable('path_to_your_file', 'auto', false, 0, 0, ['one', 'three']);

//$my_table will contain two dimensional array:
[
 ['one'=>1, 'three'=>3],
 ['one'=>1, 'three'=>3]
]

```

The presence of required field names is always checked. If at least one field is missing, Spreadsheet2ArrayException will be thrown.
 
###### complete description of all parameters: 

```
     * @param string $file_name name of file to read
     * @param string $type 'auto', 'csv', 'xls', 'xlsx', 'ods'
     *
     * @param bool|string|int $sheet Sheet Name/Index to read.
     * Index must be passed as an integer. Sheets are zero-based. I.e. first sheet has index=0
     * If false read active sheet. This is normally the sheet, that was active at the moment of Save in Excel or Open Office
     * Not applicable for csv
     *
     * @param int $first_row first row to read (zero based).
     *
     * @param int $first_col first column to read (zero based).
     *
     * @param string[]|bool $col_names array of field names or false to use $firstRow for field names
     *
     * @param bool $check_col_names Parameter works only when $col_names !== false.
     * If true - check that column names and the order of names are exactly the same as value and order of cells in the first row.<br>
     * If false - only check that all $col_names are represented in the first row
     *
```