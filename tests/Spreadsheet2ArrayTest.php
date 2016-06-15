<?php
namespace Otzy;

class SpreadSheet2ArrayTest extends \PHPUnit_Framework_TestCase{

    /**
     * TODO test csv
     *
     * @var array
     */
    public $spreadsheets = array(
        'xlsx' => __DIR__. '/data/test.xlsx',
        'ods' => __DIR__. '/data/test.ods',
        'auto' => __DIR__. '/data/test.xlsx'
    );

    public $sheets = [
        'sheet1'=>[
            ['a', 'b', 'c', 'd', null, null],
            ['aa', 'bb', 'cc', 'dd', null, null],
            [1,	2,	3,	4, null, null],
            ['one',	'two', 'three', 'four', null, null],
            ['2016-01-02 12:13:13',	'2016-01-04 12:13:13',	'2016-12-07 12:13:13',	'2016-01-02 12:13:13', null, 'x']
        ],
        'sheet2'=>[['xxx','yyy','zzz'],
            [1, 2, 3],
            [4, 5, 6]
        ]
    ];

    /**
     * two dimensional version of array_slice
     *
     * @param array $arr normal non hash array
     * @param int $first_row zero based
     * @param int $first_col zero based
     * @param int $max_rows >0
     * @param int $max_cols >0
     * @return array
     */
    private static function getSubArray($arr, $first_row, $first_col, $max_rows, $max_cols){
        $rows = array_slice($arr, $first_row, $max_rows);
        array_walk($rows, function(&$row)use($first_col, $max_cols){$row = array_slice($row, $first_col, $max_cols);});
        return $rows;
    }

    public function sheetTypeProvider() {
        $types = array_keys($this->spreadsheets);

        unset($types[2]); //xlsx and ods is enough. We are not going to test PHPExcel here

        //convert to array of arrays
        array_walk($types, function (&$v) {
            $v = [$v];
        });

        return $types;
    }


    public function testGetSheetByIndex(){
        /* @var \PHPExcel_Worksheet $sheet */
        $sheet = Spreadsheet2Array::invokePrivate('getSheet', array($this->spreadsheets['xlsx'], 'xlsx', 0));
        $this->assertEquals('sheet1', $sheet->getTitle());

        $sheet = Spreadsheet2Array::invokePrivate('getSheet', array($this->spreadsheets['xlsx'], 'auto', 1));
        $this->assertEquals('sheet2', $sheet->getTitle());
    }

    public function testGetSheetByName(){
        /* @var \PHPExcel_Worksheet $sheet */
        $sheet = Spreadsheet2Array::invokePrivate('getSheet', array($this->spreadsheets['xlsx'], 'xlsx', 'sheet1'));
        $this->assertEquals('sheet1', $sheet->getTitle());

        $sheet = Spreadsheet2Array::invokePrivate('getSheet', array($this->spreadsheets['xlsx'], 'auto', 'sheet2'));
        $this->assertEquals('sheet2', $sheet->getTitle());

        $sheet = Spreadsheet2Array::invokePrivate('getSheet', array($this->spreadsheets['xlsx'], 'auto', 'sheet3'));
        $this->assertNull($sheet);
    }

    public function testGetActiveSheet(){
        /* @var \PHPExcel_Worksheet $sheet */
        $sheet = Spreadsheet2Array::invokePrivate('getSheet', array($this->spreadsheets['ods'], 'auto'));
        $this->assertEquals('sheet2', $sheet->getTitle());
    }

    /**
     * @dataProvider sheetTypeProvider
     * @param string $spreadsheet_type
     */
    public function testReadRow($spreadsheet_type){
        /* @var \PHPExcel_Worksheet $sheet */
        $sheet = Spreadsheet2Array::invokePrivate('getSheet', array($this->spreadsheets[$spreadsheet_type], 'auto', 'sheet2'));
        $test_array = $this->sheets['sheet2'];

        $i = 0;
        foreach ($sheet->getRowIterator() as $rowObj){

            //test entire row
            $row_array = Spreadsheet2Array::invokePrivate('readRow', array($rowObj, 0));
            $this->assertEquals($test_array[$i], $row_array);

            //test from N-th cell
            $n = 1;
            $row_array = Spreadsheet2Array::invokePrivate('readRow', array($rowObj, $n));
            $this->assertEquals(array_slice($test_array[$i],1), $row_array);
            $i++;
        }
    }

    /**
     * @dataProvider sheetTypeProvider
     * @param string $spreadsheet_type
     */
    public function testReadEntireSheet($spreadsheet_type){

        //we perform test on a bigger table with empty cells - it's a "sheet1" in our sample files
        $sheet = Spreadsheet2Array::getSheet($this->spreadsheets[$spreadsheet_type], 'auto', 'sheet1');
        $test_array = $this->sheets['sheet1'];
        $result = Spreadsheet2Array::readHeadlessTable($sheet, 0, 0);

        //the last row in sheet1 contains date in Excel format. Lets convert it to string
        for($i = 0; $i<4; $i++){
            $result[4][$i] = date('Y-m-d H:i:s', Spreadsheet2Array::excelDate2Timestamp($result[4][$i]));
        }

        $this->assertEquals($test_array, $result, __FUNCTION__ . ' failed. If you an error in date/time in the last row, probably it\'s a time zone issue.');
    }

    /**
     * @dataProvider sheetTypeProvider
     * @param string $spreadsheet_type
     */
    public function testReadPartOfSheet($spreadsheet_type){

        //we perform test on a bigger table with empty cells - it's a "sheet1" in our sample files
        $sheet = Spreadsheet2Array::getSheet($this->spreadsheets[$spreadsheet_type], 'auto', 'sheet1');
        $test_array = self::getSubArray($this->sheets['sheet1'], 1, 1, 2, 2);
        $result = Spreadsheet2Array::readHeadlessTable($sheet, 1, 1, 2, 2);

        $this->assertEquals($test_array, $result, __FUNCTION__ . ' failed. If you an error in date/time in the last row, probably it\'s a time zone issue.');
    }

    /**
     * @dataProvider sheetTypeProvider
     * @param string $spreadsheet_type
     */
    public function testReadTable($spreadsheet_type){

        //valid field list
        $result = Spreadsheet2Array::readTable($this->spreadsheets[$spreadsheet_type], 'auto', 'sheet2', 0, 0, ['xxx','yyy','zzz'], true);
        $test_array = $this->plain2HashArray($this->sheets['sheet2']);
        $this->assertEquals($test_array, $result);

        //shorter field list
        $result = Spreadsheet2Array::readTable($this->spreadsheets[$spreadsheet_type], 'auto', 'sheet2', 0, 0, ['xxx', 'zzz'], false);
        //remove 'yyy'-s from $test_array
        array_walk($test_array, function(&$v){unset($v['yyy']);});
        $this->assertEquals($test_array, $result);
    }

    /**
     * @dataProvider sheetTypeProvider
     * @param string $spreadsheet_type
     * @expectedException \Otzy\Spreadsheet2ArrayException
     */
    public function testReadTableExceptionExtraFields($spreadsheet_type){
        //invalid field list
        Spreadsheet2Array::readTable($this->spreadsheets[$spreadsheet_type], 'auto', 'sheet2', 0, 0, ['xxx', 'zzz'], true);
    }

    /**
     * @dataProvider sheetTypeProvider
     * @param string $spreadsheet_type
     * @expectedException \Otzy\Spreadsheet2ArrayException
     */
    public function testReadTableExceptionMissingFields($spreadsheet_type){
        //invalid field list
        Spreadsheet2Array::readTable($this->spreadsheets[$spreadsheet_type], 'auto', 'sheet2', 0, 0, ['xxx', 'yyy' ,'zzz', 'ccc'], false);
    }

    /**
     * @dataProvider sheetTypeProvider
     * @param string $spreadsheet_type
     * @expectedException \Otzy\Spreadsheet2ArrayException
     */
    public function testReadTableExceptionWrongFields($spreadsheet_type){
        //invalid field list
        Spreadsheet2Array::readTable($this->spreadsheets[$spreadsheet_type], 'auto', 'sheet2', 0, 0, ['xxx1', 'yyy', 'zzz'], true);
    }

    /**
     * converts plain array into hash with elements of first rows as hashes (field names)
     *
     * @param $array
     * @return array
     */
    protected function plain2HashArray($array){
        $result = [];
        $fields = array_shift($array);
        foreach ($array as $row){
            $result[] = array_combine($fields, $row);
        }

        return $result;
    }

}
