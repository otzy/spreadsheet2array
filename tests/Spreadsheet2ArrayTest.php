<?php
namespace Otzy;

class SpreadSheet2ArrayTest extends \PHPUnit_Framework_TestCase{

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

    public function sheetTypeProvider() {
        $types = array_keys($this->spreadsheets);
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
    public function testReadTableException($spreadsheet_type){
        //invalid field list
        $result = Spreadsheet2Array::readTable($this->spreadsheets[$spreadsheet_type], 'auto', 'sheet2', 0, 0, ['xxx', 'zzz'], true);
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