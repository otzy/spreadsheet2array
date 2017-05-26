<?php
/*
 * The MIT License (MIT)
 *
 * Copyright (c) 2016 Evgeny Mazovetskiy
 *
 * Permission is hereby granted, free of charge, to any person obtaining a copy
 * of this software and associated documentation files (the "Software"), to deal
 * in the Software without restriction, including without limitation the rights
 * to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 * copies of the Software, and to permit persons to whom the Software is
 * furnished to do so, subject to the following conditions:
 *
 * The above copyright notice and this permission notice shall be included in all
 * copies or substantial portions of the Software.
 *
 * THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 * IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 * FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 * AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 * LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 * OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 * SOFTWARE.
 */

namespace Otzy;

class Spreadsheet2Array
{
    /**
     * Read file content to array
     * Parameters:
     * @param string $file_name name of file to read
     * @param string $type 'auto', 'csv', 'xls', 'xlsx', 'ods'
     * @param bool|string|int $sheet Sheet Name/Index to read.
     *                        Index must be passed as an <b>integer</b>. Sheets are zero-based. I.e. first sheet has index=0<br>
     *                        If <b>false</b> read active sheet. This is normally the sheet, that was active at the moment of Save in Excel or Open Office<br>
     *                        Not applicable for csv
     * @param int $first_row first row to read (zero based).
     * @param int $first_col first column to read (zero based).
     * @param string[]|bool $col_names array of field names or false to use $firstRow for field names
     * @param bool $check_col_names Parameter works only when <b>$col_names !== false</b>.<br>
     *                              If <b>true</b> - check that column names and the order of names are exactly the same as value and order of cells in the first row.<br>
     *                              If <b>false</b> - only check that all $col_names are represented in the first row
     *
     * @return array
     * @throws
     */
    public static function readTable($file_name, $type = 'auto', $sheet = false, $first_row = 0, $first_col = 0,
                                     $col_names = false, $check_col_names = false)
    {
        $objSheet = self::getSheet($file_name, $type, $sheet);

        $result = array();

        $row_number = 0;
        $header_index = array();
        foreach ($objSheet->getRowIterator($first_row + 1) as $row) {
            /* @var \PHPExcel_Worksheet_Row $row */
            $row_number++;

            if ($row_number == 1) {
                //this is a header. Read it to array
                $header = self::readRow($row, $first_col);
                
                if (!is_array($col_names)){
                    $col_names = $header;
                }

                if ($check_col_names && $header != $col_names) {
                    throw new Spreadsheet2ArrayException('Fields in the spreadsheet differ from the required ones.');
                }

                if (!$check_col_names) {
                    $diff = array_diff($col_names, $header);
                    if (count($diff) > 0) {
                        throw new Spreadsheet2ArrayException('Fields are missing in the input file: ' . implode(', ', $diff));
                    }
                }

                $header_index = array_flip($header);
                continue;
            }

            $cells = self::readRow($row, $first_col);
            $result[] = static::filterRow($cells, $col_names, $header_index);
        }

        return $result;
    }

    private static function filterRow($cells, $col_names, $header_index){
        $cells_filtered = array();
        foreach ($col_names as $field_name) {
            if (array_key_exists($header_index[$field_name], $cells)){
                $cells_filtered[$field_name] = $cells[$header_index[$field_name]];
            } else {
                //if the row is shorter than other rows we treat missing cells as empty strings
                $cells_filtered[$field_name] = '';
            }
        }

        return $cells_filtered;
    }

    /**
     * @param string $file_name
     * @param string $type
     * @param bool|string|int $sheet Sheet Name|Index to read.
     *                        Index must be passed as an <b>integer</b>. Sheets are zero-based. I.e. first sheet has index=0<br>
     *                        If <b>false</b> read active sheet. This is normally the sheet, that was active at the moment of Save in Excel or Open Office<br>
     *                        Not applicable for csv
     * @param int $firstRow
     * @param int $firstCol
     * @param int $maxRows how many rows to read
     * @param int $maxCols how many columns to read
     *
     * @return array[]
     */
    public static function readRange($file_name, $type = 'auto', $sheet = false, $firstRow = 0, $firstCol = 0,
                                     $maxRows = 0, $maxCols = 0)
    {
        $objSheet = self::getSheet($file_name, $type, $sheet);

        /* @var array[] $result */
        $result = array();
        $row_count = 0;
        $longest_row_length = 0;

        foreach ($objSheet->getRowIterator($firstRow + 1) as $row) {
            /* @var \PHPExcel_Worksheet_Row $row */
            $row_count++;
            if ($maxRows > 0 && $row_count > $maxRows) {
                break;
            }

            //Read cells, starting from $firstCol
            $cells = self::readRow($row, $firstCol);
            if ($maxCols > 0 && count($cells) > $maxCols) {
                $cells = array_slice($cells, 0, $maxCols);
            }

            $result[] = $cells;
            $longest_row_length = max($longest_row_length, count($cells));
        }

        //align the number of elements in each row. I.e add missing elements
        foreach ($result as &$v) {
            for ($i = count($v); $i < $longest_row_length; $i++) {
                $v[] = null;
            }
        }
        unset($v);

        return $result;
    }

    /**
     * @param string $file_name
     * @param string $type
     * @param bool|string|int $sheet
     * @return \PHPExcel_Worksheet
     * @throws Spreadsheet2ArrayException
     * @throws \PHPExcel_Exception
     */
    public static function getSheet($file_name, $type = 'auto', $sheet = false)
    {
        if ($type == 'auto') {
            $objPHPExcel = \PHPExcel_IOFactory::load($file_name);
        } else {
            $objReader = self::getReader($type);
            if ($type != 'csv' && is_string($sheet)) {
                //load only the needed sheet, just to save memory. This works only if $sheet passed as a string
                $objReader->setLoadSheetsOnly($sheet);
            }
            $objPHPExcel = $objReader->load($file_name);
        }

        if (is_string($sheet)) {
            $objSheet = $objPHPExcel->getSheetByName($sheet);
        } elseif (is_int($sheet)) {
            $objSheet = $objPHPExcel->getSheet($sheet);
        } elseif ($sheet === false) {
            $objSheet = $objPHPExcel->getActiveSheet();
        } else {
            throw new Spreadsheet2ArrayException('Invalid type of the $sheet parameter. It must be string, int or boolean false'); //@codeCoverageIgnore
        }

        return $objSheet;
    }

    /**
     * @param $type
     * @return \PHPExcel_Reader_CSV|\PHPExcel_Reader_Excel2007|\PHPExcel_Reader_Excel5|\PHPExcel_Reader_OOCalc
     * @throws \PHPExcel_Exception
     * @codeCoverageIgnore
     */
    private static function getReader($type)
    {
        switch ($type) {
            case 'xls':
                $objReader = new \PHPExcel_Reader_Excel5();
                break;
            case 'xlsx':
                $objReader = new \PHPExcel_Reader_Excel2007();
                break;
            case 'ods':
                $objReader = new \PHPExcel_Reader_OOCalc();
                break;
            case 'csv':
                $objReader = new \PHPExcel_Reader_CSV();
                break;
            default:
                throw new \PHPExcel_Exception("Unsupported spreadsheet type $type");
        }

        if ($type != 'csv') {
            //we don't need format
            $objReader->setReadDataOnly(true);
        }

        return $objReader;
    }

    /**
     * @param \PHPExcel_Worksheet_Row $row
     * @param int $start_cell
     * @return array
     */
    private static function readRow(\PHPExcel_Worksheet_Row $row, $start_cell = 0)
    {
        $result = array();
        $cell_iterator = $row->getCellIterator();
        $cell_iterator->setIterateOnlyExistingCells(false);
        $cell_index = -1; //actually the first is 0, but we increment it in the beginning of the loop
        $not_null_right_index = -1;
        foreach ($cell_iterator as $cell) {
            /* @var \PHPExcel_Cell $cell */
            $cell_index++;

            if ($cell_index < $start_cell) {
                continue;
            }

            $result[] = $cell->getValue();

            //remember position of the rightmost not null element, in order to remove null values later without additional loop
            if ($result[count($result) - 1] !== null) {
                $not_null_right_index = count($result) - 1;
            }

        }

        //delete all rightmost null elements (Excel issue)
        $result = array_slice($result, 0, $not_null_right_index + 1);

        return $result;
    }

    /**
     * for test
     *
     * @param string $method_name
     * @param array $arguments
     * @return mixed
     *
     * @codeCoverageIgnore
     */
    public static function invokePrivate($method_name, $arguments)
    {
        $method = (new \ReflectionClass(__CLASS__))->getMethod($method_name);
        $method->setAccessible(true);
        return $method->invokeArgs(null, $arguments);
    }

    /**
     * Converts internal representation of date to unix timestamp
     *
     * @param float $excel_value
     * @return int
     *
     * @codeCoverageIgnore
     */
    public static function excelDate2Timestamp($excel_value)
    {
        return round($excel_value * 86400, 0) - 2209165200;
    }

}
