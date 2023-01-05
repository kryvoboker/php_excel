<?php

namespace OpenExcel;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Worksheet\Worksheet;
use PhpOffice\PhpSpreadsheet\Writer\Exception;

require_once __DIR__ . '/OpenExcel.php';

mb_internal_encoding('UTF-8');

class ParseExcel extends OpenExcel
{
    private ?Spreadsheet $new_excel = null;
    private Worksheet $new_worksheet;

    public function __construct(string $excel_file_one = '', string $excel_file_two = '')
    {
        parent::__construct($excel_file_one, $excel_file_two);

        $this->createNewExcel();
    }

    /**
     * @param array $parameters
     * @throws Exception
     */
    public function createExcelFromString(array $parameters) : void
    {
        [
            $sheet_name,
            $coordinate_of_info,
            $excel_folder,
        ] = $parameters;

        $worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
        $highest_row = $worksheet->getHighestRow();
        $this->count_something = 2;

        for ($row = 2; $row <= $highest_row; $row++) {
            $info_from_row = trim($worksheet->getCell($coordinate_of_info . $row)->getValue(), " \t\n\r_,.'");

            if (empty($info_from_row)) continue;

            $this->parseString($info_from_row);

            $this->count_something++;
        }

        $this->writeExcelFile($excel_folder, $this->new_excel);
        unset($highest_row);
    }

    /**
     * @param string $info_from_row
     */
    private function parseString(string $info_from_row) : void
    {
        $info_from_row = preg_replace('/([^\w\/])+/u', ' ', $info_from_row);
        $info_from_row = preg_replace('/\bтелефон\W?\s?|\bтел\W?\s?|\bт\W?\s/ui', '', $info_from_row);

        $arr_from_string = explode(' ', $info_from_row);
        $count_arr = count($arr_from_string);

        for ($index = 0; $index <= $count_arr; $index++) {
            if (preg_match('/\d{10,12}/', $arr_from_string[$index])) {
                $this->writeInNewExcel(trim($arr_from_string[$index]), 'B', true);
                $arr_from_string[$index] = '';
            }
        }

        $new_info_str = implode(' ', array_filter($arr_from_string));

        if (preg_match('/(\d+\D*(\s\b\w+\b)+)$/ui', $new_info_str, $matches)) {
            if (!empty($matches[0])) {
                $human = preg_replace('/\d+\w*\s.+\d+\w*\s|\d+\w*\s/ui', '', $matches[0]);
                $this->writeInNewExcel($human, 'C');
                $address = str_replace($human, '', $new_info_str);
                $new_info_str = str_replace([$address, $human], '', $new_info_str);
                $this->writeInNewExcel($address, 'A');
            }
        }

        if (!empty($new_info_str)) $this->writeInNewExcel($new_info_str, 'D');
    }

    private function createNewExcel() : void
    {
        $this->new_excel = new Spreadsheet();
        $this->new_worksheet = $this->new_excel->getActiveSheet();
    }

    /**
     * @param string $match_value
     * @param string $col_coordinate
     * @param bool $append
     */
    private function writeInNewExcel(string $match_value, string $col_coordinate, bool $append = false) : void
    {
        if ($append) {
            $val = $this->new_worksheet->getCell($col_coordinate . $this->count_something)->getValue();
            $this->new_worksheet->setCellValue($col_coordinate . $this->count_something, ltrim($val . ', ' . $match_value, ' ,'));
        } else {
            $this->new_worksheet->setCellValue($col_coordinate . $this->count_something, $match_value);
        }
    }

    public function __destruct()
    {
        parent::__destruct(); // TODO: Change the autogenerated stub

        if ($this->new_excel instanceof Spreadsheet) {
            $this->new_excel->disconnectWorksheets();
            unset($this->new_excel);
        }
    }
}