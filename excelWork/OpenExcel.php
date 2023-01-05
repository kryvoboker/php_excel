<?php

namespace OpenExcel;

require_once __DIR__ . '../../vendor/autoload.php';

mb_internal_encoding('UTF-8');
//ini_set('max_execution_time', 600);

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class OpenExcel
{
    protected string $excel_file_one;
    protected string $excel_file_two;
    protected ?Spreadsheet $spreadsheet_one = null;
    protected ?Spreadsheet $spreadsheet_two = null;
    protected int $count_something = 0;

    public function __construct(string $excel_file_one = '', string $excel_file_two = '')
    {
        if (!empty($excel_file_one)) {
            $this->excel_file_one = $excel_file_one;
            $this->spreadsheet_one = IOFactory::load($excel_file_one);
        }

        if (!empty($excel_file_two)) {
            $this->excel_file_two = $excel_file_two;
            $this->spreadsheet_two = IOFactory::load($excel_file_two);
        }
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    protected function writeExcelFile(string $directory_path, Spreadsheet $spreadsheet_one = null, ?Spreadsheet $spreadsheet_two = null) : void
    {
        if ($spreadsheet_one) {
            $writer_file_one = IOFactory::createWriter($spreadsheet_one, "Xlsx");

            $array_of_file_one = explode('/', $this->excel_file_one);
            $file_name_one = end($array_of_file_one);

            $writer_file_one->save($directory_path . 'new_' . $file_name_one);

            echo('new_' . $file_name_one . ' - успішно записаний' . '<br/>');
        }

        if ($spreadsheet_two) {
            $array_of_file_two = explode('/', $this->excel_file_two);
            $file_name_two = end($array_of_file_two);

            $writer_file_two = IOFactory::createWriter($spreadsheet_two, "Xlsx");
            $writer_file_two->save($directory_path . 'new_' . $file_name_two);

            echo('new_' . $file_name_two . ' - успішно записаний' . '<br/>');
        }
    }

    public function __destruct()
    {
        if ($this->spreadsheet_one instanceof Spreadsheet) {
            $this->spreadsheet_one->disconnectWorksheets();
            unset($this->spreadsheet_one);
        }

        if ($this->spreadsheet_two instanceof Spreadsheet) {
            $this->spreadsheet_two->disconnectWorksheets();
            unset($this->spreadsheet_two);
        }
    }
}