<?php /** @noinspection PhpComposerExtensionStubsInspection */

require_once 'vendor/autoload.php';

mb_internal_encoding('UTF-8');
ini_set('max_execution_time', 600);

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class ExcelWork
{
    private string $excel_file_one;
    private string $excel_file_two;
    private ?Spreadsheet $spreadsheet_one = null;
    private ?Spreadsheet $spreadsheet_two = null;
    private int $count_something = 0;
    private $fp;

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

    public function findSomeInfoInStringInOneFile(array $array_of_parameters)
    {
        [
            $sheet_name,
            $coordinate,
            $index_for_array,
            $separator_row,
        ] = $array_of_parameters;

        $worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
        $take_highest_row = $worksheet->getHighestRow();

        for ($num_row = 2; $num_row <= $take_highest_row; $num_row++) {
            $take_info_from_cell = $worksheet->getCell($coordinate . $num_row)->getValue();

            if (!empty($take_info_from_cell)) {
                $array_from_cell = explode($separator_row, $take_info_from_cell);
            }

            if (isset($array_from_cell[$index_for_array])) {
                echo $array_from_cell[$index_for_array];
            } else {
                echo 'empty';
            }
            echo '<br/>';
        }
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    private function writeExcelFile(string $directory_path, Spreadsheet $spreadsheet_one = null, ?Spreadsheet $spreadsheet_two = null) : void
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

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function findSomeInfoInOneFileAndWriteInTwoFile(array $array_of_parameters)
    {
        /** Автоформатування змінює на деструктуризацію - так не буде працювати */
        [
            $sheet_name_file_one,
            $sheet_name_file_two,
            $coordinate_one_for_search_in_file_one,
            $coordinate_for_write_in_file_one,
            $coordinate_one_for_search_in_file_two,
            $coordinate_two_for_search_in_file_two,
            $some_info_for_search_and_write,
            $directory_path,
            $separator_for_write,
        ] = $array_of_parameters;

        $worksheet_for_file_one = $this->spreadsheet_one->getSheetByName($sheet_name_file_one);
        $take_highest_row_for_file_one = $worksheet_for_file_one->getHighestRow();
        $worksheet_for_file_two = $this->spreadsheet_two->getSheetByName($sheet_name_file_two);
        $take_highest_row_for_file_two = $worksheet_for_file_two->getHighestRow();

        for ($num_row_one = 2; $num_row_one <= $take_highest_row_for_file_one; $num_row_one++) {
            /** Work with file one */
            $take_one_info_from_cell_from_file_one = trim($worksheet_for_file_one
                    ->getCell($coordinate_one_for_search_in_file_one . $num_row_one)->getValue() ?? '');

            $take_cell_for_write_in_file_one = $worksheet_for_file_one
                ->getCell($coordinate_for_write_in_file_one . $num_row_one);

            $take_info_from_cell_for_write_from_file_one = trim($take_cell_for_write_in_file_one->getValue() ?? '');
            /** Work with file one */

            for ($num_row_two = 2; $num_row_two <= $take_highest_row_for_file_two; $num_row_two++) {
                /** Work with file two */
                $take_one_info_from_cell_from_file_two = trim($worksheet_for_file_two
                        ->getCell($coordinate_one_for_search_in_file_two . $num_row_two)->getValue() ?? '');

                $take_second_cell_for_search_in_file_two = $worksheet_for_file_two
                    ->getCell($coordinate_two_for_search_in_file_two . $num_row_two);

                $take_second_info_from_cell_from_file_two = trim($take_second_cell_for_search_in_file_two->getValue() ?? '');

                $flip_array = array_flip($some_info_for_search_and_write);
                $find_value = $this->findByPartStringInArray($flip_array, $take_second_info_from_cell_from_file_two);

                /** Work with file two */
                if (
                    $take_one_info_from_cell_from_file_one == $take_one_info_from_cell_from_file_two &&
                    !empty($take_second_info_from_cell_from_file_two) &&
                    !empty($find_value)
//                    array_key_exists($take_second_info_from_cell_from_file_two, $some_info_for_search_and_write)
                ) {
                    $take_cell_for_write_in_file_one
                        ->setValue(
                            $take_info_from_cell_for_write_from_file_one .
                            (!empty($separator_for_write) ? $separator_for_write : ' ') .
                            $find_value
//                            $some_info_for_search_and_write[$take_second_info_from_cell_from_file_two]
                        );
                    break;
                }
            }
        }

        $this->writeExcelFile($directory_path, $this->spreadsheet_one, $this->spreadsheet_two);
    }

    /**
     * @param string $sheet_name
     * @param string $coordinate_of_words
     * @param array $array_of_words_for_replace
     * @param array $array_of_words_for_search
     * @param string $directory_path
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function replaceWordsInExcel
    (
        string $sheet_name, string $coordinate_of_words, array $array_of_words_for_replace,
        array  $array_of_words_for_search, string $directory_path
    ) : void
    {
        $worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
        $take_highest_row = $worksheet->getHighestRow();

        for ($index_row = 2; $index_row <= $take_highest_row; $index_row++) {
            $cell_for_search_and_replace = $worksheet->getCell($coordinate_of_words . $index_row);
            $string_from_row = trim($cell_for_search_and_replace->getValue() ?? '');

            if (empty($string_from_row)) continue;

            $string_from_row = $this->replaceWordsInString($array_of_words_for_search, $array_of_words_for_replace, $string_from_row);

            $cell_for_search_and_replace->setValue($string_from_row);
        }

        $this->writeExcelFile($directory_path, $this->spreadsheet_one);
    }

    /**
     * @param array $array_of_words_for_search
     * @param array $array_of_words_for_replace
     * @param mixed $string_from_row
     * @return string
     */
    private function replaceWordsInString(array $array_of_words_for_search, array $array_of_words_for_replace, $string_from_row) : string
    {
        for ($index = 0; $index < count($array_of_words_for_search); $index++) {
            $string_from_row = str_replace(
                $array_of_words_for_search[$index], $array_of_words_for_replace[$array_of_words_for_search[$index]], $string_from_row
            );
        }
        return $string_from_row;
    }

    /**
     * @param array $array
     * @param string $string_where_search
     * @return string
     */
    private function findByPartStringInArray(array $array, string $string_where_search) : string
    {
        while (($value = current($array)) !== null) {
            if (strstr($string_where_search, $value)) return (string)key($array);
            next($array);
        }

        return '';
    }

    public function findAndShowProductsBySKUFromZip(array $array_of_parameters)
    {
        [
            $zip_file,
            $extract_folder,
            $sheet_name_in_one_file,
            $sheet_name_in_two_file,
            $coordinate_in_one_file,
            $coordinate_in_two_file,
        ] = $array_of_parameters;

        $this->deleteFilesInFolder($extract_folder);

        $this->extractFilesFromZip($zip_file, $extract_folder);

        $files = scandir($extract_folder);

        foreach ($files as $file) {
            if ($file == '.' || $file == '..') continue;

            $this->spreadsheet_one = IOFactory::load($extract_folder . $file);

            $worksheet_in_first_file = $this->spreadsheet_one->getSheetByName($sheet_name_in_one_file);
            $worksheet_in_second_file = $this->spreadsheet_two->getSheetByName($sheet_name_in_two_file);

            $take_highest_row_in_first_file = $worksheet_in_first_file->getHighestRow();
            $take_highest_row_in_second_file = $worksheet_in_second_file->getHighestRow();

            echo '<pre style="font-size:16px">';
            for ($index_row = 2; $index_row <= $take_highest_row_in_first_file; $index_row++) {
                $get_cell_from_first_file = trim($worksheet_in_first_file->getCell($coordinate_in_one_file . $index_row)->getValue());

                if (empty($get_cell_from_first_file)) continue;

                for ($index_row_two = 2; $index_row_two <= $take_highest_row_in_second_file; $index_row_two++) {
                    $get_cell_from_second_file = trim($worksheet_in_second_file->getCell($coordinate_in_two_file . $index_row)->getValue());

                    if (empty($get_cell_from_second_file)) continue;

                    if ($get_cell_from_first_file == $get_cell_from_second_file) {
                        print_r($index_row_two . "\n\r");
                    }
                }
            }
            echo '</pre>';
        }
    }

    /**
     * @param $extract_folder
     * @return void
     */
    private function deleteFilesInFolder($extract_folder)
    {
        if (file_exists($extract_folder)) {
            $files = scandir($extract_folder);

            foreach ($files as $file) {
                if ($file != '.' && $file != '..') unlink($extract_folder . $file);
            }
        }
    }

    /**
     * @param $zip_file
     * @param $extract_folder
     */
    private function extractFilesFromZip($zip_file, $extract_folder) : void
    {
        $zip = new ZipArchive();
        $zip->open($zip_file);
        $zip->extractTo($extract_folder);
        $zip->close();
    }

    public function createSQLForUpdateSomething(array $array_of_parameters)
    {
        [
            $zip_file,
            $extract_folder,
            $sheet_name_in_one_file,
            $coordinate_sku_in_one_file,
            $coordinate_sku_in_two_file,
            $coordinate_something_in_two_file,
            $sql_file,
            $mode_write,
            $what_update,
        ] = $array_of_parameters;

        $this->deleteFilesInFolder($extract_folder);

        $this->extractFilesFromZip($zip_file, $extract_folder);

        $files = scandir($extract_folder);

        $all_sheets = $this->spreadsheet_two->getAllSheets();
        $fp = fopen($sql_file, $mode_write);
        $empty_rows = '';

        foreach ($files as $file) {
            if ($file == '.' || $file == '..') continue;

            $this->spreadsheet_one = IOFactory::load($extract_folder . $file);

            $worksheet_in_first_file = $this->spreadsheet_one->getSheetByName($sheet_name_in_one_file);

            $take_highest_row_in_first_file = $worksheet_in_first_file->getHighestRow();

            echo '<pre style="font-size:16px">';
            for ($index_row = 2; $index_row <= $take_highest_row_in_first_file; $index_row++) {
                $product_sku_from_first_file = trim($worksheet_in_first_file->getCell($coordinate_sku_in_one_file . $index_row)->getValue());

                if (empty($product_sku_from_first_file)) continue;

                foreach ($all_sheets as $sheet) {
                    $take_highest_row = $sheet->getHighestRow();

                    for ($index_row2 = 1; $index_row2 <= $take_highest_row; $index_row2++) {
                        $product_sku_from_second_file = trim($sheet->getCell($coordinate_sku_in_two_file . $index_row2)->getValue());
                        $take_product_something = trim($sheet->getCell($coordinate_something_in_two_file . $index_row2)->getValue());

                        if (empty($product_sku_from_second_file) || empty($take_product_something)) {
                            $empty_rows .= "В рядку №$index_row2 відсутня ціна або артикул!!!\r\n";
                            continue;
                        }

                        if ($product_sku_from_first_file == $product_sku_from_second_file) {
                            $sql = "UPDATE `product` SET `" . $what_update . "` = " . $take_product_something . " WHERE `sku` = '" . $product_sku_from_second_file . "';\r\n";

                            if (fwrite($fp, $sql) === false) {
                                fclose($fp);
                                $this->__destruct();
                                exit('An error occurred while writing the file!');
                            }

                            $this->count_something++;
                        }
                    }
                }
            }
            echo '</pre>';

            $this->spreadsheet_two->disconnectWorksheets();
            unset($this->spreadsheet_two);
        }

        fclose($fp);

        $log_info = "Було створено \"$this->count_something\" запитів \r\n\n" . $empty_rows;

        echo '<pre style="font-size: 16px">';
        print_r($log_info);
        echo '</pre>';
    }

    public function createSQLForInsertSomething(array $array_of_parameters)
    {
        [
            $priority,
            $date_end,
            $customer_group_id,
            $coordinate_sku_in_file,
            $coordinate_something_in_file,
            $sql_file,
            $mode_write,
        ] = $array_of_parameters;

        $all_sheets = $this->spreadsheet_one->getAllSheets();
        $fp = fopen($sql_file, $mode_write);
        $empty_rows = '';

        foreach ($all_sheets as $sheet) {
            $take_highest_row = $sheet->getHighestRow();

            for ($index_row = 2; $index_row <= $take_highest_row; $index_row++) {
                $product_sku_from_file = trim($sheet->getCell($coordinate_sku_in_file . $index_row)->getValue());
                $take_product_something = trim($sheet->getCell($coordinate_something_in_file . $index_row)->getValue());

                if (empty($product_sku_from_file) || empty($take_product_something)) {
                    $empty_rows .= "В рядку №$index_row щось відсутнє!!!\r\n";
                    continue;
                }

                $sql = "INSERT INTO `product_special` (`product_id`, `customer_group_id`, `priority`, `price`, `date_end`) VALUES(IFNULL((SELECT `product_id` FROM `product` WHERE `sku` = '$product_sku_from_file' LIMIT 1), 0), $customer_group_id, $priority, $take_product_something, '$date_end');\r\n";

                if (fwrite($fp, $sql) === false) {
                    fclose($fp);
                    $this->__destruct();
                    exit('An error occurred while writing the file!');
                }

                $this->count_something++;
            }
        }

        fclose($fp);

        $log_info = "Було створено \"$this->count_something\" запитів \r\n\n" . $empty_rows;

        echo '<pre style="font-size: 16px">';
        print_r($log_info);
        echo '</pre>';
    }

    public function createSQLForInsertSomething_2(array $array_of_parameters)
    {
        [
            $coordinate_product_id_in_file,
            $sql_file,
            $mode_write,
        ] = $array_of_parameters;

        $fp = fopen($sql_file, $mode_write);
        $empty_rows = '';

        $worksheet = $this->spreadsheet_one->getActiveSheet();

            $take_highest_row = $worksheet->getHighestRow();

            for ($index_row = 2; $index_row <= $take_highest_row; $index_row++) {
                $product_id_from_file = trim($worksheet->getCell($coordinate_product_id_in_file . $index_row)->getValue());

                if (empty($product_id_from_file)) {
                    $empty_rows .= "В рядку №$index_row щось відсутнє!!!\r\n";
                    continue;
                }

                $sql = "INSERT INTO `order_to_1c` (`order_id`, `1c_id`) VALUES($product_id_from_file, '');\r\n";

                if (fwrite($fp, $sql) === false) {
                    fclose($fp);
                    $this->__destruct();
                    exit('An error occurred while writing the file!');
                }

                $this->count_something++;
            }

        fclose($fp);

        $log_info = "Було створено \"$this->count_something\" запитів \r\n\n $empty_rows";

        echo '<pre style="font-size: 16px">';
        print_r($log_info);
        echo '</pre>';
    }

    public function createSQLForInsertSomething_3(array $array_of_parameters)
    {
        [
            $coordinate_product_id_in_file,
            $coordinate_product_price_in_file,
            $date_end,
            $priority,
            $customer_group_id,
            $sql_file,
            $mode_write,
        ] = $array_of_parameters;

        $fp = fopen($sql_file, $mode_write);
        $empty_rows = '';

        $worksheet = $this->spreadsheet_one->getActiveSheet();

        $take_highest_row = $worksheet->getHighestRow();

        for ($index_row = 2; $index_row <= $take_highest_row; $index_row++) {
            $product_id_from_file = trim($worksheet->getCell($coordinate_product_id_in_file . $index_row)->getValue());
            $product_price_from_file = trim($worksheet->getCell($coordinate_product_price_in_file . $index_row)->getValue());

            if (empty($product_id_from_file) || empty($product_price_from_file)) {
                $empty_rows .= "В рядку №$index_row щось відсутнє!!!\r\n";
                continue;
            }

            $special_price = round($product_price_from_file - ($product_price_from_file * 0.17));

            $sql = "INSERT INTO `product_special` (`product_id`, `customer_group_id`, `priority`, `price`, `date_end`) VALUES($product_id_from_file, $customer_group_id, $priority, $special_price, '$date_end');\r\n";

            if (fwrite($fp, $sql) === false) {
                fclose($fp);
                $this->__destruct();
                exit('An error occurred while writing the file!');
            }

            $this->count_something++;
        }

        fclose($fp);

        $log_info = "Було створено \"$this->count_something\" запитів \r\n\n $empty_rows";

        echo '<pre style="font-size: 16px">';
        print_r($log_info);
        echo '</pre>';
    }

    public function createSQLForInsertSomething_5(array $array_of_parameters)
    {
        [
            $coordinate_product_sku_in_file,
            $coordinate_product_price_in_file,
            $date_end,
            $priority,
            $customer_group_id,
            $sql_file,
            $sheet_name,
            $mode_write,
        ] = $array_of_parameters;

        $fp = fopen($sql_file, $mode_write);
        $empty_rows = '';

        if ($sheet_name) {
            $worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
        } else {
            $worksheet = $this->spreadsheet_one->getActiveSheet();
        }

        $take_highest_row = $worksheet->getHighestRow();

        for ($index_row = 2; $index_row <= $take_highest_row; $index_row++) {
            $product_sku_from_file = trim($worksheet->getCell($coordinate_product_sku_in_file . $index_row)->getValue());
            $special_price = trim($worksheet->getCell($coordinate_product_price_in_file . $index_row)->getValue());

            if (empty($product_sku_from_file) || empty($special_price)) {
                $empty_rows .= "В рядку №$index_row щось відсутнє!!!\r\n";
                continue;
            }

            $sql = "INSERT INTO `gg_product_special` (`product_id`, `customer_group_id`, `priority`, `price`, `date_end`) VALUES(IFNULL((SELECT `gg_product_id` FROM `product` WHERE `sku` = '$product_sku_from_file' LIMIT 1), 0), $customer_group_id, $priority, $special_price, '$date_end');\r\n";

            if (fwrite($fp, $sql) === false) {
                fclose($fp);
                $this->__destruct();
                exit('An error occurred while writing the file!');
            }

            $this->count_something++;
        }

        fclose($fp);

        $log_info = "Було створено \"$this->count_something\" запитів \r\n\n $empty_rows";

        echo '<pre style="font-size: 16px">';
        print_r($log_info);
        echo '</pre>';
    }

    public function createSQLForInsertSomething_6(array $array_of_parameters)
    {
        [
            $coordinate_product_id_in_file,
            $sql_file,
            $sheet_name,
            $path_to_image,
            $mode_write,
        ] = $array_of_parameters;

        $fp = fopen($sql_file, $mode_write);
        $empty_rows = '';

        if ($sheet_name) {
            $worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
        } else {
            $worksheet = $this->spreadsheet_one->getActiveSheet();
        }

        $take_highest_row = $worksheet->getHighestRow();

        for ($index_row = 2; $index_row <= $take_highest_row; $index_row++) {
            $product_id_from_file = trim($worksheet->getCell($coordinate_product_id_in_file . $index_row)->getValue());

            if (empty($product_id_from_file)) {
                $empty_rows .= "В рядку №$index_row щось відсутнє!!!\r\n";
                continue;
            }

            $sql = "INSERT INTO `product_image` (`product_id`, `image`, `sort_order`) VALUES($product_id_from_file, '$path_to_image', 1);\r\n";

            if (fwrite($fp, $sql) === false) {
                fclose($fp);
                $this->__destruct();
                exit('An error occurred while writing the file!');
            }

            $this->count_something++;
        }

        fclose($fp);

        $log_info = "Було створено \"$this->count_something\" запитів \r\n\n $empty_rows";

        echo '<pre style="font-size: 16px">';
        print_r($log_info);
        echo '</pre>';
    }

    public function createSQLForInsertSomething_4(array $array_of_parameters)
    {
        [
            $files_directory,
            $sql_file,
            $mode_write,
        ] = $array_of_parameters;

        $this->fp = fopen($sql_file, $mode_write);
        $empty_rows = '';

        $files = scandir($files_directory);

        foreach ($files as $file) {
            if ($file == '.' || $file == '..') continue;

            preg_match('/^\d+/', $file, $matches);
            $sku = isset($matches[0]) && !empty($matches[0]) ? $matches[0] : '';

            if (empty($sku)) {
                $empty_rows .= $file . ' - bad name of file <br/>';
                continue;
            }

            $languages = [
                'ua' => 3,
                'en' => 4,
                'ru' => 5,
            ];

            foreach ($languages as $language) {
                if ($language == 3) {
                    $text = "&lt;p&gt;&lt;a class=&quot;sbtn sbtn-1&quot; href=&quot;../image/instructions/$file&quot;&gt;Переглянути інструкцію&lt;/a&gt;&lt;/p&gt;";
                } else if ($language == 4) {
                    $text = "&lt;p&gt;&lt;a class=&quot;sbtn sbtn-1&quot; href=&quot;../image/instructions/$file&quot;&gt;Look instruction&lt;/a&gt;&lt;/p&gt;";
                } else {
                    $text = "&lt;p&gt;&lt;a class=&quot;sbtn sbtn-1&quot; href=&quot;../image/instructions/$file&quot;&gt;Смотреть инструкцию&lt;/a&gt;&lt;/p&gt;";
                }

                $sql = "
                        INSERT INTO `oct_product_extra_tabs` 
                            (`product_id`, `extra_tab_id`, `language_id`, `text`) 
                            VALUES(IFNULL((SELECT `product_id` FROM `product` 
                                        WHERE sku = '$sku'
                                            LIMIT 1),0), 
                            1, $language, '$text')
                            ON DUPLICATE KEY UPDATE `text` = '$text';\r\n";

                if (fwrite($this->fp, $sql) === false) {
                    $this->__destruct();
                    exit('An error occurred while writing the file!');
                }

                $this->count_something++;
            }
        }

        $log_info = "Було створено \"$this->count_something\" запитів \r\n\n $empty_rows";

        echo '<pre style="font-size: 16px">';
        print_r($log_info);
        echo '</pre>';
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function checkProductImages(array $parameters)
    {
        [$image_coordinate, $id_coordinate] = $parameters;
        $images_folder = 'C:/OpenServer/domains/oc2store.nicktoys.loc/httpdocs/image/catalog/products';

        $images_from_folder = scandir($images_folder);

        $worksheet = $this->spreadsheet_one->getSheet(0);
        $highest_row = $worksheet->getHighestRow();

        $non_existent_images = [];

        for ($index = 2; $index <= $highest_row; $index++) {
            $array_of_cell = explode('/', $worksheet->getCell($image_coordinate . $index)->getValue());

            if (!in_array($array_of_cell[array_key_last($array_of_cell)], $images_from_folder)) {
                array_push($non_existent_images, $worksheet->getCell($id_coordinate . $index)->getValue());
            }
        }

        echo '<pre>';
        if (!empty($non_existent_images)) {
            echo count($non_existent_images) . '<br/>';
            foreach ($non_existent_images as $info) {
                echo $info . ', ';
            }
        } else {
            echo 'Non-existent images not found';
        }
        echo '</pre>';
    }

    /**
     * @throws \PhpOffice\PhpSpreadsheet\Exception
     */
    public function difficultSearchOfPartOfStringInOneFile(array $array_of_parameters)
    {
        [
            $sheet_name,
            $coordinate_where_search,
            $regular_expression,
            $coordinate_product_id,
            $coordinate_for_write_product_id,
            $coordinate_for_write_result,
            $folder,
        ] = $array_of_parameters;

        if (empty($regular_expression)) {
            $this->__destruct();
            exit('Regular expression is empty!!!');
        }

        $worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
        $new_worksheet = $this->spreadsheet_two->getActiveSheet();
        $take_highest_row = $worksheet->getHighestRow();

        $this->count_something = 2;

        for ($num_row = 2; $num_row <= $take_highest_row; $num_row++) {
            $take_info_from_cell = $worksheet->getCell($coordinate_where_search . $num_row)->getValue();
            $take_product_id = $worksheet->getCell($coordinate_product_id . $num_row)->getValue();

            if (preg_match($regular_expression, $take_info_from_cell, $matches) == 1) {
                $new_worksheet->getCell($coordinate_for_write_product_id . $this->count_something)->setValue($take_product_id);
                $new_worksheet->getCell($coordinate_for_write_result . $this->count_something)->setValue($matches[0]);
                $this->count_something++;
            }
        }

        $this->writeExcelFile($folder, null, $this->spreadsheet_two);
    }

    public function createSQLForUpdateSomething_2(array $array_of_parameters)
    {
        [
            $zip_file,
            $extract_folder,
            $coordinate_of_product_id,
            $coordinate_of_size,
            $sql_file,
            $mode_write,
            $what_update,
            $attribute_id,
        ] = $array_of_parameters;

        $this->deleteFilesInFolder($extract_folder);

        $this->extractFilesFromZip($zip_file, $extract_folder);

        $files = scandir($extract_folder);

        $fp = fopen($sql_file, $mode_write);
        $empty_rows = '';

        foreach ($files as $file) {
            if ($file == '.' || $file == '..') continue;

            $this->spreadsheet_one = IOFactory::load($extract_folder . $file);

            $worksheet = $this->spreadsheet_one->getActiveSheet('Products');
            $take_highest_row = $worksheet->getHighestRow();

            echo '<pre style="font-size:16px">';

            for ($index_row = 2; $index_row <= $take_highest_row; $index_row++) {
                $product_id = (int)trim($worksheet->getCell($coordinate_of_product_id . $index_row)->getValue());
                $product_size = trim($worksheet->getCell($coordinate_of_size . $index_row)->getValue());

                if (empty($product_id) || empty($product_size)) {
                    $empty_rows .= "В рядку №$index_row відсутня ціна або артикул!!!\r\n";
                    continue;
                }

                $sql = "UPDATE `product_attribute` SET `" . $what_update . "` = '" . $product_size . "' WHERE `product_id` = " . $product_id . " AND `attribute_id` = " . $attribute_id . " AND `" . $what_update . "` = '';\r\n";

                if (fwrite($fp, $sql) === false) {
                    fclose($fp);
                    $this->__destruct();
                    exit('An error occurred while writing the file!');
                }

                $this->count_something++;
            }
            echo '</pre>';

            if ($file == $files[array_key_last($files)]) continue;

            $this->spreadsheet_one->disconnectWorksheets();
            unset($this->spreadsheet_one);
        }

        fclose($fp);

        $log_info = "Було створено \"$this->count_something\" запитів \r\n\n" . $empty_rows;

        echo '<pre style="font-size: 16px">';
        print_r($log_info);
        echo '</pre>';
    }

    public function createSQLForUpdateSomething_3(array $array_of_parameters)
    {
        [
            $sql_file,
            $mode_write,
            $what_update,
            $attribute_id,
        ] = $array_of_parameters;

        $fp = fopen($sql_file, $mode_write);
        $empty_rows = '';
        $take_highest_id = 2650;

        echo '<pre style="font-size:16px">';

        for ($product_id = 2; $product_id <= $take_highest_id; $product_id++) {

            $sql = "UPDATE `product_attribute` SET `" . $what_update . "` = '1 кг' WHERE `product_id` = " . $product_id . " AND `attribute_id` = " . $attribute_id . " AND `" . $what_update . "` = '' AND (`language_id` = 3 OR `language_id` = 5);\r\n";

            if (fwrite($fp, $sql) === false) {
                fclose($fp);
                exit('An error occurred while writing the file!');
            }

            $this->count_something++;
        }
        echo '</pre>';

        fclose($fp);

        $log_info = "Було створено \"$this->count_something\" запитів \r\n\n" . $empty_rows;

        echo '<pre style="font-size: 16px">';
        print_r($log_info);
        echo '</pre>';
    }

    public function createSQLForUpdateSomething_5(array $array_of_parameters)
    {
        [
            $coordinate_product_sku_in_file,
            $coordinate_product_price_in_file,
            $sql_file,
            $sheet_name,
            $mode_write,
        ] = $array_of_parameters;

        $fp = fopen($sql_file, $mode_write);
        $empty_rows = '';

        if ($sheet_name) {
            $worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
        } else {
            $worksheet = $this->spreadsheet_one->getActiveSheet();
        }

        $take_highest_row = $worksheet->getHighestRow();

        for ($index_row = 2; $index_row <= $take_highest_row; $index_row++) {
            $product_sku_from_file = trim($worksheet->getCell($coordinate_product_sku_in_file . $index_row)->getValue());
            $price = trim($worksheet->getCell($coordinate_product_price_in_file . $index_row)->getValue());

            if (empty($price) || !is_numeric($price) || empty($product_sku_from_file)) {
                $empty_rows .= "В рядку №$index_row щось відсутнє!!!\r\n";
                continue;
            }

            $sql = "UPDATE `gg_product` SET `price` = $price WHERE `sku` = '$product_sku_from_file';\r\n";

            if (fwrite($fp, $sql) === false) {
                fclose($fp);
                exit('An error occurred while writing the file!');
            }

            $this->count_something++;
        }
        echo '</pre>';

        fclose($fp);

        $log_info = "Було створено \"$this->count_something\" запитів \r\n\n" . $empty_rows;

        echo '<pre style="font-size: 16px">';
        print_r($log_info);
        echo '</pre>';
    }

    public function createSQLForUpdateSomething_4(array $array_of_parameters)
    {
        [
            $zip_file,
            $extract_folder,
            $coordinate_sku,
            $coordinate_something,
            $sql_file,
            $mode_write,
            $what_update,
            $attribute_id,
        ] = $array_of_parameters;

        $this->deleteFilesInFolder($extract_folder);

        $this->extractFilesFromZip($zip_file, $extract_folder);

        $files = scandir($extract_folder);

        $fp = fopen($sql_file, $mode_write);
        $empty_rows = '';

        foreach ($files as $file) {
            if ($file == '.' || $file == '..') continue;

            $this->spreadsheet_one = IOFactory::load($extract_folder . $file);
            $all_sheets = $this->spreadsheet_one->getAllSheets();

            echo '<pre style="font-size:16px">';
            foreach ($all_sheets as $sheet) {
                $take_highest_row = $sheet->getHighestRow();

                for ($index_row = 2; $index_row <= $take_highest_row; $index_row++) {
                    $product_sku = trim($sheet->getCell($coordinate_sku . $index_row)->getValue());
                    $take_product_something = trim($sheet->getCell($coordinate_something . $index_row)->getValue());

                    if (empty($product_sku) || empty($take_product_something)) {
                        $empty_rows .= "В рядку №$index_row щось відсутнє!!!\r\n";
                        continue;
                    }

                    $sql = "UPDATE `product_attribute` SET `" . $what_update . "` = '" . $take_product_something . "' WHERE `product_id` = (IFNULL((SELECT `product_id` FROM `product` WHERE `sku` = '" . $product_sku . "' LIMIT 1), 0)) AND `attribute_id` = " . $attribute_id . " AND `text` = '';\r\n";

                    if (fwrite($fp, $sql) === false) {
                        fclose($fp);
                        $this->__destruct();
                        exit('An error occurred while writing the file!');
                    }

                    $this->count_something++;
                }
            }
            echo '</pre>';

            if ($file == $files[array_key_last($files)]) continue;

            $this->spreadsheet_one->disconnectWorksheets();
            unset($this->spreadsheet_one);
        }

        fclose($fp);

        $log_info = "Було створено \"$this->count_something\" запитів \r\n\n" . $empty_rows;

        echo '<pre style="font-size: 16px">';
        print_r($log_info);
        echo '</pre>';
    }

    /**
     * @param array $parameters
     * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
     */
    public function createAttributes(array $parameters) : void
    {
        [
            $folder_excels,
            $sheet_name,
        ] = $parameters;

        $worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
        $highest_row = $worksheet->getHighestRow();

        $new_spreadsheet = new Spreadsheet();
        $new_worksheet = $new_spreadsheet->getActiveSheet();

        $col_product_id = trim($worksheet->getCell('A1')->getValue());
        $col_manufacturer = trim($worksheet->getCell('B1')->getValue());
        $col_canvas_mat = trim($worksheet->getCell('C1')->getValue());
        $col_weight = trim($worksheet->getCell('D1')->getValue());
        $col_size = trim($worksheet->getCell('E1')->getValue());
        $col_package = trim($worksheet->getCell('F1')->getValue());
        $col_num_of_colors = trim($worksheet->getCell('G1')->getValue());
        $col_difficulty_level = trim($worksheet->getCell('H1')->getValue());

        $col_dop_1 = trim($worksheet->getCell('I1')->getValue());
        $col_dop_2 = trim($worksheet->getCell('J1')->getValue());
        $col_dop_3 = trim($worksheet->getCell('K1')->getValue());

        $this->count_something = 2;

        for ($row = 2; $row <= $highest_row; $row++) {
            $manufacturer_val = trim($worksheet->getCell('B' . $row)->getValue());

            if (!empty($manufacturer_val)) {
                $new_worksheet->setCellValue('A' . $this->count_something, $worksheet->getCell('A' . $row)->getValue());
                $new_worksheet->setCellValue('B' . $this->count_something, $col_manufacturer);
                $new_worksheet->setCellValue('C' . $this->count_something, $manufacturer_val);

                $this->count_something++;
            }

            $canvas_mat_val = trim($worksheet->getCell('C' . $row)->getValue());

            if (!empty($canvas_mat_val)) {
                $new_worksheet->setCellValue('A' . $this->count_something, $worksheet->getCell('A' . $row)->getValue());
                $new_worksheet->setCellValue('B' . $this->count_something, $col_canvas_mat);
                $new_worksheet->setCellValue('C' . $this->count_something, $canvas_mat_val);

                $this->count_something++;
            }

            $weight_val = trim($worksheet->getCell('D' . $row)->getValue());

            if (!empty($weight_val)) {
                $new_worksheet->setCellValue('A' . $this->count_something, $worksheet->getCell('A' . $row)->getValue());
                $new_worksheet->setCellValue('B' . $this->count_something, $col_weight);
                $new_worksheet->setCellValue('C' . $this->count_something, $weight_val);

                $this->count_something++;
            }

            $size_val = trim($worksheet->getCell('E' . $row)->getValue());

            if (!empty($size_val)) {
                $new_worksheet->setCellValue('A' . $this->count_something, $worksheet->getCell('A' . $row)->getValue());
                $new_worksheet->setCellValue('B' . $this->count_something, $col_size);
                $new_worksheet->setCellValue('C' . $this->count_something, $size_val);

                $this->count_something++;
            }

            $package_val = trim($worksheet->getCell('F' . $row)->getValue());

            if (!empty($package_val)) {
                $new_worksheet->setCellValue('A' . $this->count_something, $worksheet->getCell('A' . $row)->getValue());
                $new_worksheet->setCellValue('B' . $this->count_something, $col_package);
                $new_worksheet->setCellValue('C' . $this->count_something, $package_val);

                $this->count_something++;
            }

            $num_of_colors_val = trim($worksheet->getCell('G' . $row)->getValue());

            if (!empty($num_of_colors_val)) {
                $new_worksheet->setCellValue('A' . $this->count_something, $worksheet->getCell('A' . $row)->getValue());
                $new_worksheet->setCellValue('B' . $this->count_something, $col_num_of_colors);
                $new_worksheet->setCellValue('C' . $this->count_something, $num_of_colors_val);

                $this->count_something++;
            }

            $difficulty_level_val = trim($worksheet->getCell('H' . $row)->getValue());

            if (!empty($difficulty_level_val)) {
                $new_worksheet->setCellValue('A' . $this->count_something, $worksheet->getCell('A' . $row)->getValue());
                $new_worksheet->setCellValue('B' . $this->count_something, $col_difficulty_level);
                $new_worksheet->setCellValue('C' . $this->count_something, $difficulty_level_val);

                $this->count_something++;
            }

            /** For nicktoys.com.ua */

            $col_dop_1_val = trim($worksheet->getCell('I' . $row)->getValue());

            if (!empty($col_dop_1_val)) {
                $new_worksheet->setCellValue('A' . $this->count_something, $worksheet->getCell('A' . $row)->getValue());
                $new_worksheet->setCellValue('B' . $this->count_something, $col_dop_1);
                $new_worksheet->setCellValue('C' . $this->count_something, $col_dop_1_val);

                $this->count_something++;
            }

            $col_dop_2_val = trim($worksheet->getCell('J' . $row)->getValue());

            if (!empty($col_dop_2_val)) {
                $new_worksheet->setCellValue('A' . $this->count_something, $worksheet->getCell('A' . $row)->getValue());
                $new_worksheet->setCellValue('B' . $this->count_something, $col_dop_2);
                $new_worksheet->setCellValue('C' . $this->count_something, $col_dop_2_val);

                $this->count_something++;
            }

            $col_dop_3_val = trim($worksheet->getCell('K' . $row)->getValue());

            if (!empty($col_dop_3_val)) {
                $new_worksheet->setCellValue('A' . $this->count_something, $worksheet->getCell('A' . $row)->getValue());
                $new_worksheet->setCellValue('B' . $this->count_something, $col_dop_3);
                $new_worksheet->setCellValue('C' . $this->count_something, $col_dop_3_val);

                $this->count_something++;
            }
        }

        if (!file_exists($folder_excels)) mkdir($folder_excels);

        $this->writeExcelFile($folder_excels, $new_spreadsheet);

        $new_spreadsheet->disconnectWorksheets();
        unset($new_spreadsheet);
    }

    public function __destruct()
    {
        if (isset($this->fp)) fclose($this->fp);

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

/*$files = 'F:/strateg/Лайно/інструкції';

$parameters = [
    'F:/strateg/Лайно/інструкції',
    'F:/strateg/Лайно/інструкції/create_insert_pdf_links_in_product.sql',
    'w',
];

$create_insert_pdf_links_in_product_sql = new ExcelWork();
$create_insert_pdf_links_in_product_sql->createSQLForInsertSomething_4($parameters);*/

/*$file = 'F:/strateg/Добавление всякой нездравой фигни/акції avstore.com.ua 26.10.2022/products_from_avstore.com.ua_26.10.2022.xlsx';

$parameters = [
    'A',
    'B',
    '2022-11-03',
    1,
    1,
    'F:/strateg/Добавление всякой нездравой фигни/акції avstore.com.ua 26.10.2022/insert_product_specials_avstore.com.ua.sql',
    'w',
];

$create_insert_product_ids_sql = new ExcelWork($file);
$create_insert_product_ids_sql->createSQLForInsertSomething_3($parameters);*/

$file = 'F:/strateg/Лайно/2. Картини.xlsx';

$parameters = [
    'A',
    'D',
    '2022-11-30',
    1,
    1,
    'F:/strateg/Лайно/specials_price_for_pictures_nicktoys.com.ua.sql',
    'Ніктойс',
    'w',
];

$create_insert_product_ids_sql = new ExcelWork($file);
$create_insert_product_ids_sql->createSQLForInsertSomething_5($parameters);

/*$file = 'F:/strateg/Оновлення не здової фігні/08.11.2022 додаткові зображення avstore.com.ua/product_ids_special-diamond-mosaic_from_avstore.com.ua_08.11.2022.xlsx';

$parameters = [
    'A',
    'F:/strateg/Оновлення не здової фігні/08.11.2022 додаткові зображення avstore.com.ua/insert_craft-box-for-special-diamond-mosaic_avstore.com.ua.sql',
    null,
    'catalog/product/craft-box-diamond-mosaic_40x50_v.1.jpg',
    'w',
];

$create_insert_product_ids_sql = new ExcelWork($file);
$create_insert_product_ids_sql->createSQLForInsertSomething_6($parameters);*/

/*$file = 'F:/strateg/Оновлення не здової фігні/01.11.2022 акції та РРЦ ціна nicktoys.com.ua/(4. Технок) nicktoys.com.ua_01.11.2022.xlsx';

$parameters = [
    'A',
    'B',
    'F:/strateg/Оновлення не здової фігні/01.11.2022 акції та РРЦ ціна nicktoys.com.ua/update_product_price_nicktoys.com.ua_2.sql',
    'Нова основна ціна ніктойс',
    'w',
];

$create_insert_product_ids_sql = new ExcelWork($file);
$create_insert_product_ids_sql->createSQLForUpdateSomething_5($parameters);*/

/*$file = 'F:/strateg/Лайно/order_ids.xlsx';
$parameters = [
    'A',
    'F:/strateg/Лайно/insert_order_ids.sql',
    'w',
];

$create_insert_product_ids_sql = new ExcelWork($file);
$create_insert_product_ids_sql->createSQLForInsertSomething_2($parameters);*/

/*$file = 'F:/strateg/Добавление новых позиций/26.09.2022_avstore.com.ua/new_products_excelport_bulk_en_gb_avstore_com_ua_2022_09_22_11_25_34.xlsx';

$params = [
    1,
    '2022-09-18',
    1,
    'D',
    'K',
    'F:/strateg/Оновлення не здової фігні/12.09.2022_акції_avstore.com.ua/sqls_insert/inserts_for_diamond_mosaic.sql',
    'w',
];

$create_sql_specials = new ExcelWork($file);
$create_sql_specials->createSQLForInsertSomething($params);*/

/*$file = 'F:/strateg/Добавление новых позиций/08.11.2022_nicktoys.com.ua/uk-ua__nicktoys.com.ua_ 03.11.2022.xlsx';
$parameters = [
    'F:/strateg/Добавление новых позиций/08.11.2022_nicktoys.com.ua/excels/',
    'Attribute',
];

$create_attributes = new ExcelWork($file);
try {
    $create_attributes->createAttributes($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
    echo $e->getMessage();
}*/


/*$file_one = 'F:/strateg/avstore/backup_товаров_avstore/30.08.2022/products_excelport_bulk_uk-ua__avstore.com.ua_2022-08-30_12-09-39_1600.xlsx';
$file_two = 'F:/strateg/avstore/backup_товаров_avstore/30.08.2022/3.xlsx';

$parameters = [
    'Products',
    'B',
    '/(\d{2}\s?(х|x)\s?\d{2})/ui',
    'A',
    'A',
    'B',
    'F:/strateg/avstore/backup_товаров_avstore/30.08.2022/'
];

$difficult_search = new ExcelWork($file_one, $file_two);
try {
    $difficult_search->difficultSearchOfPartOfStringInOneFile($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
    echo $e->getMessage();
}*/

/*$file_one = 'F:/strateg/Оновлення не здової фігні/26.09.2022_avstore.com.ua_specials/Знижка_на_алмазку_35_відсотків_від_26_09_2022.xlsx';

$parameters = [
    1,
    '2022-10-10',
    1,
    'D',
    'K',
    'F:/strateg/Оновлення не здової фігні/26.09.2022_avstore.com.ua_specials/sql_insert_new_special.sql',
    'w',
];

$create_sql_insert = new ExcelWork($file_one);
$create_sql_insert->createSQLForInsertSomething($parameters);*/

/*$parameters = [
    'F:/strateg/avstore/backup_товаров_avstore/30.08.2022/1.zip',
    'F:/strateg/avstore/backup_товаров_avstore/30.08.2022/extract_excels/',
    'A',
    'B',
    'F:/strateg/avstore/backup_товаров_avstore/30.08.2022/update.sql',
    'w+',
    'text',
    14,
];

$create_sql_update = new ExcelWork();
$create_sql_update->createSQLForUpdateSomething_2($parameters);*/

/*$parameters = [
    'F:/strateg/avstore/backup_товаров_avstore/30.08.2022/update.sql',
    'w+',
    'text',
    14,
];

$create_sql_update = new ExcelWork();
$create_sql_update->createSQLForUpdateSomething_3($parameters);*/

/*$parameters = [
    'F:/strateg/avstore/backup_товаров_avstore/30.08.2022/1.zip',
    'F:/strateg/avstore/backup_товаров_avstore/30.08.2022/extract_excels/',
    'D',
    'E',
    'F:/strateg/avstore/backup_товаров_avstore/30.08.2022/update.sql',
    'w+',
    'text',
    18,
];

$create_sql_update = new ExcelWork();
$create_sql_update->createSQLForUpdateSomething_4($parameters);*/

/*$file = 'F:/strateg/Добавление новых позиций/02.09.2022_nicktoys.com.ua/products_excelport_bulk_uk_ua_nicktoys_com_ua_2022_08_31_10_52_02.xlsx';

$parameters = [
    'Products',
    'B',
    'X',
    'en',
    'F:/strateg/Добавление новых позиций/02.09.2022_nicktoys.com.ua/Новая папка/',
    '-ua',
];

$create_seo_urls = new ExcelWork($file);
try {
    $create_seo_urls->createSeoUrl($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
    echo $e->getMessage();
}*/
