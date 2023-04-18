<?php /** @noinspection SqlResolve */

/** @noinspection SqlDialectInspection */
/** @noinspection SqlNoDataSourceInspection */

require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Csv;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

mb_internal_encoding('UTF-8');

class PhpExcel
{
	private ?Spreadsheet $spreadsheet_one = null;
	private ?Spreadsheet $spreadsheet_two = null;
	private ?Csv $csv_one = null;
	private ?Csv $csv_two = null;
	private string $file_name_one;
	private string $file_name_two;
	private int $count_something = 0;

	/**
	 * @throws \PhpOffice\PhpSpreadsheet\Reader\Exception
	 */
	public function __construct(string $spreadsheet_file_name_one, bool $is_need_create_file = false, string $spreadsheet_file_name_two = '', array $parameters = [])
	{
		$arr_of_name_one = explode('.', $spreadsheet_file_name_one);
		$arr_of_name_two = ($spreadsheet_file_name_two ? explode('.', $spreadsheet_file_name_two) : []);

		$type_of_file_one = end($arr_of_name_one);
		$type_of_file_two = (end($arr_of_name_two) ?: 'xlsx');

		if (!$is_need_create_file) {
			if (!$spreadsheet_file_name_two && ($type_of_file_one == 'xls' || $type_of_file_one == 'xlsx')) {
				$this->spreadsheet_one = IOFactory::load($spreadsheet_file_name_one);
			} else if (!$spreadsheet_file_name_two && $type_of_file_one == 'csv') {
				$this->csv_one = new Csv();

				if (isset($parameters['input_encoding'])) $this->csv_one->setInputEncoding($parameters['input_encoding']);
				if (isset($parameters['delimiter'])) $this->csv_one->setDelimiter($parameters['delimiter']);
				if (isset($parameters['enclosure'])) $this->csv_one->setEnclosure($parameters['enclosure']);
				if (isset($parameters['sheet_index'])) $this->csv_one->setSheetIndex($parameters['sheet_index']);

				$this->spreadsheet_one = $this->csv_one->load($spreadsheet_file_name_one);
			} else if (
				($type_of_file_one == 'xls' || $type_of_file_one == 'xlsx') &&
				($type_of_file_two == 'xlsx' || $type_of_file_two == 'xls')
			) {
				$this->spreadsheet_one = IOFactory::load($spreadsheet_file_name_one);
				$this->spreadsheet_two = IOFactory::load($spreadsheet_file_name_two);
			} else if (
				($type_of_file_one == 'xls' || $type_of_file_one == 'xlsx') &&
				$type_of_file_two == 'csv'
			) {
				$this->spreadsheet_one = IOFactory::load($spreadsheet_file_name_one);

				$this->csv_one = new Csv();

				if (isset($parameters['input_encoding'])) $this->csv_one->setInputEncoding($parameters['input_encoding']);
				if (isset($parameters['delimiter'])) $this->csv_one->setDelimiter($parameters['delimiter']);
				if (isset($parameters['enclosure'])) $this->csv_one->setEnclosure($parameters['enclosure']);
				if (isset($parameters['sheet_index'])) $this->csv_one->setSheetIndex($parameters['sheet_index']);

				$this->spreadsheet_two = $this->csv_one->load($spreadsheet_file_name_two);
			} else if ($type_of_file_one == 'csv' && $type_of_file_two == 'csv') {
				$this->csv_one = new Csv();

				if (isset($parameters['input_encoding_one'])) $this->csv_one->setInputEncoding($parameters['input_encoding_one']);
				if (isset($parameters['delimiter_one'])) $this->csv_one->setDelimiter($parameters['delimiter_one']);
				if (isset($parameters['enclosure_one'])) $this->csv_one->setEnclosure($parameters['enclosure_one']);
				if (isset($parameters['sheet_index_one'])) $this->csv_one->setSheetIndex($parameters['sheet_index_one']);

				$this->spreadsheet_one = $this->csv_one->load($spreadsheet_file_name_one);

				$this->csv_two = new Csv();

				if (isset($parameters['input_encoding_two'])) $this->csv_two->setInputEncoding($parameters['input_encoding_two']);
				if (isset($parameters['delimiter_two'])) $this->csv_two->setDelimiter($parameters['delimiter_two']);
				if (isset($parameters['enclosure_two'])) $this->csv_two->setEnclosure($parameters['enclosure_two']);
				if (isset($parameters['sheet_index_two'])) $this->csv_two->setSheetIndex($parameters['sheet_index_two']);

				$this->spreadsheet_two = $this->csv_two->load($spreadsheet_file_name_two);
			} else if (
				$type_of_file_one == 'csv' &&
				($type_of_file_two == 'xlsx' || $type_of_file_two == 'xls')
			) {
				$this->csv_one = new Csv();

				if (isset($parameters['input_encoding_one'])) $this->csv_one->setInputEncoding($parameters['input_encoding_one']);
				if (isset($parameters['delimiter_one'])) $this->csv_one->setDelimiter($parameters['delimiter_one']);
				if (isset($parameters['enclosure_one'])) $this->csv_one->setEnclosure($parameters['enclosure_one']);
				if (isset($parameters['sheet_index_one'])) $this->csv_one->setSheetIndex($parameters['sheet_index_one']);

				$this->spreadsheet_one = $this->csv_one->load($spreadsheet_file_name_one);

				$this->spreadsheet_two = IOFactory::load($spreadsheet_file_name_two);
			}
		} else {
			$this->spreadsheet_one = new Spreadsheet();
		}

		$this->file_name_one = basename($spreadsheet_file_name_one);
		$this->file_name_two = basename($spreadsheet_file_name_two);
	}

	/**
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	private function writeExcelFile(string $path_to_files, $file_one, ?Spreadsheet $new_spreadsheet = null) : void
	{
		if ($new_spreadsheet instanceof Spreadsheet) {
			$writer = new Xlsx($new_spreadsheet);
		} else {
			$writer = new Xlsx($this->spreadsheet_one);
		}

		if ($path_to_files[mb_strlen($path_to_files) - 1] == '/') {
			$writer->save($path_to_files . 'new_' . $file_one);
		} else {
			$writer->save($path_to_files . '/new_' . $file_one);
		}

		echo "Success for file new_$file_one";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createUpdateSqlForProductPrice(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_sku = mb_strtoupper($coordinate_prod_sku);
		$coordinate_prod_price = mb_strtoupper($coordinate_prod_price);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		if ($is_platoshka_db) {
			$table_property = 'model';
		} else {
			$table_property = 'sku';
		}

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_sku = trim($worksheet->getCell($coordinate_prod_sku . $row_number)->getValue());
			$prod_price = (float)trim($worksheet->getCell($coordinate_prod_price . $row_number)->getValue());

			if (empty($prod_sku) || !is_numeric($prod_price) || empty($prod_price)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! - prod_sku: $prod_sku; prod_price: $prod_price \n\r";
				continue;
			}

			$sql = "UPDATE `{$db_prefix}product` SET `price` = $prod_price WHERE `product_id` = IFNULL((SELECT `product_id` FROM `{$db_prefix}product` WHERE `$table_property` = '$prod_sku' LIMIT 1), 0);\n";

			fwrite($fp, $sql);
			$this->count_something++;
		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createInsertInProductCategory(array $parameters) : void
	{
		extract($parameters);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		$coordinate_prod_sku = mb_strtoupper($coordinate_prod_sku);
		$coordinate_cat_name_level_1 = mb_strtoupper($coordinate_cat_name_level_1);
		$coordinate_cat_name_level_2 = mb_strtoupper($coordinate_cat_name_level_2);

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_sku = trim($worksheet->getCell($coordinate_prod_sku . $row_number)->getValue());
			$cat_name_level_1 = trim($worksheet->getCell($coordinate_cat_name_level_1 . $row_number)->getValue());
			$cat_name_level_2 = trim($worksheet->getCell($coordinate_cat_name_level_2 . $row_number)->getValue());

			if (empty($prod_sku) || (empty($cat_name_level_1) && empty($cat_name_level_2))) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! \n\r";
				continue;
			}

			$cat_name_level_1 = str_replace("'", "\'", $cat_name_level_1);

			if (!empty($cat_name_level_2)) {
				$cat_name_level_2 = str_replace("'", "\'", $cat_name_level_2);
				$sql = "INSERT INTO `{$db_prefix}product_to_category` (`product_id`, `category_id`) VALUES (IFNULL((SELECT product_id FROM `{$db_prefix}product` WHERE sku = '$prod_sku' LIMIT 1), 0), IFNULL((SELECT cd.category_id FROM `{$db_prefix}category_description` AS cd RIGHT JOIN `{$db_prefix}category` AS c ON (cd.category_id = c.category_id) WHERE cd.`name` = '$cat_name_level_2' AND c.parent_id = (SELECT cd.category_id FROM `{$db_prefix}category` AS c RIGHT JOIN `{$db_prefix}category_description` AS cd ON (c.category_id = cd.category_id) WHERE cd.`name` = '$cat_name_level_1' AND c.parent_id = $main_parent_cat_id LIMIT 1) LIMIT 1), 0)) ON DUPLICATE KEY UPDATE `category_id` = `category_id`;\n";

				fwrite($fp, $sql);
				$this->count_something++;
			}


			if (!empty($cat_name_level_1)) {
				$sql = "INSERT INTO `{$db_prefix}product_to_category` (`product_id`, `category_id`) VALUES (IFNULL((SELECT product_id FROM `{$db_prefix}product` WHERE sku = '$prod_sku' LIMIT 1), 0), IFNULL((SELECT c.category_id FROM `{$db_prefix}category_description` AS cd RIGHT JOIN `{$db_prefix}category` AS c ON (c.category_id = cd.category_id) WHERE c.parent_id = $main_parent_cat_id AND cd.`name` = '$cat_name_level_1' LIMIT 1), 0)) ON DUPLICATE KEY UPDATE `category_id` = `category_id`;\n";

				fwrite($fp, $sql);
				$this->count_something++;
			}

		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createInsertInProductCategory_2(array $parameters) : void
	{
		extract($parameters);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		$coordinate_prod_sku = mb_strtoupper($coordinate_prod_sku);

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_sku = trim($worksheet->getCell($coordinate_prod_sku . $row_number)->getValue());

			if (empty($prod_sku)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! \n\r";
				continue;
			}

			$sql = "DELETE FROM `{$db_prefix}product_to_category` WHERE product_id = IFNULL((SELECT product_id FROM `{$db_prefix}product` WHERE sku = '$prod_sku' LIMIT 1), 0);\n";

			$sql .= "INSERT INTO `{$db_prefix}product_to_category` (`product_id`, `category_id`, `main_category`) VALUES (IFNULL((SELECT product_id FROM `{$db_prefix}product` WHERE sku = '$prod_sku' LIMIT 1), 0), $sub_cat_id, $main_cat_id);\n";

			fwrite($fp, $sql);
			$this->count_something++;
		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createInsertSqlInProductSpecials(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_sku = mb_strtoupper($coordinate_prod_sku);
		$coordinate_special_price = mb_strtoupper($coordinate_special_price);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		if ($is_platoshka_db) {
			$table_property = 'model';
		} else {
			$table_property = 'sku';
		}

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_sku = trim($worksheet->getCell($coordinate_prod_sku . $row_number)->getValue());
			$special_price = (float)trim($worksheet->getCell($coordinate_special_price . $row_number)->getValue());

			if (empty($prod_sku) || empty($special_price) && !is_numeric($special_price)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! \n\r";
				continue;
			}

			/*$sql = "DELETE FROM `{$db_prefix}product_special` WHERE product_id = IFNULL((SELECT `product_id` FROM `{$db_prefix}product` WHERE `$table_property` = '$prod_sku' LIMIT 1), 0) AND date_start = '$date_start' AND date_end = '$date_end';\n";

			fwrite($fp, $sql);
			$this->count_something++;*/

			$sql = "INSERT INTO `{$db_prefix}product_special` (`product_id`, `customer_group_id`, `priority`, `price`, `date_start`, `date_end`) VALUES(IFNULL((SELECT `product_id` FROM `{$db_prefix}product` WHERE `$table_property` = '$prod_sku' LIMIT 1), 0), $customer_group_id, $priority, $special_price, '$date_start', '$date_end');\n";

			fwrite($fp, $sql);
			$this->count_something++;
		}

		$sql = "DELETE FROM `{$db_prefix}product_special` WHERE `product_id` = 0;";

		fwrite($fp, $sql);
		$this->count_something++;

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createDeleteFromProductSpecials(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_sku = mb_strtoupper($coordinate_prod_sku);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_sku = trim($worksheet->getCell($coordinate_prod_sku . $row_number)->getValue());

			if (empty($prod_sku)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! \n\r";
				continue;
			}

			$sql = "DELETE FROM gg_product_special WHERE product_id = IFNULL((SELECT product_id FROM gg_product WHERE sku = '$prod_sku' LIMIT 1), 0) AND customer_group_id = $customer_group_id;\n";

			fwrite($fp, $sql);
			$this->count_something++;
		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function createAttributes(array $parameters) : void
	{
		extract($parameters);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		$highest_row = $worksheet->getHighestRow();

		$new_spreadsheet = new Spreadsheet();
		$new_worksheet = $new_spreadsheet->getActiveSheet();

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

		$main_col_name_product_id = 'Product ID';
		$main_col_name_attribute = 'Attribute';
		$main_col_name_text = 'Text';

		$this->count_something = 2;

		$new_worksheet->setCellValue('A1', $main_col_name_product_id);
		$new_worksheet->setCellValue('B1', $main_col_name_attribute);
		$new_worksheet->setCellValue('C1', $main_col_name_text);

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

		$this->writeExcelFile($folder_excels, $this->file_name_one, $new_spreadsheet);

		$new_spreadsheet->disconnectWorksheets();
		unset($new_spreadsheet);
	}

	/**
	 * @param array $parameters
	 * @return void
	 * @throws Exception
	 */
	public function createSqlForProductRating(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_id = mb_strtoupper($coordinate_prod_id);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_id = trim($worksheet->getCell($coordinate_prod_id . $row_number)->getValue());

			if (empty($prod_id) || !is_numeric($prod_id)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! \n\r";
				continue;
			}

			$random_rating = random_int(4, 5);
			$random_view = random_int(3, 1000);

			$sql = "INSERT INTO `review` (`product_id`, `rating`, `status`) VALUES($prod_id, $random_rating, 1);\n";

			fwrite($fp, $sql);
			$this->count_something++;

			$sql = "UPDATE product SET viewed = (viewed + $random_view) WHERE product_id = $prod_id;\n";

			fwrite($fp, $sql);
			$this->count_something++;
		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createUpdateSqlInProductSpecials(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_sku = mb_strtoupper($coordinate_prod_sku);
		$coordinate_special_price = mb_strtoupper($coordinate_special_price);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		if (empty($date_end)) $date_end = 'NOW()';
		else $date_end = "'$date_end'";

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_sku = trim($worksheet->getCell($coordinate_prod_sku . $row_number)->getValue());
			$special_price = trim($worksheet->getCell($coordinate_special_price . $row_number)->getValue());

			if (empty($prod_sku) || empty($special_price) && !is_numeric($special_price)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! \n\r";
				continue;
			}

			$sql = "UPDATE `product_special` SET `price` = $special_price WHERE `product_id` = IFNULL((SELECT `product_id` FROM `product` WHERE `sku` = '$prod_sku' LIMIT 1), 0) AND `customer_group_id` = $customer_group_id AND `date_end` >= $date_end;\n";

			fwrite($fp, $sql);
			$this->count_something++;
		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createInsertSqlInSeoUrlTable(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_id = mb_strtoupper($coordinate_prod_id);
		$coordinate_product_seo_url = mb_strtoupper($coordinate_product_seo_url);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_id = trim($worksheet->getCell($coordinate_prod_id . $row_number)->getValue());
			$product_seo_url = trim($worksheet->getCell($coordinate_product_seo_url . $row_number)->getValue());

			if (empty($prod_id) || empty($product_seo_url) && !is_numeric($product_seo_url)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! \n\r";
				continue;
			}

			$sql = "UPDATE `oc_seo_url` SET `query` = 'product_id=$prod_id', `store_id` = 0, `language_id` = 1, `keyword` = '$product_seo_url' WHERE NOT EXISTS(SELECT * FROM oc_seo_url WHERE `query` = 'product_id=$prod_id' AND `language_id` = 1);\n";
//			вину0м-ua
			fwrite($fp, $sql);
			$this->count_something++;
		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createUpdateAndDeleteSqlInProductSpecials(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_sku = mb_strtoupper($coordinate_prod_sku);
		$coordinate_special_price = mb_strtoupper($coordinate_special_price);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		if (empty($date_end)) $date_end = 'NOW()';
		else $date_end = "'$date_end'";

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_sku = trim($worksheet->getCell($coordinate_prod_sku . $row_number)->getValue());
			$special_price = (float)trim($worksheet->getCell($coordinate_special_price . $row_number)->getValue());

			if (empty($prod_sku) || !is_numeric($special_price)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! - prod_sku: $prod_sku; special_price: $special_price \n\r";
			} else if (empty($special_price)) {
				$sql = "DELETE FROM `$db_prefix" . "product_special` WHERE `product_id` = IFNULL((SELECT `product_id` FROM `$db_prefix" . "product` WHERE `sku` = '$prod_sku' LIMIT 1), 0) AND `customer_group_id` = $customer_group_id AND `date_end` >= $date_end;\n";

				fwrite($fp, $sql);
				$this->count_something++;
			} else {
				$sql = "UPDATE `$db_prefix" . "product_special` SET `price` = $special_price WHERE `product_id` = IFNULL((SELECT `product_id` FROM `$db_prefix" . "product` WHERE `sku` = '$prod_sku' LIMIT 1), 0) AND `customer_group_id` = $customer_group_id AND `date_end` >= $date_end;\n";

				fwrite($fp, $sql);
				$this->count_something++;
			}
		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @param array $products
	 * @return void
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function sumAndUniqueArrayOfProducts(array $parameters, array $products) : void
	{
		extract($parameters);

		$new_arr_prods = [];
		$checked_prods = [];

		foreach ($products as $product) {
			if (!in_array($product['product_id'], $checked_prods)) {
				$checked_prods[] = $product['product_id'];
				$new_arr_prods[$product['product_id']] = $product;
			} else {
				$new_arr_prods[$product['product_id']]['count_orders'] += (float)$product['count_orders'];
				$new_arr_prods[$product['product_id']]['total_sum_orders'] += (float)$product['total_sum_orders'];
			}
		}

		unset($checked_prods);

		$call_back = function ($product_one, $product_two) {
			if ($product_one['count_orders'] == $product_two['count_orders']) return 0;

			return ($product_one['count_orders'] > $product_two['count_orders']) ? -1 : 1;
		};

		usort($new_arr_prods, $call_back);

		$product_arrays = array_chunk($new_arr_prods, (int)$length_for_chunk);

		$new_worksheet = $this->spreadsheet_one->getActiveSheet();

		$new_worksheet->setCellValue('A1', 'product_id');
		$new_worksheet->setCellValue('B1', 'sku');
		$new_worksheet->setCellValue('C1', 'name');
		$new_worksheet->setCellValue('D1', 'date_modified');
		$new_worksheet->setCellValue('E1', 'order_status');
		$new_worksheet->setCellValue('F1', 'count_orders');
		$new_worksheet->setCellValue('G1', 'total_sum_orders');

		$this->count_something = 2;

		foreach ($product_arrays[0] as $product) {
			$new_worksheet->setCellValue('A' . $this->count_something, $product['product_id']);
			$new_worksheet->setCellValue('B' . $this->count_something, $product['sku']);
			$new_worksheet->setCellValue('C' . $this->count_something, $product['name']);
			$new_worksheet->setCellValue('D' . $this->count_something, $product['date_modified']);
			$new_worksheet->setCellValue('E' . $this->count_something, $product['order_status']);
			$new_worksheet->setCellValue('F' . $this->count_something, $product['count_orders']);
			$new_worksheet->setCellValue('G' . $this->count_something, $product['total_sum_orders']);

			$this->count_something++;
		}

		if (!file_exists($folder_excel)) mkdir($folder_excel);

		unset($new_worksheet);

		$this->writeExcelFile($folder_excel, $this->file_name_one, $new_spreadsheet);
	}

	/**
	 * Searched identical SKU in both files and write MODEL of product in new EXCEL file and SQL query where we are
	 * take MODEL from second file if it not empty
	 *
	 * @param array $parameters
	 * @return void
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function findProductModelAndHisDuplicateOfSku(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_id_one = mb_strtoupper($coordinate_prod_id_one);
		$coordinate_prod_sku_one = mb_strtoupper($coordinate_prod_sku_one);
		$coordinate_prod_name_one = mb_strtoupper($coordinate_prod_name_one);
		$coordinate_prod_sku_two = mb_strtoupper($coordinate_prod_sku_two);
		$coordinate_prod_model_two = mb_strtoupper($coordinate_prod_model_two);
		$coordinate_prod_name_two = mb_strtoupper($coordinate_prod_name_two);

		$new_spreadsheet = new Spreadsheet();
		$new_worksheet = $new_spreadsheet->getActiveSheet();

		$products_form_file_two = [];
		$arr_of_duplicates = [];

		if ($sheet_name_one) {
			$worksheet_one = $this->spreadsheet_one->getSheetByName($sheet_name_one);
		} else {
			$worksheet_one = $this->spreadsheet_one->getActiveSheet();
		}

		if ($sheet_name_two) {
			$worksheet_two = $this->spreadsheet_two->getSheetByName($sheet_name_two);
		} else {
			$worksheet_two = $this->spreadsheet_two->getActiveSheet();
		}

		$highest_row_one = $worksheet_one->getHighestRow();
		$highest_row_two = $worksheet_two->getHighestRow();
		$empty_rows = '';

		$new_worksheet->setCellValue('A1', 'Product Id');
		$new_worksheet->setCellValue('B1', 'Product Name (nicktoys.com.ua)');
		$new_worksheet->setCellValue('C1', 'Product Name (platoshka.com.ua)');
		$new_worksheet->setCellValue('D1', 'Product SKU');
		$new_worksheet->setCellValue('E1', 'Product Model');
		$new_worksheet->setCellValue('F1', 'Duplicate products');

		$this->count_something = 2;

		for ($row_number_two = 2; $row_number_two <= $highest_row_two; $row_number_two++) {
			$prod_name_two = trim($worksheet_two->getCell($coordinate_prod_name_two . $row_number_two)->getValue());
			$prod_sku_two = trim($worksheet_two->getCell($coordinate_prod_sku_two . $row_number_two)->getValue());
			$prod_model_two = trim($worksheet_two->getCell($coordinate_prod_model_two . $row_number_two)->getValue());

			if (empty($prod_sku_two)) {
				$empty_rows .= "В строкі №$row_number_two щось відсутнє!!! \n\r";
				continue;
			}

			if (empty($prod_model_two)) {
				$empty_rows .= "В строці №$row_number_two відсутня модель!!! \n\r";
				continue;
			}

			if (!isset($products_form_file_two[$prod_sku_two])) {
				$products_form_file_two[$prod_sku_two] = [
					'name' => $prod_name_two,
					'sku' => $prod_sku_two,
					'model' => $prod_model_two,
				];
			} else {
				$arr_of_duplicates[$prod_sku_two . "_$this->count_something"] = [
					'name' => $prod_name_two,
					'sku' => $prod_sku_two,
					'model' => $prod_model_two,
				];

				$this->count_something++;
			}
		}

		$this->count_something = 2;

		for ($row_number_one = 2; $row_number_one <= $highest_row_one; $row_number_one++) {
			$prod_id_one = (int)trim($worksheet_one->getCell($coordinate_prod_id_one . $row_number_one)->getValue());
			$prod_name_one = trim($worksheet_one->getCell($coordinate_prod_name_one . $row_number_one)->getValue());
			$prod_sku_one = trim($worksheet_one->getCell($coordinate_prod_sku_one . $row_number_one)->getValue());

			if (empty($prod_sku_one) || empty($prod_id_one) && !is_numeric($prod_id_one)) {
				$empty_rows .= "В строкі №$row_number_one щось відсутнє!!! \n\r";
				continue;
			}

			if (isset($products_form_file_two[$prod_sku_one])) {
				$new_worksheet->setCellValue('A' . $this->count_something, $prod_id_one);
				$new_worksheet->setCellValue('B' . $this->count_something, $prod_name_one);
				$new_worksheet->setCellValue('C' . $this->count_something, $products_form_file_two[$prod_sku_one]['name']);
				$new_worksheet->setCellValue('D' . $this->count_something, $prod_sku_one);
				$new_worksheet->setCellValue('E' . $this->count_something, $products_form_file_two[$prod_sku_one]['model']);

				$this->count_something++;
			}

			foreach ($arr_of_duplicates as $prod_sku => $product) {
				if ($prod_sku_one == preg_replace('/_\d+$/ui', '', $prod_sku)) {
					$new_worksheet->setCellValue('A' . $this->count_something, $prod_id_one);
					$new_worksheet->setCellValue('B' . $this->count_something, $prod_name_one);
					$new_worksheet->setCellValue('C' . $this->count_something, $product['name']);
					$new_worksheet->setCellValue('D' . $this->count_something, $prod_sku_one);
					$new_worksheet->setCellValue('E' . $this->count_something, $product['model']);
					$new_worksheet->setCellValue('F' . $this->count_something, "duplicate $this->count_something");

					$this->count_something++;
				}
			}
		}

		$this->writeExcelFile($path_to_new_excel, $new_excel_file_name, $new_spreadsheet);

		$new_spreadsheet->disconnectWorksheets();
		unset($new_spreadsheet, $new_worksheet);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * Searched identical SKU in both files and write MODEL of product in new EXCEL file and SQL query where we are
	 * take MODEL from second file if it not empty
	 *
	 * @param array $parameters
	 * @return void
	 * @throws \PhpOffice\PhpSpreadsheet\Writer\Exception
	 */
	public function createInsertSqlInProductModel(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_id_one = mb_strtoupper($coordinate_prod_id_one);
		$coordinate_prod_sku_one = mb_strtoupper($coordinate_prod_sku_one);
		$coordinate_prod_name_one = mb_strtoupper($coordinate_prod_name_one);

		$coordinate_prod_sku_two = mb_strtoupper($coordinate_prod_sku_two);
		$coordinate_prod_model_two = mb_strtoupper($coordinate_prod_model_two);
		$coordinate_prod_name_two = mb_strtoupper($coordinate_prod_name_two);

		$new_spreadsheet = new Spreadsheet();
		$new_worksheet = $new_spreadsheet->getActiveSheet();

		$fp = fopen($sql_file, $mode_write);

		$products_form_file_two = [];

		if ($sheet_name_one) {
			$worksheet_one = $this->spreadsheet_one->getSheetByName($sheet_name_one);
		} else {
			$worksheet_one = $this->spreadsheet_one->getActiveSheet();
		}

		if ($sheet_name_two) {
			$worksheet_two = $this->spreadsheet_two->getSheetByName($sheet_name_two);
		} else {
			$worksheet_two = $this->spreadsheet_two->getActiveSheet();
		}

		$highest_row_one = $worksheet_one->getHighestRow();
		$highest_row_two = $worksheet_two->getHighestRow();
		$empty_rows = '';

		$new_worksheet->setCellValue('A1', 'Product Id');
		$new_worksheet->setCellValue('B1', 'Product Name (nicktoys.com.ua)');
		$new_worksheet->setCellValue('C1', 'Product Name (platoshka.com.ua)');
		$new_worksheet->setCellValue('D1', 'Product SKU');
		$new_worksheet->setCellValue('E1', 'Product Model');

		$this->count_something = 2;

		for ($row_number_two = 2; $row_number_two <= $highest_row_two; $row_number_two++) {
			$prod_name_two = trim($worksheet_two->getCell($coordinate_prod_name_two . $row_number_two)->getValue());
			$prod_sku_two = trim($worksheet_two->getCell($coordinate_prod_sku_two . $row_number_two)->getValue());
			$prod_model_two = trim($worksheet_two->getCell($coordinate_prod_model_two . $row_number_two)->getValue());

			$products_form_file_two[$prod_model_two] = [
				'name' => $prod_name_two,
				'sku' => $prod_sku_two,
				'model' => $prod_model_two,
			];
		}

		$this->count_something = 2;

		for ($row_number_one = 2; $row_number_one <= $highest_row_one; $row_number_one++) {
			$prod_id_one = (int)trim($worksheet_one->getCell($coordinate_prod_id_one . $row_number_one)->getValue());
			$prod_name_one = trim($worksheet_one->getCell($coordinate_prod_name_one . $row_number_one)->getValue());
			$prod_sku_one = trim($worksheet_one->getCell($coordinate_prod_sku_one . $row_number_one)->getValue());
			$prod_model_one = trim($worksheet_one->getCell($coordinate_prod_model_one . $row_number_one)->getValue());

			if (isset($products_form_file_two[$prod_model_one]) && $prod_sku_one == $products_form_file_two[$prod_model_one]['sku']) {
				$new_worksheet->setCellValue('A' . $this->count_something, $prod_id_one);
				$new_worksheet->setCellValue('B' . $this->count_something, $prod_name_one);
				$new_worksheet->setCellValue('C' . $this->count_something, $products_form_file_two[$prod_model_one]['name']);
				$new_worksheet->setCellValue('D' . $this->count_something, $prod_sku_one);
				$new_worksheet->setCellValue('E' . $this->count_something, $products_form_file_two[$prod_model_one]['model']);

				$this->count_something++;

				$sql = "UPDATE `{$db_prefix}product` SET `model` = '$prod_model_one' WHERE product_id = $prod_id_one; \n";
				fwrite($fp, $sql);
			}
		}

		fclose($fp);

		$this->writeExcelFile($path_to_new_excel, $new_excel_file_name, $new_spreadsheet);

		$new_spreadsheet->disconnectWorksheets();
		unset($new_spreadsheet, $new_worksheet);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createInsertSqlInProductModel_2(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_id = mb_strtoupper($coordinate_prod_id);
		$coordinate_prod_model = mb_strtoupper($coordinate_prod_model);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_id = trim($worksheet->getCell($coordinate_prod_id . $row_number)->getValue());
			$prod_model = trim($worksheet->getCell($coordinate_prod_model . $row_number)->getValue());

			if (empty($prod_id) || empty($prod_model)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! \n\r";
				continue;
			}

			$sql = "UPDATE `{$db_prefix}product` SET `model` = '$prod_model' WHERE product_id = $prod_id;\n";
			fwrite($fp, $sql);
			$this->count_something++;
		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createUpdateSqlInProductImages(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_sku = mb_strtoupper($coordinate_prod_sku);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$prod_image = 'catalog/products/craft-box_v.2.jpg';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_sku = trim($worksheet->getCell($coordinate_prod_sku . $row_number)->getValue());

			if (empty($prod_sku)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! - prod_sku: $prod_sku; \n\r";
				continue;
			}

			$sql = "UPDATE `{$db_prefix}product_image` SET image = '$prod_image' WHERE `product_id` = IFNULL((SELECT `product_id` FROM `{$db_prefix}product` WHERE `sku` = '$prod_sku' LIMIT 1), 0) AND `sort_order` = 1;\n";
			fwrite($fp, $sql);
			$this->count_something++;
		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createUpdateSqlInProductImages2(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_mpn = mb_strtoupper($coordinate_prod_mpn);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$arr_prod_images = [];
		$fp = fopen($sql_file, $mode_write);

		$prod_images = scandir($path_to_product_images);

		foreach ($prod_images as $image) {
			if ($image == '.' || $image == '..') {
				continue;
			}

			$image_name = basename($image, '_00.jpg');
			$arr_prod_images[$image_name] = 'catalog/products/img_products/' . $image;
		}

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_mpn = trim($worksheet->getCell($coordinate_prod_mpn . $row_number)->getValue());

			if (empty($prod_mpn)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! - prod_mpn: $prod_mpn; \n\r";
				continue;
			}

			if (isset($arr_prod_images[$prod_mpn])) {
				$sql = "UPDATE `{$db_prefix}product` SET image = '$arr_prod_images[$prod_mpn]' WHERE `mpn` = '$prod_mpn';\n";
				fwrite($fp, $sql);
				$this->count_something++;
			}
		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createInsertSqlForProductOption(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_sku = mb_strtoupper($coordinate_prod_sku);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		if ($is_platoshka_db) {
			$table_property = 'model';
		} else {
			$table_property = 'sku';
		}

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_sku = trim($worksheet->getCell($coordinate_prod_sku . $row_number)->getValue());

			if (empty($prod_sku)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! - prod_sku: $prod_sku; \n\r";
				continue;
			}

			$sql = "INSERT INTO {$db_prefix}product_option (product_id, option_id, `value`, required) VALUES(IFNULL((SELECT product_id FROM {$db_prefix}product WHERE $table_property = '$prod_sku' LIMIT 1), 0), $option_id, '', 0);\n";

			fwrite($fp, $sql);
			$this->count_something++;
		}

		unset($prod_sku);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_sku = trim($worksheet->getCell($coordinate_prod_sku . $row_number)->getValue());

			if (empty($prod_sku)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! - prod_sku: $prod_sku; \n\r";
				continue;
			}

			$sql = "INSERT INTO {$db_prefix}product_option_value 
    					(product_option_id, product_id, option_id, option_value_id, quantity, subtract, price, 
    					 price_prefix, points, points_prefix, weight, weight_prefix) 
	
					VALUES(
						IFNULL((SELECT product_option_id FROM {$db_prefix}product_option 
									WHERE product_id = IFNULL(
											(SELECT product_id FROM {$db_prefix}product WHERE $table_property = '$prod_sku' LIMIT 1), 0) 
									  AND option_id = $option_id LIMIT 1)
							, 0),
							
						IFNULL((SELECT {$db_prefix}product_id FROM {$db_prefix}product WHERE $table_property = '$prod_sku' LIMIT 1), 0),
						$option_id, $option_value_id, $quantity, $subtract, $price, '$price_prefix', $weight, '$points_prefix', $points, '$points_prefix'
					);\n";

			fwrite($fp, $sql);
			$this->count_something++;
		}

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
	}

	/**
	 * @param array $parameters
	 * @return void
	 */
	public function createInsertSqlInProductImages(array $parameters) : void
	{
		extract($parameters);

		$coordinate_prod_id = mb_strtoupper($coordinate_prod_id);

		if ($sheet_name) {
			$worksheet = $this->spreadsheet_one->getSheetByName($sheet_name);
		} else {
			$worksheet = $this->spreadsheet_one->getActiveSheet();
		}

		$highest_row = $worksheet->getHighestRow();
		$empty_rows = '';
		$fp = fopen($sql_file, $mode_write);

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_id = trim($worksheet->getCell($coordinate_prod_id . $row_number)->getValue());

			if (empty($prod_id)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! \n\r";
				continue;
			}

			$sql = "INSERT INTO `{$db_prefix}product_image` (`product_id`, `image`, `sort_order`) VALUES($prod_id, '$image', '$sort_order');\n";

			fwrite($fp, $sql);
			$this->count_something++;
		}

		fwrite($fp, $sql);
		$this->count_something++;

		fclose($fp);

		echo "Було записано $this->count_something строк\n\r $empty_rows";
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

		if ($this->csv_one instanceof Csv) unset($this->csv_one);
		if ($this->csv_two instanceof Csv) unset($this->csv_two);
	}
}

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');
$file = $path . '/products-strateg.ua-13.04.2023.xlsx';

$parameters = [
	'coordinate_prod_id' => 'a',
	'image' => 'catalog/1Cproducts/craft-box_v.2.jpg',
	'sort_order' => 1,
	'sql_file' => $path . '/insert-in-product-image-strateg.ua-13.04.2023.sql',
	'sheet_name' => '',
	'db_prefix' => '',
	'mode_write' => 'w',
];

try {
	$excel = new PhpExcel($file);
	$excel->createInsertSqlInProductImages($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');
$file = $path . '/products-sku-without-some-options-avstore.com.ua-13.04.2023.xlsx';

$parameters = [
	'coordinate_prod_sku' => 'a',
	'option_id' => 17,
	'option_value_id' => 53,
	'quantity' => 100,
	'subtract' => 0,
	'price' => 50,
	'price_prefix' => '+',
	'points' => 0,
	'points_prefix' => '+',
	'weight' => 0,
	'weight_prefix' => '+',
	'sql_file' => $path . '/insert-in-product-option-avstore.com.ua-13.04.2023.sql',
	'sheet_name' => '',
	'db_prefix' => '',
	'is_platoshka_db' => false,
	'mode_write' => 'w',
];

try {
	$excel = new PhpExcel($file);
	$excel->createInsertSqlForProductOption($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\сделать что-то\platoshka.com.ua\17.02.2023');
$file = $path . '/Кода.xlsx';

$parameters = [
	'coordinate_prod_mpn' => 'a',
	'sql_file' => $path . '/update-main-products-images-platoshka.com.ua-17.02.2023.sql',
	'sheet_name' => '',
	'path' => $path . '/',
	'path_to_product_images' => $path . '/product_images',
	'db_prefix' => 'oc_',
	'mode_write' => 'w+',
];

try {
	$excel = new PhpExcel($file);
	$excel->createUpdateSqlInProductImages2($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\сделать что-то\nicktoys.com.ua\16.02.2023');
$file = $path . '/Артикула картин без лака.xlsx';

$parameters = [
	'coordinate_prod_sku' => 'a',
	'sql_file' => $path . '/update-products-images-nicktoys.com.ua-16.02.2023.sql',
	'sheet_name' => '',
	'db_prefix' => 'gg_',
	'mode_write' => 'w+',
];

try {
	$excel = new PhpExcel($file);
	$excel->createUpdateSqlInProductImages($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');

$file_one = $path . '/правки givno-products-from-nicktoys.com.ua-09.01.2023.xlsx';

$parameters = [
	'coordinate_prod_id' => 'b',
	'coordinate_prod_model' => 'e',

	'sql_file' => $path . '/update-product-model-nicktoys.com.ua-2.sql',
	'sheet_name' => '',
	'db_prefix' => 'gg_',
	'mode_write' => 'w',
];

try {
	$excel = new PhpExcel($file_one, false);
	$excel->createInsertSqlInProductModel_2($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');
$csv_file_one = $path . '/products-from-nicktoys.com.ua-09.01.2023.csv';
$csv_file_two = $path . '/products-from-platoshka.com.ua-09.01.2023.csv';
//$csv_file_two = $path . '/Стратег+технок.xls';

$csv_parameters = [
	'input_encoding_one' => 'UTF-8',
	'input_encoding_two' => 'UTF-8',
	'delimiter_one' => ';',
	'delimiter_two' => ';',
];

$parameters = [
	'coordinate_prod_id_one' => 'a',
	'coordinate_prod_sku_one' => 'd',
	'coordinate_prod_name_one' => 'b',

	'coordinate_prod_sku_two' => 'b',
	'coordinate_prod_model_two' => 'c',
	'coordinate_prod_name_two' => 'a',

	'new_excel_file_name' => 'all-products-from-nicktoys.com.ua.xlsx',
	'path_to_new_excel' => $path,
	'sheet_name_one' => '',
	'sheet_name_two' => '',
];

try {
	$excel = new PhpExcel($csv_file_one, false, $csv_file_two);

	try {
		$excel->findProductModelAndHisDuplicateOfSku($parameters);
	} catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
		echo $e->getMessage();
	}
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
	echo $e->getMessage();
}

$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');

$file_one = $path . '/new_all-products-from-nicktoys.com.ua.xlsx';
$file_two = $path . '/правки givno-products-from-nicktoys.com.ua-09.01.2023.xlsx';

$parameters = [
	'coordinate_prod_id_one' => 'a',
	'coordinate_prod_sku_one' => 'd',
	'coordinate_prod_name_one' => 'b',
	'coordinate_prod_model_one' => 'e',

	'coordinate_prod_sku_two' => 'c',
	'coordinate_prod_model_two' => 'e',
	'coordinate_prod_name_two' => 'a',

	'sql_file' => $path . '/update-product-model-nicktoys.com.ua-2.sql',
	'new_excel_file_name' => 'unique-products-from-nicktoys.com.ua-2.xlsx',
	'path_to_new_excel' => $path,
	'sheet_name_one' => '',
	'sheet_name_two' => '',
	'db_prefix' => 'gg_',
	'mode_write' => 'w',
];

try {
	$excel = new PhpExcel($file_one, false, $file_two);

	try {
		$excel->createInsertSqlInProductModel($parameters);
	} catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
		echo $e->getMessage();
	}
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');
$file = $path . '/best-sale-toys-strateg.ua-from-2022.09.01-to-2023.01.05.xlsx';

$parameters = [
	'folder_excel' => $path,
	'length_for_chunk' => 100,
];

$excel = new PhpExcel($file, true);

require_once './1.php';

/** @var array $best_sale_toys_strateg_ua *\/
$products = $best_sale_toys_strateg_ua;

try {
	$excel->sumAndUniqueArrayOfProducts($parameters, $products);
} catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\добавление переоценки\nicktoys.com.ua\09.12.2022');
$file = $path . '/Зниження_цін_до_стандартних_для_ніктойс_+_акційні_ціни_09_12_2022.xlsx';

$excel = new PhpExcel($file);

$parameters = [
	'coordinate_prod_sku' => 'b',
	'coordinate_special_price' => 'e',
	'coordinate_prod_price' => 'd',
	'customer_group_id' => 1,
	'date_end' => '',
	'sql_file' => $path . '/update-and-delete-prices-in-specials-and-products-tables.sql',
	'sheet_name' => 'Лист1',
	'db_prefix' => 'gg_',
	'mode_write' => 'a+',
];

$excel->createUpdateAndDeleteSqlInProductSpecials($parameters);
$excel->createUpdateSqlForProductPrice($parameters);*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');
$file = $path . '/new_product-names-from-platoshka.com.ua-08.11.2022.xlsx';

$excel = new PhpExcel($file);

$parameters = [
	'coordinate_prod_id' => 'b',
	'coordinate_product_seo_url' => 'e',
	'sql_file' => $path . '/update-products-price-nicktoys.com.ua-06.12.2022.sql',
	'sheet_name' => 'product-names',
	'mode_write' => 'w+',
];

$excel->createInsertSqlInSeoUrlTable($parameters);*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\добавление переоценки\strateg.ua_06.12.2022');
$file = $path . '/Переоцінка стратег 05.12.2022.xlsx';

$excel = new PhpExcel($file);

$parameters = [
	'coordinate_prod_sku' => 'a',
	'coordinate_special_price' => 'e',
	'customer_group_id' => 1,
	'date_end' => '',
	'sql_file' => $path . '/update-product-special-strateg.ua-06.12.2022.sql',
	'sheet_name' => 'Лист1',
	'mode_write' => 'w',
];

$excel->createUpdateSqlInProductSpecials($parameters);*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');
$file = $path . '/product_ids_from_v22.strateg.ua_05.12.2022.xlsx';

$parameters = [
	'coordinate_prod_id' => 'a',
	'sql_file' => $path . '/rating-for-products-v22.strateg.ua-05.12.2022.sql',
	'sheet_name' => null,
	'mode_write' => 'w',
];

$excel = new PhpExcel($file);

try {
	$excel->createSqlForProductRating($parameters);
} catch (Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно\Новая папка');
$file = $path . '/ua__nicktoys.com.ua_2023-04-11_15-45-51_5600.xlsx';

$parameters = [
	'folder_excels' => $path . '/created-attributes/',
	'sheet_name' => 'Attribute',
];

try {
	$excel = new PhpExcel($file);
	$excel->createAttributes($parameters);
} catch (Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\добавление акционных цен\avstore.com.ua\11.04.2023');
$file = $path . '/30х40_нова_та_акціна_products_paintings_by_numbers_avstore_com_ua.xlsx';

$parameters = [
	'coordinate_prod_sku' => 'c',
	'coordinate_prod_price' => 'd',
	'sql_file' => $path . '/update-products-price-avstore.com.ua-11.04.2023.sql',
	'sheet_name' => '',
	'db_prefix' => '',
	'is_platoshka_db' => false,
	'mode_write' => 'w',
];

try {
	$excel = new PhpExcel($file);
	$excel->createUpdateSqlForProductPrice($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\сделать что-то\strateg.ua\17.02.2023');
$file = $path . '/Алмазка.xls';

$parameters = [
	'coordinate_prod_sku' => 'c',
	'coordinate_cat_name_level_1' => 'f',
	'coordinate_cat_name_level_2' => 'h',
	'main_parent_cat_id' => 4,
	'sql_file' => $path . '/products-in-cats-paintings-by-number-strateg.ua-17.02.2023.sql',
	'sheet_name' => 'Товар тут',
	'db_prefix' => '',
	'mode_write' => 'w',
];

try {
	$excel = new PhpExcel($file);
	$excel->createInsertInProductCategory($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
	echo $e->getMessage();
}*/

$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');
$file = $path . '/картины_на_ИМ_Платошка_с_остатком_1_шт.xls';

$parameters = [
	'coordinate_prod_sku' => 'a',
	'main_cat_id' => 0,
	'sub_cat_id' => 201,
	'sql_file' => $path . '/products-for-sale-category-avstore.com.ua-14.04.2023.sql',
	'sheet_name' => '',
	'db_prefix' => '',
	'mode_write' => 'w',
];

try {
	$excel = new PhpExcel($file);
	$excel->createInsertInProductCategory_2($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Reader\Exception $e) {
	echo $e->getMessage();
}

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');
$file = $path . '/Остаток на основном складе.xlsx';

$platoshka_toys_cat_ids = '
	219,220,221,222,223,224,225,226,227,228,229,230,231,232,233,234,235,236,237,238,239,240,241,242,243,244,245,246,
	247,248,249,250,251,252,253,254,255,256,257,258,259,260,261,262,263,264,265,266,267,268,269,270,271,272,273,274,
	275,276,277,278,279,280,281,282,283,284,285,286,287,288,289,290,291,292,293,294,295,296,297,298,299,300,301,302,
	303,304,305,306,307,308,309,310,311,312,313,314,315,389,401,403,405,406,407,408,409,410,411,412,413,418,419,424,
	425,436
';

$nicktoys_toys_cat_ids = '
	93,94,95,96,97,98,99,100,101,102,103,104,105,106
';

$parameters = [
	'coordinate_prod_sku' => 'a',
	'coordinate_special_price' => 'e',
	'customer_group_id' => 1,
	'priority' => 1,
	'date_start' => '2023-04-14',
	'date_end' => '2035-04-14',
	'sql_file' => $path . '/product-special-products-avstore.com.ua-14.04.2023.sql',
	'sheet_name' => '',
	'db_prefix' => '',
	'is_platoshka_db' => false,
	'mode_write' => 'w',
];

try {
	$excel = new PhpExcel($file);
	$excel->createInsertSqlInProductSpecials($parameters);
} catch (Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');
$file = $path . '/Прибрати знижку.xlsx';

$excel = new PhpExcel($file);

$parameters = [
	'coordinate_prod_sku' => 'a',
	'customer_group_id' => 1,
	'sql_file' => $path . '/create-delete-sql-from-product-special-for-all-sites-without-nicktoys-05.12.2022.sql',
	'sheet_name' => 'Лист1',
	'mode_write' => 'w',
];

$excel->createDeleteFromProductSpecials($parameters);*/