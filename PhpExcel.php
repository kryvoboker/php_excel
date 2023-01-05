<?php

/** @noinspection SqlDialectInspection */
/** @noinspection SqlNoDataSourceInspection */

require_once __DIR__ . '/vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

mb_internal_encoding('UTF-8');

class PhpExcel
{
	private ?Spreadsheet $spreadsheet_one;
	private string $file_name_one;
	private int $count_something = 0;

	public function __construct(string $spreadsheet_file_name, bool $is_need_create_file = false)
	{
		if (!$is_need_create_file) {
			$this->spreadsheet_one = IOFactory::load($spreadsheet_file_name);
		} else {
			$this->spreadsheet_one = new Spreadsheet();
		}

		$this->file_name_one = basename($spreadsheet_file_name);
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

		if ($path_to_files[strlen($path_to_files) - 1] == '/') {
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

			$sql = "UPDATE `$db_prefix" . "product` SET `price` = $prod_price WHERE `product_id` = IFNULL((SELECT `product_id` FROM `$db_prefix" . "product` WHERE `sku` = '$prod_sku' LIMIT 1), 0);\n";

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
				$sql = "INSERT INTO gg_product_to_category (`product_id`, `category_id`) VALUES (IFNULL((SELECT product_id FROM gg_product WHERE sku = '$prod_sku' LIMIT 1), 0), IFNULL((SELECT cd.category_id FROM gg_category_description AS cd RIGHT JOIN gg_category AS c ON (cd.category_id = c.category_id) WHERE cd.`name` = '$cat_name_level_2' AND c.parent_id = (SELECT cd.category_id FROM gg_category AS c RIGHT JOIN gg_category_description AS cd ON (c.category_id = cd.category_id) WHERE cd.`name` = '$cat_name_level_1' AND c.parent_id = $main_parent_cat_id LIMIT 1) LIMIT 1), 0)) ON DUPLICATE KEY UPDATE `category_id` = `category_id`;\n";

				fwrite($fp, $sql);
				$this->count_something++;
			}


			if (!empty($cat_name_level_1)) {
				$sql = "INSERT INTO gg_product_to_category (`product_id`, `category_id`) VALUES (IFNULL((SELECT product_id FROM gg_product WHERE sku = '$prod_sku' LIMIT 1), 0), IFNULL((SELECT c.category_id FROM gg_category_description AS cd RIGHT JOIN gg_category AS c ON (c.category_id = cd.category_id) WHERE c.parent_id = $main_parent_cat_id AND cd.`name` = '$cat_name_level_1' LIMIT 1), 0)) ON DUPLICATE KEY UPDATE `category_id` = `category_id`;\n";

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

		for ($row_number = 2; $row_number <= $highest_row; $row_number++) {
			$prod_sku = trim($worksheet->getCell($coordinate_prod_sku . $row_number)->getValue());
			$special_price = (float)trim($worksheet->getCell($coordinate_special_price . $row_number)->getValue());

			if (empty($prod_sku) || empty($special_price) && !is_numeric($special_price)) {
				$empty_rows .= "В строкі №$row_number щось відсутнє!!! \n\r";
				continue;
			}

			$sql = "INSERT INTO `{$db_prefix}product_special` (`product_id`, `customer_group_id`, `priority`, `price`, `date_end`) VALUES(IFNULL((SELECT `product_id` FROM `{$db_prefix}product` WHERE `sku` = '$prod_sku' LIMIT 1), 0), $customer_group_id, $priority, $special_price, '$date_end');\n";

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

		$new_spreadsheet = $this->spreadsheet_one;
		$new_worksheet = $new_spreadsheet->getActiveSheet();

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

		$this->writeExcelFile($folder_excel, $this->file_name_one, $new_spreadsheet);
	}

	public function __destruct()
	{
		if ($this->spreadsheet_one instanceof Spreadsheet) {
			$this->spreadsheet_one->disconnectWorksheets();
			unset($this->spreadsheet_one);
		}
	}
}

$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\Лайно');
$file = $path . '/best-sale-toys-strateg.ua-from-2022.09.01-to-2023.01.05.xlsx';

$parameters = [
	'folder_excel' => $path,
	'length_for_chunk' => 100,
];

$excel = new PhpExcel($file, true);

require_once './1.php';

/** @var array $best_sale_toys_strateg_ua */
$products = $best_sale_toys_strateg_ua;

try {
	$excel->sumAndUniqueArrayOfProducts($parameters, $products);
} catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
	echo $e->getMessage();
}

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

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\добавление новых позицый\03.01.2023\nicktoys.com.ua');
$file = $path . '/uk-ua__nicktoys.com.ua_2023-01-02_11-56-19_4000.xlsx';

$parameters = [
	'folder_excels' => $path . '/new_excels/',
	'sheet_name' => 'Attribute',
];

$excel = new PhpExcel($file);

try {
	$excel->createAttributes($parameters);
} catch (\PhpOffice\PhpSpreadsheet\Writer\Exception $e) {
	echo $e->getMessage();
}*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\добавление переоценки\nicktoys.com.ua_06.12.2022');
$file = $path . '/Переоцінка ніктойс 05.12.2022.xlsx';

$excel = new PhpExcel($file);

$parameters = [
	'coordinate_prod_sku' => 'a',
	'coordinate_prod_price' => 'd',
	'sql_file' => $path . '/update-products-price-nicktoys.com.ua-06.12.2022.sql',
	'sheet_name' => 'Лист1',
	'mode_write' => 'w',
];

$excel->createUpdateSqlForProductPrice($parameters);*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\добавление чего-то\подкатегории-для-nicktoys.com.ua-13.12.2022');
$file = $path . '/Выгрузка_30112022_11_17_54-paintings-by-number.xls';

$excel = new PhpExcel($file);

$parameters = [
	'coordinate_prod_sku' => 'c',
	'coordinate_cat_name_level_1' => 'f',
	'coordinate_cat_name_level_2' => 'h',
	'main_parent_cat_id' => 60,
	'sql_file' => $path . '/products-in-cats-paintings-by-number-nicktoys.com.ua-13.12.2022.sql',
	'sheet_name' => null,
	'mode_write' => 'w',
];

$excel->createInsertInProductCategory($parameters);*/

/*$path = str_replace('\\', '/', 'C:\Users\User\Documents\Роман\добавление акционных цен\avstore.com.ua\16.12.2022');
$file = $path . '/АВ стор 17 процентов 16.12.2022.xlsx';

$excel = new PhpExcel($file);

$parameters = [
	'coordinate_prod_sku' => 'b',
	'coordinate_special_price' => 'e',
	'customer_group_id' => 1,
	'priority' => 1,
	'date_end' => '2023-01-01',
	'sql_file' => $path . '/product-special-painting-by-numbers-avstore.com.ua-16.12.2022.sql',
	'sheet_name' => 'Картина',
	'db_prefix' => '',
	'mode_write' => 'w',
];

$excel->createInsertSqlInProductSpecials($parameters);*/

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