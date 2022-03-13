<?php

//import.php

include 'vendor/autoload.php';


try {
	$connect = new PDO("mysql:host=localhost;dbname=PR", "root", "admin", array(PDO::ATTR_ERRMODE => PDO::ERRMODE_WARNING));
	if ($_FILES["import_excel"]["name"] != '') {
		$allowed_extension = array('xls', 'csv', 'xlsx');
		$file_array = explode(".", $_FILES["import_excel"]["name"]);
		$file_extension = end($file_array);

		if (in_array($file_extension, $allowed_extension)) {
			$file_name = 'upload/' . time() . '.' . $file_extension;;
			move_uploaded_file($_FILES['import_excel']['tmp_name'], $file_name);
			$file_type = \PhpOffice\PhpSpreadsheet\IOFactory::identify($file_name);
			$reader = \PhpOffice\PhpSpreadsheet\IOFactory::createReader($file_type);

			$spreadsheet = $reader->load($file_name);
			//unlink($file_name);

			$data = $spreadsheet->getActiveSheet()->toArray();


			foreach ($data as $key => $row) {
				// print_r($row);
				// exit;
				if ($key == 0)
					continue;

				$desc = addslashes($row[5]);
				$qty = trim($row[6]) ? trim($row[6]) : 0;
				$total = str_replace(',', '', $row[24]);
				$other = str_replace(',', '', $row[23]);
				$other_transaction_fees = str_replace(',', '', $row[22]);
				$query = "
			INSERT INTO PR(`date_time`, `settlement_id`, `type`, `order_id`, `sku`, `description`, `quantity`, `marketplace`, `account_type`, `fulfillment`, `order_city`, `order_state`, `order_postal`, `product_sales`, `shipping_credits`, `promotional_rebates`, `GST`, `TCS_CGST`, `TCS_SGST`, `TCS_IGST`, `selling_fees`, `fba_fees`, `other_transaction_fees`, `other`, `total`)
			VALUES ('$row[0]','$row[1]','$row[2]','$row[3]','$row[4]','$desc',$qty,'$row[7]','$row[8]','$row[9]','$row[10]','$row[11]','$row[12]','$row[13]','$row[14]','$row[15]','$row[16]','$row[17]','$row[18]','$row[19]','$row[20]','$row[21]',$other_transaction_fees,$other,$total)
			";
				$statement = $connect->query($query);
				if (!$statement) {
					echo $query;
					echo "<br>";
				}
			}
			$message = '<div class="alert alert-success">Data Imported Successfully</div>';
		} else {
			$message = '<div class="alert alert-danger">Only .xls .csv or .xlsx file allowed</div>';
		}
	} else {
		$message = '<div class="alert alert-danger">Please Select File</div>';
	}
} catch (PDOException $e) {
	echo 'Connection failed: ' . $e->getMessage();
	exit;
}
echo $message;
