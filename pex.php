<?php
include 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;

$connect = new PDO("mysql:host=localhost;dbname=PR", "root", "admin");
$query = "SELECT * FROM PR ORDER BY id DESC";
$statement = $connect->prepare($query);
$statement->execute();
$result = $statement->fetchAll(PDO::FETCH_ASSOC);
// print_r($result);

// print_r($_POST);
// exit;
if (isset($_POST["export"])) {

  if (count($_POST['skus'])) {
    $query = "SELECT order_id FROM PR WHERE sku IN ('" . implode("','", $_POST['skus']) . "') ORDER BY id DESC";
    $statement = $connect->prepare($query);
    $statement->execute();
    $orderIds = $statement->fetchAll(PDO::FETCH_COLUMN);
    // echo "<pre>";
    // print_r($orderIds);
    $query = "SELECT * FROM PR WHERE order_id IN ('" . implode("','", $orderIds) . "') ORDER BY id DESC";
    $statement = $connect->prepare($query);
    $statement->execute();
    $result = $statement->fetchAll(PDO::FETCH_ASSOC);
    // print_r($result);
    // exit;
  }


  $file = new Spreadsheet();
  $active_sheet = $file->getActiveSheet();
  $active_sheet->setCellValue('A1', 'Order Id');
  $active_sheet->setCellValue('B1', 'Sku');
  $active_sheet->setCellValue('C1', 'Qty');
  $active_sheet->setCellValue('D1', 'Type');
  $active_sheet->setCellValue('E1', 'Total');

  $count = 2;

  foreach ($result as $row) {
    $active_sheet->setCellValue('A' . $count, $row["order_id"]);
    $active_sheet->setCellValue('B' . $count, $row["sku"]);
    $active_sheet->setCellValue('C' . $count, $row["quantity"]);
    $active_sheet->setCellValue('D' . $count, $row["type"]);
    $active_sheet->setCellValue('E' . $count, $row["total"]);

    $count = $count + 1;
  }

  // Create a new worksheet called "My Data"
  $myWorkSheet = new \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet($file, 'Summary');
  $file->addSheet($myWorkSheet, 1);
  $active_sheet = $file->getSheetByName('Summary');
  $active_sheet->setCellValue('A1', 'ff');


  $writer = \PhpOffice\PhpSpreadsheet\IOFactory::createWriter($file, $_POST["file_type"]);
  $file_name = time() . '.' . strtolower($_POST["file_type"]);
  $writer->save($file_name);

  header('Content-Type: application/x-www-form-urlencoded');
  header('Content-Transfer-Encoding: Binary');
  header("Content-disposition: attachment; filename=\"" . $file_name . "\"");
  readfile($file_name);
  unlink($file_name);
  exit;
}

?>
<!DOCTYPE html>
<html>

<head>
  <title>PR</title>
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.0.1/dist/css/bootstrap.min.css" rel="stylesheet" integrity="sha384-+0n0xVW2eSR5OomGNYDnhzAbDsOXxcvSN1TPprVMTNDbiYZCxYbOOl7+AMvyTG2x" crossorigin="anonymous">
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-select/1.14.0-beta2/css/bootstrap-select.min.css" integrity="sha512-mR/b5Y7FRsKqrYZou7uysnOdCIJib/7r5QeJMFvLNHNhtye3xJp1TdJVPLtetkukFn227nKpXD9OjUc09lx97Q==" crossorigin="anonymous" referrerpolicy="no-referrer" />
</head>

<body>
  <div class="full-container">
    <br />
    <h3 align="center">PR Report</h3>
    <br />
    <div class="panel panel-default">
      <div class="panel-heading">
        <form method="post">
          <div class="row">
            <div class="col-md-6">
              <select class="selectpicker" multiple aria-label="size 3 select" name="skus[]">
                <?php
                $query = "SELECT DISTINCT sku FROM PR";
                $statement = $connect->prepare($query);
                $statement->execute();
                $skus = $statement->fetchAll(PDO::FETCH_ASSOC);
                foreach ($skus as $row) {
                ?>
                  <option value="<?php echo $row["sku"]; ?>"><?php echo $row["sku"]; ?></option>
                <?php } ?>
              </select>
            </div>
            <div class="col-md-4">
              <select name="file_type" class="form-control input-sm">
                <option value="Xlsx">Xlsx</option>
                <option value="Xls">Xls</option>
                <option value="Csv">Csv</option>
              </select>
            </div>
            <div class="col-md-2">
              <input type="submit" name="export" class="btn btn-primary btn-sm" value="Export" />
            </div>
          </div>
        </form>
      </div>
      <div class="panel-body mt-5">
        <div class="table-responsive">
          <table class="table table-striped table-bordered">
            <?php
            $header = '';
            $rows = '';

            foreach ($result as $row) {
              if ($header == '') {
                $header .= '<tr>';
                $rows .= '<tr>';
                foreach ($row as $key => $value) {
                  $header .= '<th>' . $key . '</th>';
                  $rows .= '<td>' . $value . '</td>';
                }
                $header .= '</tr>';
                $rows .= '</tr>';
              } else {
                $rows .= '<tr>';
                foreach ($row as $value) {
                  $rows .= "<td>" . $value . "</td>";
                }
                $rows .= '</tr>';
              }
            }
            echo $header . $rows; ?>
          </table>
        </div>
      </div>
    </div>
  </div>
  <br />
  <br />

  <script src="https://code.jquery.com/jquery-3.6.0.min.js" integrity="sha256-/xUj+3OJU5yExlq6GSYGSHk7tPXikynS7ogEvDej/m4=" crossorigin="anonymous"></script>
  <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.0.1/dist/js/bootstrap.bundle.min.js" integrity="sha384-gtEjrD/SeCtmISkJkNUaaKMoLD0//ElJ19smozuHV6z3Iehds+3Ulb9Bn9Plx0x4" crossorigin="anonymous"></script>
  <script src="https://cdnjs.cloudflare.com/ajax/libs/bootstrap-select/1.14.0-beta2/js/bootstrap-select.min.js" integrity="sha512-FHZVRMUW9FsXobt+ONiix6Z0tIkxvQfxtCSirkKc5Sb4TKHmqq1dZa8DphF0XqKb3ldLu/wgMa8mT6uXiLlRlw==" crossorigin="anonymous" referrerpolicy="no-referrer"></script>
</body>

</html>