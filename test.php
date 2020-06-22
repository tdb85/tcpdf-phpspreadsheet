<?php

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Style\Color;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;
use PhpOffice\PhpSpreadsheet\Writer\Pdf\Tcpdf;

require __DIR__ . '/vendor/autoload.php';

$spreadsheet = new Spreadsheet;
$spreadsheet->setActiveSheetIndex(0);
$worksheet = $spreadsheet->getActiveSheet();
$worksheet->setTitle("Completion Report");

$worksheet->getStyle("A1:F1")->getFont()->setBold(true);
$worksheet->setCellValue("A1", "Last Name");
$worksheet->setCellValue("B1", "First Name");
$worksheet->setCellValue("C1", "Course Name");
$worksheet->setCellValue("D1", "Version");
$worksheet->setCellValue("E1", "Completion Date");
$worksheet->setCellValue("F1", "Expiration Date");
foreach (range('A', 'F') as $column) {
    $worksheet->getColumnDimension($column)->setAutoSize(true);
}
$worksheet->freezePane("A2");
$worksheet->setAutoFilter($worksheet->calculateWorksheetDimension());

$rowNum = 2;
$results = [
    ['LAST_NAME' => 'Duck', 'FIRST_NAME' => 'Donald', 'CERT_NAME' => 'ACME Course', 'CERT_VERSION' => '1.0', 'VERSION_DT' => '2020-JAN-01', 'COMPLETE_DT' => '2020-JUN-22', 'EXPIRY_DT' => '2023-JUN-22' ],
    ['LAST_NAME' => 'Bunny', 'FIRST_NAME' => 'Bugs', 'CERT_NAME' => 'ACME Course', 'CERT_VERSION' => '1.0', 'VERSION_DT' => '2020-JAN-01', 'COMPLETE_DT' => '2020-JUN-22', 'EXPIRY_DT' => '2023-JUN-22' ],
    ['LAST_NAME' => 'Fudd', 'FIRST_NAME' => 'Elmer', 'CERT_NAME' => 'Looney Tunes', 'CERT_VERSION' => '2.0', 'VERSION_DT' => '2020-MAY-30', 'COMPLETE_DT' => '2020-JUN-03', 'EXPIRY_DT' => '2023-JUN-03' ],
];
foreach ($results as $row) {
    $worksheet->setCellValue("A{$rowNum}", $row['LAST_NAME']);
    $worksheet->setCellValue("B{$rowNum}", $row['FIRST_NAME']);
    $worksheet->setCellValue("C{$rowNum}", $row['CERT_NAME']);
    $worksheet->setCellValue("D{$rowNum}", sprintf('="%s"', rtrim(sprintf("%s, %s", $row['CERT_VERSION'], $row['VERSION_DT']), ", ")));
    $worksheet->setCellValue("E{$rowNum}", $row['COMPLETE_DT']);
    $worksheet->setCellValue("F{$rowNum}", $row['EXPIRY_DT']);
    $rowNum++;
}


$worksheet->getStyle('A1:F5')
    ->getBorders()
    ->getAllBorders()
    ->setBorderStyle(Border::BORDER_HAIR)
    ->setColor(new Color(Color::COLOR_BLUE));

$writer = new Tcpdf($spreadsheet);
$writer->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
$writer->save('output.pdf');