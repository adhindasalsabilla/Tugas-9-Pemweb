<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Nama Ayah');
$sheet->setCellValue('B1', 'Tahun Lahir Ayah');
$sheet->setCellValue('C1', 'Pendidikan Ayah');
$sheet->setCellValue('D1', 'Pekerjaan Ayah');
$sheet->setCellValue('E1', 'Penghasilan Ayah');
$sheet->setCellValue('F1', 'Kebutuhan Khusus');
$sheet->setCellValue('G1', 'Nama Ibu');
$sheet->setCellValue('H1', 'Tahun Lahir Ibu');
$sheet->setCellValue('I1', 'Pendidikan Ibu');
$sheet->setCellValue('J1', 'Pekerjaan Ibu');
$sheet->setCellValue('K1', 'Penghasilan Ibu');
$sheet->setCellValue('L1', 'Kebutuhan Khusus');

$koneksi = mysqli_connect("localhost", "root", "", "psd_baru2");
$query = mysqli_query($koneksi, "SELECT * FROM data_ortu");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $row['nama_ayah']);
    $sheet->setCellValue('B' . $i, $row['tahun_lahir_ayah']);
    $sheet->setCellValue('C' . $i, $row['pendidikan_ayah']);
    $sheet->setCellValue('D' . $i, $row['pekerjaan_ayah']);
    $sheet->setCellValue('E' . $i, $row['penghasilan_ayah']);
    $sheet->setCellValue('F' . $i, $row['khusus_ayah']);
    $sheet->setCellValue('G' . $i, $row['nama_ibu']);
    $sheet->setCellValue('H' . $i, $row['tahun_lahir_ibu']);
    $sheet->setCellValue('I' . $i, $row['pendidikan_ibu']);
    $sheet->setCellValue('J' . $i, $row['pekerjaan_ibu']);
    $sheet->setCellValue('K' . $i, $row['penghasilan_ibu']);
    $sheet->setCellValue('L' . $i, $row['khusus_ibu']);
}

$styleArray = [
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
        ],
    ],
];
$sheet->getStyle('A1:D' . ($i - 1))->applyFromArray($styleArray);

$writer = new Xlsx($spreadsheet);
$writer->save('Report Data Ortu.xlsx');
?>