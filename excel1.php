<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'No');
$sheet->setCellValue('B1', 'Jenis Pendaftaran');
$sheet->setCellValue('C1', 'Tanggal Masuk');
$sheet->setCellValue('D1', 'NIS');
$sheet->setCellValue('E1', 'Nomor Ujian');
$sheet->setCellValue('F1', 'Paud');
$sheet->setCellValue('G1', 'TK');
$sheet->setCellValue('H1', 'Nomor SKHUN');
$sheet->setCellValue('I1', 'Nomor Ijazah');
$sheet->setCellValue('J1', 'Hobi');
$sheet->setCellValue('K1', 'Cita-Cita');

$koneksi = mysqli_connect("localhost", "root", "", "psd_baru2");
$query = mysqli_query($koneksi, "SELECT * FROM data_pribadi");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $no++);
    $sheet->setCellValue('B' . $i, $row['jenis_daftar']);
    $sheet->setCellValue('C' . $i, $row['tgl_masuk']);
    $sheet->setCellValue('D' . $i, $row['nis']);
    $sheet->setCellValue('E' . $i, $row['nomor_ujian']);
    $sheet->setCellValue('F' . $i, $row['pernah_paud']);
    $sheet->setCellValue('G' . $i, $row['pernah_tk']);
    $sheet->setCellValue('H' . $i, $row['nomor_skhun']);
    $sheet->setCellValue('I' . $i, $row['nomor_ijazah']);
    $sheet->setCellValue('J' . $i, $row['hobi']);
    $sheet->setCellValue('K' . $i, $row['cita_cita']);
    $i++;
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
$writer->save('Report Data Pribadi.xlsx');
?>