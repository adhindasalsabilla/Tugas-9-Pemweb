<?php
require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();
$sheet->setCellValue('A1', 'Nama Lengkap');
$sheet->setCellValue('B1', 'Jenis Kelamin');
$sheet->setCellValue('C1', 'NISN');
$sheet->setCellValue('D1', 'NIK');
$sheet->setCellValue('E1', 'Tempat Lahir');
$sheet->setCellValue('F1', 'Tanggal Lahir');
$sheet->setCellValue('G1', 'Agama');
$sheet->setCellValue('H1', 'Berkebutuhan Khusus');
$sheet->setCellValue('I1', 'Alamat Jalan');
$sheet->setCellValue('J1', 'RT');
$sheet->setCellValue('K1', 'RW');
$sheet->setCellValue('L1', 'Nama Dusun');
$sheet->setCellValue('M1', 'Nama Kelurahan');
$sheet->setCellValue('N1', 'Kecamatan');
$sheet->setCellValue('O1', 'Kode Pos');
$sheet->setCellValue('P1', 'Tempat Tinggal');
$sheet->setCellValue('Q1', 'Moda Transportasi');
$sheet->setCellValue('R1', 'Nomor HP');
$sheet->setCellValue('S1', 'Nomor Telepon');
$sheet->setCellValue('T1', 'Email');
$sheet->setCellValue('U1', 'KPS/PKH/KIP');
$sheet->setCellValue('V1', 'Nomor KPS');
$sheet->setCellValue('W1', 'Kewarganegaraan');
$sheet->setCellValue('X1', 'Nama Negara');

$koneksi = mysqli_connect("localhost", "root", "", "psd_baru2");
$query = mysqli_query($koneksi, "SELECT * FROM data_lengkap");
$i = 2;
$no = 1;
while ($row = mysqli_fetch_array($query)) {
    $sheet->setCellValue('A' . $i, $row['nama_lengkap']);
    $sheet->setCellValue('B' . $i, $row['jenis_kelamin']);
    $sheet->setCellValue('C' . $i, $row['nisn']);
    $sheet->setCellValue('D' . $i, $row['nik']);
    $sheet->setCellValue('E' . $i, $row['tempat_lahir']);
    $sheet->setCellValue('F' . $i, $row['tgl_lahir']);
    $sheet->setCellValue('G' . $i, $row['agama']);
    $sheet->setCellValue('H' . $i, $row['kebutuhan_khusus']);
    $sheet->setCellValue('I' . $i, $row['alamat_jalan']);
    $sheet->setCellValue('J' . $i, $row['rt']);
    $sheet->setCellValue('K' . $i, $row['rw']);
    $sheet->setCellValue('L' . $i, $row['nama_dusun']);
    $sheet->setCellValue('M' . $i, $row['nama_kelurahan']);
    $sheet->setCellValue('N' . $i, $row['kecamatan']);
    $sheet->setCellValue('O' . $i, $row['kode_pos']);
    $sheet->setCellValue('P' . $i, $row['tempat_tinggal']);
    $sheet->setCellValue('Q' . $i, $row['moda_transport']);
    $sheet->setCellValue('R' . $i, $row['nomor_hp']);
    $sheet->setCellValue('S' . $i, $row['nomor_telp']);
    $sheet->setCellValue('T' . $i, $row['email']);
    $sheet->setCellValue('U' . $i, $row['kps_pkh_kip']);
    $sheet->setCellValue('V' . $i, $row['nomor_kps']);
    $sheet->setCellValue('W' . $i, $row['kewarganegaraan']);
    $sheet->setCellValue('X' . $i, $row['nama_negara']);
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
$writer->save('Report Data Lengkap.xlsx');
?>