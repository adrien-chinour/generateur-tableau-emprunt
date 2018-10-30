<?php

require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;


$FILE = 'tableau_emprunt.xlsx';
$URL = "http://$_SERVER[HTTP_HOST]$_SERVER[REQUEST_URI]";

// récupération des données du formulaire
$taux = $_POST["taux"] / 12;
$annee = $_POST["annee"];
$emprunt = $_POST["montant"];

// init Spreadsheet object
$spreadsheet = new Spreadsheet();
$sheet = $spreadsheet->getActiveSheet();

// affichage des données de base
$sheet->setCellValue('A1', 'Montant de l\'emprunt');
$sheet->setCellValue('A2', 'Nombre de mensualité');
$sheet->setCellValue('A3', 'Taux d\'emprunt');

$sheet->setCellValue('B1', $emprunt);
$sheet->setCellValue('B2', $annee * 12);
$sheet->setCellValue('B3', $taux * 12);

$sheet->setCellValue('D1', 'Mois');
$sheet->setCellValue('E1', 'Mensualite');
$sheet->setCellValue('F1', 'Amortissement');
$sheet->setCellValue('G1', 'Intérêt');
$sheet->setCellValue('H1', 'Restant');

// mensualite constante
$mensualite = $emprunt * ($taux / (1 - pow((1 + $taux), -($annee * 12))));

// calcul
for ($i = 1; $i <= $annee * 12; $i++) {
    $interet = $emprunt * $taux;
    $emprunt -= $mensualite - $interet;
    $sheet->setCellValue('D' . ($i + 1), $i);
    $sheet->setCellValue('E' . ($i + 1), $mensualite);
    $sheet->setCellValue('G' . ($i + 1), $interet);
    $sheet->setCellValue('F' . ($i + 1), $mensualite - $interet);
    $sheet->setCellValue('H' . ($i + 1), $emprunt);
}

// sauvegarde dans le fichier .xlxs
$writer = new Xlsx($spreadsheet);
$writer->save($FILE);

header('Location: ' . substr($URL, 0, sizeof($url)-16) . $FILE);