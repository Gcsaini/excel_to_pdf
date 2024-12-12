<?php

require_once 'vendor/autoload.php';
include_once 'config.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Writer\Csv;

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();

$spreadsheet = $reader->load('files\by-election.xlsx');

$sheetData = $spreadsheet->getSheetByName("Sheet2")->toArray();

if (!empty($sheetData)) {
    insertByElection($conn, $sheetData);
    echo 'data inserted into table';
} else {
    echo "Sheet is empty";
}

function insertByElection($conn, $data)
{
    $array = [];

    for ($i = 1; $i < count($data); $i++) {
        $fields = [$data[$i][0], (string)$data[$i][1], (string)$data[$i][1], $data[$i][2], (string)$data[$i][3], (string)$data[$i][3], (string)$data[$i][4], (string)$data[$i][4], (string)$data[$i][5], (string)$data[$i][6], (string)$data[$i][7]];
        array_push($array, $fields);
    }
    $stmt = $conn->prepare("INSERT INTO by_election_constituency_master (
        state_id,state,state_h,cid,cname,cname_h,candidate,candidate_h,party,status,type
      ) VALUES (?,?,?,?,?,?,?,?,?,?,?)");
    try {
        $conn->beginTransaction();
        foreach ($array as $row) {
            $stmt->execute($row);
        }
        $conn->commit();
    } catch (Exception $e) {
        $conn->rollback();
        throw $e;
    }
}
