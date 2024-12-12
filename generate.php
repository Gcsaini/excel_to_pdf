<?php
require_once 'vendor/autoload.php';

use Dompdf\Dompdf;

$reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
setlocale(LC_ALL, 'hi_IN.UTF-8');

$spreadsheet = $reader->load('files\Final Data.xlsx');

$data = $spreadsheet->getSheetByName("Sheet1")->toArray();


if (!empty($data)) {

    $outputDir = __DIR__ . '/generated_pdfs';
    $batchSize = 15;
    $totalRecords = count($data);

    for ($batchStart = 1; $batchStart < $totalRecords; $batchStart += $batchSize) {
        $batchEnd = min($batchStart + $batchSize - 1, $totalRecords - 1);

        for ($i = $batchStart; $i <= $batchEnd; $i++) {
            $dompdf = new Dompdf();
            $dompdf->set_option('isRemoteEnabled', true);
            $baseUrl = __DIR__ . '/assets';

            $clientId = $data[$i][0] ?? '';
            $clientName = ucwords(strtolower($data[$i][1] ?? ''));
            $clientAge = isset($data[$i][2]) ? $data[$i][2] : '';
            $clientDob = isset($data[$i][3]) ? $data[$i][3] : '';
            $clientGender = isset($data[$i][5]) ? $data[$i][5] : '';
            $clientMail = isset($data[$i][6]) ? $data[$i][6] : '';
            $value1 = isset($data[$i][8]) ? $data[$i][8] : '-';
            $value2 = isset($data[$i][9]) ? $data[$i][9] : '-';
            $value3 = isset($data[$i][10]) ? $data[$i][10] : '-';
            $value4 = isset($data[$i][11]) ? $data[$i][11] : '-';
            $value5 = isset($data[$i][12]) ? $data[$i][12] : '-';
            $value6 = isset($data[$i][13]) ? $data[$i][13] : '-';
            $comb = getTopThreeCombinations([$data[$i][8], $data[$i][9], $data[$i][10], $data[$i][11], $data[$i][12], $data[$i][13]]);
            $html = <<<HTML
            <!DOCTYPE html>
            <html>
            <head>
            <style>@page{margin:0}body{margin:0;font-family:Arial,sans-serif}.header-img,.footer-img{position:fixed;left:0;width:100%;height:40px}.header-img{top:0}.footer-img{bottom:0}table{border-collapse:collapse;width:100%;text-align:center;border:2px solid}td{padding:8px;border-left:2px solid}td.no-border{border-bottom:0}th{padding:.8em;border:2px solid;background-color:#c1c1c4;font-weight:900;font-size:16px}.watermark{position:fixed;top:0;left:0;width:100%;height:100%;z-index:-1;background-image:url('http://localhost/dpk/assets/favicon.png');background-repeat:no-repeat;background-position:center;background-size:50%;opacity:.2}.page-break{page-break-before:always}</style>
            </head>
            <body>
            <img src="http://localhost/dpk/assets/ci_2.png" class=header-img />
            <img src="http://localhost/dpk/assets/ci_2.png" class=footer-img />
            <div class=watermark></div>
            <div style="width:100%;margin:70px 60px">
            <img src="http://localhost/dpk/assets/favicon.png" height=80 width=80 />
            <div style=margin-top:-80px;margin-left:100px>
            <span style=font-weight:900;font-size:20px>CHOOSE YOUR THERAPIST LLP</span><br />
            <span style=font-size:17px>A platform where we provide a path for your<br />mental wellness.</span>
            </div>
            <div style=margin-left:240px;margin-top:70px>
            <span style=font-weight:800;font-size:20px>Assessment Report</span>
            </div>
            <br/>
            <div style="width:625px;border:2px solid;padding:10px">
            <div style=display:table;width:100%>
            <div style=display:table-row;line-height:1.7rem>
            <div style=display:table-cell;padding-right:20px><span style=font-weight:900;font-size:14px>Name: </span><span style=font-size:14px>$clientName</span></div>
            <div style=display:table-cell;padding-right:20px><span style=font-weight:900;font-size:14px>Client ID: </span><span style=font-size:14px>$clientId </span></div>
            </div>
            <div style=display:table-row;line-height:1.7rem>
            <div style=display:table-cell;font-size:14px><span style=font-weight:900;font-size:14px>Email ID: </span><span style=font-size:14px>$clientMail</span></div>
            <div style=display:table-cell;font-size:14px><span style=font-weight:900;font-size:14px>DOB: </span><span style=font-size:14px>$clientDob</span></div>
            </div>
            <div style=display:table-row;line-height:1.7rem>
            <div style=display:table-cell;font-size:14px><span style=font-weight:900;font-size:14px>Gender: </span><span style=font-size:14px>$clientGender</span></div>
            <div style=display:table-cell;font-size:14px><span style=font-weight:900;font-size:14px>Age: </span><span style=font-size:14px>$clientAge Y</span></div>
            </div>
            </div>
            </div>
            <br />
            <span style=font-weight:800;font-size:18px>Assessment Tool</span>
            <br />
            <div style=margin-top:10px>
            <span style=font-weight:800;font-size:14px>Name:</span>
            <span style=font-size:14px>Holland's RIASEC Model</span>
            </div>
            <div style=margin-bottom:20px>
            <span style=font-weight:800;font-size:15px>Purpose:</span>
            <span style=font-size:14px>
            Assess vocational interests based on six personality types:
            Realistic, Investigative,<br/> Artistic, Social, Enterprising, and Conventional.
            </span>
            </div>
            <table style=width:650px>
            <tr>
            <th>RIASEC Domain</th>
            <th>Score</th>
            <th>Interest Code</th>
            </tr>
            <tr>
            <td class=no-border style=font-weight:900;font-size:14px>REALISTIC (R)</td>
            <td class=no-border>$value1</td>
            <td rowspan=6>$comb</td>
            </tr>
            <tr>
            <td class=no-border style=font-weight:900;font-size:14px>INVESTIGATIVE (I) </td>
            <td class=no-border>$value2</td>
            </tr>
            <tr>
            <td style=font-weight:900;font-size:14px>ARTISTIC (A)</td>
            <td>$value3</td>
            </tr>
            <tr>
            <td style=font-weight:900;font-size:14px>SOCIAL (S) </td>
            <td>$value4</td>
            </tr>
            <tr>
            <td style=font-weight:900;font-size:14px>ENTERPRISING (E)</td>
            <td>$value5</td>
            </tr>
            <tr>
            <td style=font-weight:900;font-size:14px>CONVENTIONAL (C)</td>
            <td>$value6</td>
            </tr>
            </table>
            <br />
            <span style=font-size:14px>
            A detailed interpretation of each RIASEC domain and corresponding
            career recommendations is provided<br/> on the next page.
            </span>
            <br /><br />
            <div>
                <span style=font-weight:800;font-size:14px>Date: </span>
                <span style=font-size:14px>12/12/2024</span>
            </div>
            <div style=margin-top:10px;margin-left:40px>
                <img src="http://localhost/dpk/assets/ri_2_new.png" alt=sign />
            </div>
            <div style=margin-top:10px>
            <span style=font-size:18px;font-weight:900>Choose Your Therapist LLP</span>
            </div>
            <div style=margin-top:15px>
            <div style=margin-left:440px>
            <img src="http://localhost/dpk/assets/globe.png" height=20 width=20 />
            </div>
            <div style=margin-top:-45px;margin-left:464px>
            <span style=font-size:12px>www.chooseyourtherapist.in</span>
            </div>
            </div>
            </div>
            <div class=page-break></div>
            <div style="margin:100px 60px">
            <img src="http://localhost/dpk/assets/Abcd_page-0002.jpg" height=800 width="600"/>
            </div>
            </body>
            </html>
            HTML;

            $dompdf->loadHtml($html);
            $dompdf->setPaper('A4', 'portrait');
            $dompdf->render();

            $filePath = $outputDir . "/{$clientId}.pdf";
            file_put_contents($filePath, $dompdf->output());
        }
        unset($dompdf);
        gc_collect_cycles();
    }
    echo "file generated successfully";
} else {
    echo "Sheet is empty";
}


function getTopThreeCombinations($data)
{
    $headers = ['R', 'I', 'A', 'S', 'E', 'C'];
    $valuesWithHeaders = array_combine($headers, $data);

    arsort($valuesWithHeaders);

    $topHeaders = array_keys(array_slice($valuesWithHeaders, 0, 3, true));
    $combinations = [];
    for ($i = 0; $i < count($topHeaders); $i++) {
        for ($j = $i + 1; $j < count($topHeaders); $j++) {
            for ($k = $j + 1; $k < count($topHeaders); $k++) {
                $combinations[] = $topHeaders[$i] . $topHeaders[$j] . $topHeaders[$k];
            }
        }
    }

    return implode(", ", $combinations);
}
