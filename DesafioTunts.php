<?php

require __DIR__ . '/vendor/autoload.php';


//Start setting the Google Sheets API user
$client = new \Google_Client();
$client->setApplicationName('Sheets');
$client->setScopes([\Google_Service_Sheets::SPREADSHEETS]);
$client->setAccessType('offline');
$client->setAuthConfig(__DIR__ . '/credentials.json');
$service = new Google_Service_Sheets($client);
$spreadsheetId = "1ZDTIani5q4yqeKYfwSsL8_xcm3uaxoR9m-zdhekROAg";
//Finish setting the Google Sheets API user

$fp = fopen("log.txt", "a");

//Set the Range of Rows in the CSV file
$range = "engenharia!B4:F27";
$response = $service->spreadsheets_values->get($spreadsheetId, $range);
//Values will receive the Rows of the Google Sheet
$values = $response->getValues();

//Possible Situations for this project
$situation1 = [
    ["Aprovado"]
];

$situation2 = [
    ["Reprovado por Nota"]
];

$situation3 = [
    ["Exame Final"]
];

$situation4 = [
    ["Reprovado por Faltas"]
];



if (empty($values)) { //IF Values don't receive any value
    print "No data found.\n";
} else {
    $count = 4; //Count will represent the actual row in the loop, starts in the rows number 4



    $params = [ //Parameters for the API 
        'valueInputOption' => 'USER_ENTERED'
    ];



    foreach ($values as $row) { //loop for each row in the Sheet
        $range2 = 'engenharia!G' . $count . ':G' . $count;

        $avarage = ($row[2] + $row[3] + $row[4]) / 3;

        if ($row[1] <= 15) {

            if ($avarage >= 70) { //IF student is approved   

                $body = new Google_Service_Sheets_ValueRange([
                    'values' => $situation1
                ]);

                $result = $service->spreadsheets_values->update(
                    $spreadsheetId,
                    $range2,
                    $body,
                    $params
                );

                $range2 = 'engenharia!H' . $count . ':H' . $count;

                $FinalGrade = [
                    ["0"]
                ];

                $body = new Google_Service_Sheets_ValueRange([
                    'values' => $FinalGrade
                ]);


                $result = $service->spreadsheets_values->update(
                    $spreadsheetId,
                    $range2,
                    $body,
                    $params
                );

                echo 'Row: ' . $count - 3 . ' Aluno: ' . $row[0] . ' Situação: Aprovado    Naf: 0 ';
                echo "\r\n";

                fwrite($fp, 'Row: ' . $count - 3 . ' Aluno: ' . $row[0] . ' Situação: Aprovado    Naf: 0 ');
                fwrite($fp, "\r\n");
            } else if ($avarage >= 50) { //IF student goes to the Final Exam 
                $body = new Google_Service_Sheets_ValueRange([
                    'values' => $situation3
                ]);

                $result = $service->spreadsheets_values->update(
                    $spreadsheetId,
                    $range2,
                    $body,
                    $params
                );

                $range2 = 'engenharia!H' . $count . ':H' . $count;

                $FirstNaf = 100 - $avarage;
                $FinalNaf = ceil($FirstNaf);

                $FinalGrade = [
                    [$FinalNaf]
                ];

                $body = new Google_Service_Sheets_ValueRange([
                    'values' => $FinalGrade
                ]);


                $result = $service->spreadsheets_values->update(
                    $spreadsheetId,
                    $range2,
                    $body,
                    $params
                );

                echo 'Row: ' . $count - 3 . ' Aluno: ' . $row[0] . ' Situação: Exame Final    Naf: ' . $FinalNaf;
                echo "\r\n";
                fwrite($fp, 'Row: ' . $count - 3 . ' Aluno: ' . $row[0] . ' Situação: Exame Final    Naf: ' . $FinalNaf);
                fwrite($fp, "\r\n");
            } else { //IF student have Failed in the Class by his Grade
                $body = new Google_Service_Sheets_ValueRange([
                    'values' => $situation2
                ]);

                $result = $service->spreadsheets_values->update(
                    $spreadsheetId,
                    $range2,
                    $body,
                    $params
                );


                $range2 = 'engenharia!H' . $count . ':H' . $count;

                $FinalGrade = [
                    ["0"]
                ];

                $body = new Google_Service_Sheets_ValueRange([
                    'values' => $FinalGrade
                ]);


                $result = $service->spreadsheets_values->update(
                    $spreadsheetId,
                    $range2,
                    $body,
                    $params
                );

                echo 'Row: ' . $count - 3 . ' Aluno: ' . $row[0] . ' Situação: Reprovado por Nota    Naf: 0 ';
                echo "\r\n";
                fwrite($fp, 'Row: ' . $count - 3 . ' Aluno: ' . $row[0] . ' Situação: Reprovado por Nota    Naf: 0 ');
                fwrite($fp, "\r\n");
            }
        } else { //IF student have Failed in the Class by his lack of Presence
            $body = new Google_Service_Sheets_ValueRange([
                'values' => $situation4
            ]);

            $result = $service->spreadsheets_values->update(
                $spreadsheetId,
                $range2,
                $body,
                $params
            );

            $range2 = 'engenharia!H' . $count . ':H' . $count;

            $FinalGrade = [
                ["0"]
            ];

            $body = new Google_Service_Sheets_ValueRange([
                'values' => $FinalGrade
            ]);


            $result = $service->spreadsheets_values->update(
                $spreadsheetId,
                $range2,
                $body,
                $params
            );

            echo 'Row: ' . $count - 3 . ' Aluno: ' . $row[0] . ' Situação: Reprovado por Falta    Naf: 0';
            echo "\r\n";
            fwrite($fp, 'Row: ' . $count - 3 . ' Aluno: ' . $row[0] . ' Situação: Reprovado por Falta    Naf: 0 ');
            fwrite($fp, "\r\n");
        }

        $count += 1;
    }
}

fclose($fp);
