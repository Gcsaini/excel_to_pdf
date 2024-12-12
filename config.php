<?php
$servername = '10.5.210.195';
$username = 'itgd_office';
$password = '!tgd_0ff111';
$dbname = "election_data";
try {
    $conn = new PDO("mysql:host=$servername;dbname=$dbname", $username, $password);
    $conn->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);
} catch (PDOException $e) {
    echo $e->getMessage();
}
