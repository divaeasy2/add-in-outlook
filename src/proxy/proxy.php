<?php
// proxy.php

header("Content-Type: application/json");
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Headers: Content-Type");


if ($_SERVER["REQUEST_METHOD"] !== "POST") {
    http_response_code(405);
    echo json_encode([
        "ok" => false,
        "error" => "Method not allowed"
    ]);
    exit;
}


$input = file_get_contents("php://input");

if (!$input) {
    http_response_code(400);
    echo json_encode([
        "ok" => false,
        "error" => "Empty body"
    ]);
    exit;
}


$apiUrl = "https://remote.divy-si.fr:8443/DhsDivaltoServiceDivaApiRest/api/v1/Webhook/5DED7C6421BE4694A7D992BE08D93D2F0278797F";

// cURL
$ch = curl_init($apiUrl);
curl_setopt_array($ch, [
    CURLOPT_POST => true,
    CURLOPT_HTTPHEADER => [
        "Content-Type: application/json"
    ],
    CURLOPT_POSTFIELDS => $input,
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_SSL_VERIFYPEER => false,
    CURLOPT_SSL_VERIFYHOST => false,
    CURLOPT_TIMEOUT => 30
]);

$response = curl_exec($ch);
$error    = curl_error($ch);
$httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);

curl_close($ch);


if ($response === false) {
    http_response_code(500);
    echo json_encode([
        "ok" => false,
        "error" => "cURL error",
        "details" => $error
    ]);
    exit;
}


$decoded = json_decode($response, true);


echo json_encode([
    "ok" => true,
    "httpCode" => $httpCode,
    "isJson" => $decoded !== null,
    "raw" => $response,
    "json" => $decoded
]);
