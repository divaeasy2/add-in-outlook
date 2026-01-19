<?php
header("Content-Type: application/json");
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Headers: Content-Type");

if ($_SERVER["REQUEST_METHOD"] !== "POST") {
    http_response_code(405);
    echo json_encode(["error"=>"Method not allowed"]);
    exit;
}

$body = file_get_contents("php://input");
if(!$body){
    http_response_code(400);
    echo json_encode(["error"=>"Empty body"]);
    exit;
}

$apiUrl = "https://remote.divy-si.fr:8443/DhsDivaltoServiceDivaApiRest/api/v1/Webhook/034CFA063EB54E99A574955F88B68828050D7209";

$ch = curl_init($apiUrl);
curl_setopt_array($ch, [
    CURLOPT_POST => true,
    CURLOPT_POSTFIELDS => $body,
    CURLOPT_RETURNTRANSFER => true,
    CURLOPT_HTTPHEADER => ["Content-Type: application/json"],
    CURLOPT_SSL_VERIFYPEER => false,
    CURLOPT_SSL_VERIFYHOST => false
]);

$response = curl_exec($ch);
$err = curl_error($ch);
curl_close($ch);

if($response === false){
    echo json_encode(["error"=>"curl error", "details"=>$err]);
    exit;
}

echo json_encode(["raw"=>$response]);
