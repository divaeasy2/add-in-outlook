<?php
// proxy_child.php - REST API for fetching linked events

header("Content-Type: application/json; charset=utf-8");
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Headers: Content-Type");

// Set default charset for all string operations
ini_set('default_charset', 'utf-8');

error_reporting(E_ALL);
ini_set('display_errors', 0);

// Log file for debugging
$logFile = __DIR__ . '/proxy_child_debug.log';

function logDebug($message) {
    global $logFile;
    try {
        $timestamp = date('Y-m-d H:i:s');
        $logMessage = "[$timestamp] $message\n";
        
        // Ensure directory exists
        if (!is_dir(dirname($logFile))) {
            @mkdir(dirname($logFile), 0777, true);
        }
        
        // Write to log file
        $bytes = @file_put_contents($logFile, $logMessage, FILE_APPEND | LOCK_EX);
        if ($bytes === false) {
            error_log("Failed to write to $logFile");
        }
    } catch (Exception $e) {
        error_log("LogDebug exception: " . $e->getMessage());
    }
}

logDebug("\n\n" . str_repeat("=", 150));
logDebug("🚀 PROXY_CHILD.PHP INITIALIZED - " . date('Y-m-d H:i:s'));
logDebug(str_repeat("=", 150));

function decryptPassword($encryptedPassword) {
    logDebug("\n════════════════════════════════════════════════════════════════════════════════════════════");
    logDebug("🔍 DECRYPTION PROCESS STARTED");
    logDebug("════════════════════════════════════════════════════════════════════════════════════════════");
    logDebug("   Encrypted Password (Base64): " . substr($encryptedPassword, 0, 50) . "...");
    
    try {
        // Standard AES-256-CBC with documented key derivation
        // Key: 'divalto2025' padded to 32 bytes with null bytes
        // IV: 16 zero bytes (deterministic)
        
        $key = 'divalto2025';
        $key_padded = str_pad($key, 32, "\0");
        $iv = str_repeat("\0", 16);
        
        logDebug("   Key Derivation: Raw key padded to 32 bytes");
        logDebug("      Key: '$key'");
        logDebug("      Key (padded hex): " . bin2hex($key_padded));
        logDebug("      IV (hex): " . bin2hex($iv));
        
        // Decrypt with automatic PKCS7 padding
        logDebug("   Decryption Method: AES-256-CBC with PKCS7 padding");
        $decrypted = openssl_decrypt(
            $encryptedPassword,
            'aes-256-cbc',
            $key_padded,
            0,  // 0 = with automatic PKCS7 padding
            $iv
        );
        
        if ($decrypted === false) {
            logDebug("   ❌ DECRYPTION FAILED: " . openssl_error_string());
            logDebug("════════════════════════════════════════════════════════════════════════════════════════════\n");
            throw new Exception("Decryption failed: " . openssl_error_string());
        }
        
        $decryptedTrimmed = trim($decrypted);
        logDebug("   ✅ DECRYPTION SUCCESSFUL");
        logDebug("      Password length: " . strlen($decryptedTrimmed) . " characters");
        logDebug("      Password preview: " . substr($decryptedTrimmed, 0, 10) . "***");
        logDebug("════════════════════════════════════════════════════════════════════════════════════════════\n");
        
        return $decryptedTrimmed;
        
    } catch (Exception $e) {
        logDebug("   ❌ Exception: " . $e->getMessage());
        logDebug("════════════════════════════════════════════════════════════════════════════════════════════\n");
        throw $e;
    }
}

if ($_SERVER["REQUEST_METHOD"] !== "POST") {
    logDebug("❌ Invalid request method: " . $_SERVER["REQUEST_METHOD"]);
    http_response_code(405);
    echo json_encode(["error"=>"Method not allowed"]);
    exit;
}

$input = file_get_contents("php://input");
logDebug("📨 REQUEST RECEIVED - Method: " . $_SERVER["REQUEST_METHOD"] . ", Content-Type: " . ($_SERVER['CONTENT_TYPE'] ?? 'N/A'));

if (!$input) {
    logDebug("⚠️  Empty request body received");
    http_response_code(400);
    echo json_encode(["error"=>"Empty body"]);
    exit;
}

logDebug("📨 Incoming request payload:");
logDebug("   " . substr($input, 0, 300) . (strlen($input) > 300 ? "..." : ""));

try {
    $data = json_decode($input, true);
    
    if ($data === null) {
        logDebug("❌ JSON decode failed: " . json_last_error_msg());
        throw new Exception("Invalid JSON payload");
    }
    
    logDebug("✅ JSON decoded successfully");

    // Load environment variables from .env file
    $envPaths = [
        __DIR__ . '/../../.env',        
        __DIR__ . '/../../../.env',    
        '/var/www/html/.env',          
        getenv('HOME') . '/.env',       
    ];

    $envFile = null;
    foreach ($envPaths as $path) {
        if (file_exists($path)) {
            $envFile = $path;
            break;
        }
    }

    if ($envFile) {
        $envLines = file($envFile, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES);
        foreach ($envLines as $line) {
            if (strpos($line, '=') !== false && $line[0] !== '#') {
                list($key, $value) = explode('=', $line, 2);
                $key = trim($key);
                $value = trim($value);
                if (!getenv($key)) {
                    putenv("$key=$value");
                }
            }
        }
    }

    // Database connection
    $dbHost = getenv('DB_HOST');
    $dbUser = getenv('DB_USER');
    $dbPass = getenv('DB_PASS');
    $dbName = getenv('DB_NAME');
    
    // Fallback if variables are not set
    if (!$dbHost) {
        $dbHost = 'localhost';
        $dbUser = 'DivyAddIN';
        $dbPass = 'i$aKtRG48hjffr0?2';
        $dbName = 'DivyADDIN';
    }

    $conn = new mysqli($dbHost, $dbUser, $dbPass, $dbName);
    if ($conn->connect_error) {
        throw new Exception("Database connection failed: " . $conn->connect_error);
    }

    // Set charset to utf8mb4 for proper character encoding
    $conn->set_charset("utf8mb4");

    // Fetch credentials from database
    $stmt = $conn->prepare("SELECT domain, user, password, env, auth_api, action_api FROM `auth-add-in` LIMIT 1");
    if (!$stmt) {
        throw new Exception("Database query failed: " . $conn->error);
    }

    $stmt->execute();
    $result = $stmt->get_result();

    if ($result->num_rows === 0) {
        throw new Exception("No credentials found in database");
    }

    $row = $result->fetch_assoc();
    
    logDebug("� DATABASE VALUES RETRIEVED:");
    logDebug("   Domain: " . $row['domain']);
    logDebug("   User: " . $row['user']);
    logDebug("   Encrypted Password (from DB): " . substr($row['password'], 0, 50) . "...");
    logDebug("   Encrypted Password Full: " . $row['password']);
    logDebug("   Env: " . $row['env']);
    logDebug("   Auth API: " . $row['auth_api']);
    logDebug("   Action API: " . $row['action_api']);
    logDebug("");
    
    logDebug("�🔐 Attempting to decrypt password from database...");
    // Decrypt the password from database
    try {
        $decryptedPassword = decryptPassword($row['password']);
        logDebug("✅ Password decrypted successfully");
    } catch (Exception $e) {
        logDebug("❌ Decryption failed: " . $e->getMessage());
        throw $e;
    }
    
    // Get API URLs from database, fallback to .env or defaults
    $authUrl = $row['auth_api'] ?: (getenv('AUTH_API'));;
    $executionUrl = $row['action_api'] ?: (getenv('ACTION_API'));
    
    if (empty($authUrl) || empty($executionUrl)) {
        throw new Exception("API URLs not configured");
    }
    
    $creds = [
        'domain' => $row['domain'],
        'user' => $row['user'],
        'password' => $decryptedPassword,
        'env' => $row['env']
    ];

    $stmt->close();
    $conn->close();

    // Step 1: Get a fresh token from Auth API
    logDebug("📤 Calling Auth API: " . $authUrl);

    $ch = curl_init($authUrl);
    curl_setopt_array($ch, [
        CURLOPT_POST => true,
        CURLOPT_HTTPHEADER => ["Content-Type: application/json"],
        CURLOPT_POSTFIELDS => json_encode($creds),
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_SSL_VERIFYPEER => false,
        CURLOPT_SSL_VERIFYHOST => false,
        CURLOPT_TIMEOUT => 30,
        CURLOPT_CONNECTTIMEOUT => 10
    ]);

    $tokenResponse = curl_exec($ch);
    $error = curl_error($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($tokenResponse === false) {
        throw new Exception("Failed to reach Auth API: " . $error);
    }

    logDebug("🔙 Auth API Response (HTTP $httpCode): " . substr($tokenResponse, 0, 200));

    $tokenData = json_decode($tokenResponse, true);
    if (!$tokenData) {
        throw new Exception("Invalid JSON response from Auth API");
    }

    if ($tokenData['error'] != 0) {
        throw new Exception("Auth API returned error: " . ($tokenData['error'] ?? 'Unknown'));
    }

    $token = $tokenData['access_token'];
    logDebug("✅ Token retrieved successfully");

    // Step 2: Build the linked events request body
    $utilizador = $data['evenement']['utilisateur'] ?? '';
    $tiers = $data['evenement']['tiers'] ?? '';
    $realizeOkValue = $data['evenement']['RealiseOk'] ?? 0;
    
    // Pass RealiseOk value directly to the API
    // 0 = available only, 1 = all events

    $linkedEventsPayload = [
        "action" => "WEB_SERVICE_INFINITY",
        "access_token" => $token,
        "param" => json_encode([
            "action" => [
                "swinfinity" => "dv_lister_evt"
            ],
            "data" => [
                "evenement" => [
                    "utilisateur" => $utilizador,
                    "tiers" => $tiers,
                    "RealiseOk" => $realizeOkValue
                ]
            ]
        ])
    ];

    logDebug("📤 Calling WebService/Execute for linked events");
    logDebug("   Utilisateur: " . $utilizador);
    logDebug("   Tiers: " . $tiers);
    logDebug("   RealiseOk Filter: " . $realizeOkValue);

    // Step 3: Call the WebService/Execute endpoint
    $ch = curl_init($executionUrl);
    curl_setopt_array($ch, [
        CURLOPT_POST => true,
        CURLOPT_HTTPHEADER => ["Content-Type: application/json"],
        CURLOPT_POSTFIELDS => json_encode($linkedEventsPayload),
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_SSL_VERIFYPEER => false,
        CURLOPT_SSL_VERIFYHOST => false,
        CURLOPT_TIMEOUT => 60,
        CURLOPT_CONNECTTIMEOUT => 10
    ]);

    $response = curl_exec($ch);
    $error = curl_error($ch);
    $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($response === false) {
        throw new Exception("Failed to reach WebService API: " . $error);
    }

    logDebug("🔙 WebService Response (HTTP $httpCode): " . substr($response, 0, 200));

    echo $response;
    exit;

} catch (Exception $e) {
    logDebug("❌ Error: " . $e->getMessage());
    http_response_code(500);
    echo json_encode([
        "error" => $e->getMessage()
    ]);
}
