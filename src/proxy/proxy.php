<?php
// proxy.php - REST API Token Retrieval & Action Handler

header("Content-Type: application/json");
header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Headers: Content-Type");

error_reporting(E_ALL);
ini_set('display_errors', 0);

// Log file for debugging
$logFile = __DIR__ . '/proxy_debug.log';

function logDebug($message) {
    global $logFile;
    $timestamp = date('Y-m-d H:i:s');
    file_put_contents($logFile, "[$timestamp] $message\n", FILE_APPEND);
}

// Load environment variables from .env file
$envPaths = [
    __DIR__ . '/../../.env',           // src/proxy/../../.env
    __DIR__ . '/../../../.env',        // src/proxy/../../../.env
    '/var/www/html/.env',             // Common server path
    getenv('HOME') . '/.env',          // Home directory
];

logDebug("🔍 Looking for .env file...");
$envFile = null;

foreach ($envPaths as $path) {
    logDebug("  Checking: $path");
    if (file_exists($path)) {
        $envFile = $path;
        logDebug("  ✅ Found at: $path");
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
                logDebug("  📝 Loaded env var: $key=$value");
            }
        }
    }
    logDebug("✅ .env file loaded successfully");
} else {
    logDebug("⚠️  .env file not found in any path, relying on system environment");
}

/**
 * Decrypt password from database (encrypted with AES-256-CBC)
 * The encryption key must be exactly 32 bytes
 */
function decryptPassword($encryptedPassword) {
    $encryptionKeyString = getenv('ENCRYPTION_KEY') ?: 'divalto2025';
    
    // Ensure key is exactly 32 bytes (pad with '!' if needed)
    $encryptionKey = substr(str_pad($encryptionKeyString, 32, '!'), 0, 32);
    
    // Decode from base64
    $decodedPassword = base64_decode($encryptedPassword, true);
    if ($decodedPassword === false) {
        throw new Exception("Failed to decode encrypted password");
    }
    
    // Extract IV (first 16 bytes) and encrypted data
    if (strlen($decodedPassword) < 16) {
        throw new Exception("Invalid encrypted password format");
    }
    
    $iv = substr($decodedPassword, 0, 16);
    $encrypted = substr($decodedPassword, 16);
    
    // Decrypt
    $decrypted = openssl_decrypt($encrypted, 'AES-256-CBC', $encryptionKey, OPENSSL_RAW_DATA, $iv);
    
    if ($decrypted === false) {
        throw new Exception("Failed to decrypt password: " . openssl_error_string());
    }
    
    return trim($decrypted);
}

if ($_SERVER["REQUEST_METHOD"] !== "POST") {
    http_response_code(405);
    echo json_encode(["ok" => false, "error" => "Method not allowed"]);
    exit;
}

$input = file_get_contents("php://input");
if (!$input) {
    http_response_code(400);
    echo json_encode(["ok" => false, "error" => "Empty body"]);
    exit;
}

logDebug("📨 Incoming request: " . substr($input, 0, 200));

try {
    $data = json_decode($input, true);
    
    if ($data === null) {
        throw new Exception("Invalid JSON payload");
    }

    // Check if this is a token retrieval request (no action/param fields) or action request
    $isTokenRequest = !isset($data['action']) && !isset($data['param']);
    
    if ($isTokenRequest || empty($data)) {
        // Token retrieval request - fetch credentials from database
        logDebug("🔐 Token retrieval request detected");
        
        // Database connection
        $dbHost = getenv('DB_HOST');
        $dbUser = getenv('DB_USER');
        $dbPass = getenv('DB_PASS');
        $dbName = getenv('DB_NAME');
        
        // Log all variables for debugging
        logDebug("🔗 DB_HOST: " . ($dbHost ? '✅ Set' : '❌ Not set') . " ($dbHost)");
        logDebug("🔗 DB_USER: " . ($dbUser ? '✅ Set' : '❌ Not set') . " ($dbUser)");
        logDebug("🔗 DB_PASS: " . ($dbPass ? '✅ Set' : '❌ Not set'));
        logDebug("🔗 DB_NAME: " . ($dbName ? '✅ Set' : '❌ Not set') . " ($dbName)");
        
        // Fallback if variables are not set
        if (!$dbHost) {
            logDebug("⚠️  Using fallback credentials");
            $dbHost = 'maisogv978.mysql.db';
            $dbUser = 'maisogv978';
            $dbPass = 'DivaEasy2025';
            $dbName = 'maisogv978';
        }
        
        logDebug("🔗 Attempting DB connection - Host: $dbHost, User: $dbUser, DB: $dbName");

        $conn = new mysqli($dbHost, $dbUser, $dbPass, $dbName);

        if ($conn->connect_error) {
            logDebug("❌ DB Connection Error: " . $conn->connect_error);
            throw new Exception("Database connection failed: " . $conn->connect_error);
        }

        logDebug("✅ Database connected");

        // Fetch credentials from database
        $stmt = $conn->prepare("SELECT domain, user, password, env FROM `auth-add-in` LIMIT 1");
        if (!$stmt) {
            throw new Exception("Database query failed: " . $conn->error);
        }

        $stmt->execute();
        $result = $stmt->get_result();

        if ($result->num_rows === 0) {
            throw new Exception("No credentials found in database");
        }

        $row = $result->fetch_assoc();
        
        // Decrypt the password from database
        $decryptedPassword = decryptPassword($row['password']);
        
        $creds = [
            'domain' => $row['domain'],
            'user' => $row['user'],
            'password' => $decryptedPassword,
            'env' => $row['env']
        ];

        logDebug("✅ Credentials fetched and password decrypted");

        $stmt->close();
        $conn->close();

        // Forward to Authent/Auth endpoint
        $authUrl = "https://remote.divy-si.fr:8443/DhsDivaltoServiceDivaApiRest/api/v1/Authent/Auth";
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

        $response = curl_exec($ch);
        $error = curl_error($ch);
        $httpCode = curl_getinfo($ch, CURLINFO_HTTP_CODE);
        curl_close($ch);

        if ($response === false) {
            throw new Exception("Failed to reach Auth API: " . $error);
        }

        logDebug("🔙 Auth API Response (HTTP $httpCode): " . substr($response, 0, 200));

        $decoded = json_decode($response, true);
        if (!$decoded) {
            throw new Exception("Invalid JSON response from Auth API");
        }

        if ($decoded['error'] != 0) {
            throw new Exception("Auth API returned error: " . ($decoded['error'] ?? 'Unknown'));
        }

        // Return token to client
        echo json_encode([
            "ok" => true,
            "token" => $decoded['access_token'],
            "message" => "✅ Token retrieved successfully"
        ]);
        exit;

    } else {
        // Action request - forward to WebService/Execute
        logDebug("⚡ Action request detected");
        
        $executionUrl = "https://remote.divy-si.fr:8443/DhsDivaltoServiceDivaApiRest/api/v1/WebService/Execute";
        logDebug("📤 Calling WebService/Execute: " . $executionUrl);

        $ch = curl_init($executionUrl);
        curl_setopt_array($ch, [
            CURLOPT_POST => true,
            CURLOPT_HTTPHEADER => ["Content-Type: application/json"],
            CURLOPT_POSTFIELDS => json_encode($data),
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
    }

} catch (Exception $e) {
    logDebug("❌ Error: " . $e->getMessage());
    http_response_code(500);
    echo json_encode([
        "ok" => false,
        "error" => $e->getMessage()
    ]);
}
