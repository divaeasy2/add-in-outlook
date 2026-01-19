<?php
/**
 * Password Encryption Tool
 * Encrypts passwords for storage in the database
 * Use the output to store in the auth-add-in table password field
 */

header('Content-Type: application/json');

// Function to encrypt password (reverse of decryptPassword in proxy.php)
function encryptPassword($password, $encryptionKey = 'divalto2025') {
    // Pad the key to 32 bytes
    $encryptionKey = substr(str_pad($encryptionKey, 32, '!'), 0, 32);
    
    // Generate a random IV (initialization vector)
    $iv = openssl_random_pseudo_bytes(16);
    
    // Encrypt the password
    $encrypted = openssl_encrypt($password, 'AES-256-CBC', $encryptionKey, OPENSSL_RAW_DATA, $iv);
    
    if ($encrypted === false) {
        throw new Exception("Encryption failed: " . openssl_error_string());
    }
    
    // Combine IV + encrypted data and base64 encode
    $encryptedPassword = base64_encode($iv . $encrypted);
    
    return $encryptedPassword;
}

// Get password from query parameter or POST
$passwordToEncrypt = null;
$encryptionKeyUsed = 'divalto2025'; // Default key

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    $input = json_decode(file_get_contents('php://input'), true);
    $passwordToEncrypt = $input['password'] ?? null;
    $encryptionKeyUsed = $input['encryption_key'] ?? 'divalto2025';
} elseif (isset($_GET['password'])) {
    $passwordToEncrypt = $_GET['password'];
    $encryptionKeyUsed = $_GET['encryption_key'] ?? 'divalto2025';
}

$result = array(
    'success' => false,
    'error' => null,
    'encrypted_password' => null,
    'instructions' => 'Send a POST request with {"password": "your-password-here", "encryption_key": "optional-custom-key"} or use ?password=your-password'
);

try {
    if (!$passwordToEncrypt) {
        throw new Exception("No password provided. Send it as 'password' parameter or in JSON body");
    }
    
    if (strlen($passwordToEncrypt) === 0) {
        throw new Exception("Password cannot be empty");
    }
    
    $encrypted = encryptPassword($passwordToEncrypt, $encryptionKeyUsed);
    
    $result['success'] = true;
    $result['encrypted_password'] = $encrypted;
    $result['original_password'] = $passwordToEncrypt;
    $result['encryption_key_used'] = $encryptionKeyUsed;
    $result['how_to_use'] = 'Copy the "encrypted_password" value and insert it into the auth-add-in table password field in your database';
    
} catch (Exception $e) {
    $result['error'] = $e->getMessage();
}

echo json_encode($result, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);
?>
