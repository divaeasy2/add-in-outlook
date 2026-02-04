<?php
// diagnostic_encryption.php - Analyze encryption mismatch

echo "=== ENCRYPTION DIAGNOSTIC ===\n\n";

// Values from the website screenshot
$website_encrypted = 'zbO20+c2Z9JOxfvtFhPD9rwvC+ZQD7jhGzNmtHSVZic=';
$website_plaintext = 'sCsKRW8ro0RfG4t;JeIUZ_yg';
$website_key = 'divaito2025';

echo "1. WEBSITE VALUES (from screenshot):\n";
echo "   Plaintext: " . $website_plaintext . "\n";
echo "   Plaintext length: " . strlen($website_plaintext) . " chars\n";
echo "   Encrypted: " . $website_encrypted . "\n";
echo "   Key: " . $website_key . "\n\n";

// Decode the ciphertext to see what we have
$ciphertext = base64_decode($website_encrypted);
echo "2. CIPHERTEXT ANALYSIS:\n";
echo "   Ciphertext (hex): " . bin2hex($ciphertext) . "\n";
echo "   Ciphertext length: " . strlen($ciphertext) . " bytes\n";
echo "   Expected with padding: ";
$plaintext_len = strlen($website_plaintext);
$blocks_needed = ceil($plaintext_len / 16);
$expected_padded_len = $blocks_needed * 16;
echo $expected_padded_len . " bytes (PKCS7)\n\n";

// Try to reverse-engineer the key/IV by attempting different approaches
echo "3. KEY DERIVATION TESTING:\n";

// Test 1: Standard EVP_BytesToKey (no salt)
echo "\n   Test A: EVP_BytesToKey (no salt, MD5)\n";
$key_a = deriveKey($website_key, '');
echo "      Key (hex): " . bin2hex($key_a['key']) . "\n";
echo "      IV (hex): " . bin2hex($key_a['iv']) . "\n";
$result = tryDecrypt($ciphertext, $key_a['key'], $key_a['iv']);
echo "      Result: " . ($result ? "✅ SUCCESS: " . $result : "❌ FAILED") . "\n";

// Test 2: What if the key itself is the hash?
echo "\n   Test B: SHA256 of key as encryption key\n";
$key_b = hash('sha256', $website_key, true); // 32 bytes
$iv_b = hash('md5', $website_key, true); // 16 bytes
echo "      Key (hex): " . bin2hex($key_b) . "\n";
echo "      IV (hex): " . bin2hex($iv_b) . "\n";
$result = tryDecrypt($ciphertext, $key_b, $iv_b);
echo "      Result: " . ($result ? "✅ SUCCESS: " . $result : "❌ FAILED") . "\n";

// Test 3: Just the raw key padded to 32 bytes
echo "\n   Test C: Raw key padded to 32 bytes\n";
$key_c = str_pad($website_key, 32, "\0");
$iv_c = str_repeat("\0", 16);
echo "      Key (hex): " . bin2hex($key_c) . "\n";
echo "      IV (hex): " . bin2hex($iv_c) . "\n";
$result = tryDecrypt($ciphertext, $key_c, $iv_c);
echo "      Result: " . ($result ? "✅ SUCCESS: " . $result : "❌ FAILED") . "\n";

// Test 4: What if we need to encrypt to understand?
echo "\n4. REVERSE ENGINEERING - Try to encrypt with different methods:\n";
echo "   Plaintext to encrypt: " . $website_plaintext . "\n\n";

echo "   Method A: EVP_BytesToKey encryption\n";
$encrypted_a = encryptWithEVPBytesToKey($website_plaintext, $website_key);
echo "      Encrypted (base64): " . $encrypted_a . "\n";
echo "      Match website? " . ($encrypted_a === $website_encrypted ? "✅ YES!" : "❌ NO") . "\n";

echo "\n   Method B: SHA256 key encryption\n";
$encrypted_b = encryptWithSHA256($website_plaintext, $website_key);
echo "      Encrypted (base64): " . $encrypted_b . "\n";
echo "      Match website? " . ($encrypted_b === $website_encrypted ? "✅ YES!" : "❌ NO") . "\n";

echo "\n5. NEXT STEPS:\n";
echo "   - Check database: What's the actual encrypted password stored?\n";
echo "   - Try decrypting database value with website plaintext to reverse-engineer key\n";
echo "   - Ask admin: Exactly what password was entered on the website before encryption?\n";

function deriveKey($key, $salt) {
    $hash = '';
    $d = '';
    while (strlen($hash) < 48) {
        $d = md5($d . $key . $salt, true);
        $hash .= $d;
    }
    return [
        'key' => substr($hash, 0, 32),
        'iv' => substr($hash, 32, 16)
    ];
}

function tryDecrypt($ciphertext, $key, $iv) {
    $result = @openssl_decrypt($ciphertext, 'aes-256-cbc', $key, OPENSSL_RAW_DATA, $iv);
    if ($result === false) {
        return false;
    }
    return trim($result);
}

function encryptWithEVPBytesToKey($plaintext, $key) {
    $salt = '';
    $hash = '';
    $d = '';
    while (strlen($hash) < 48) {
        $d = md5($d . $key . $salt, true);
        $hash .= $d;
    }
    $derived_key = substr($hash, 0, 32);
    $iv = substr($hash, 32, 16);
    
    $encrypted = openssl_encrypt($plaintext, 'aes-256-cbc', $derived_key, OPENSSL_RAW_DATA, $iv);
    return base64_encode($encrypted);
}

function encryptWithSHA256($plaintext, $key) {
    $derived_key = hash('sha256', $key, true);
    $iv = hash('md5', $key, true);
    
    $encrypted = openssl_encrypt($plaintext, 'aes-256-cbc', $derived_key, OPENSSL_RAW_DATA, $iv);
    return base64_encode($encrypted);
}
?>
