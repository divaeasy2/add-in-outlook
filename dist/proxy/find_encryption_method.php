<?php
// find_encryption_method.php - Exhaustive search for the encryption method

echo "=== FINDING THE ENCRYPTION METHOD ===\n\n";

$plaintext = 'sCsKRW8ro0RfG4t;JeIUZ_yg';
$key = 'divalto2025';
$target_encrypted = 'zbO20+c2Z9JOxfvtFhPD9rwvC+ZQD7jhGzNmtHSVZic=';

echo "Searching for encryption method...\n";
echo "Plaintext: $plaintext\n";
echo "Key: $key\n";
echo "Target (what we need to match): $target_encrypted\n\n";

// Test 1: Different key derivations with AES-256-CBC
echo "TEST 1: Different Key Derivation Methods (AES-256-CBC)\n";
echo "─────────────────────────────────────────────────────\n\n";

// 1A: PBKDF2 with SHA1
$derivedkey_1a = hash_pbkdf2('sha1', $key, '', 1000, 32, true);
$iv_1a = str_repeat("\0", 16);
$enc_1a = openssl_encrypt($plaintext, 'aes-256-cbc', $derivedkey_1a, OPENSSL_RAW_DATA, $iv_1a);
$b64_1a = base64_encode($enc_1a);
echo "1A. PBKDF2-SHA1 (1000 iterations): $b64_1a\n";
echo "    Match? " . ($b64_1a === $target_encrypted ? "✅ YES!!!" : "❌ NO") . "\n\n";

// 1B: PBKDF2 with SHA256
$derivedkey_1b = hash_pbkdf2('sha256', $key, '', 1000, 32, true);
$iv_1b = str_repeat("\0", 16);
$enc_1b = openssl_encrypt($plaintext, 'aes-256-cbc', $derivedkey_1b, OPENSSL_RAW_DATA, $iv_1b);
$b64_1b = base64_encode($enc_1b);
echo "1B. PBKDF2-SHA256 (1000 iterations): $b64_1b\n";
echo "    Match? " . ($b64_1b === $target_encrypted ? "✅ YES!!!" : "❌ NO") . "\n\n";

// 1C: Simple MD5 hash
$derivedkey_1c = hash('md5', $key, true);
$derivedkey_1c = $derivedkey_1c . hash('md5', $derivedkey_1c . $key, true); // Pad to 32 bytes
$iv_1c = hash('md5', $key . '1', true);
$enc_1c = openssl_encrypt($plaintext, 'aes-256-cbc', $derivedkey_1c, OPENSSL_RAW_DATA, $iv_1c);
$b64_1c = base64_encode($enc_1c);
echo "1C. MD5 double hash (custom): $b64_1c\n";
echo "    Match? " . ($b64_1c === $target_encrypted ? "✅ YES!!!" : "❌ NO") . "\n\n";

// 1D: SHA256 hash
$derivedkey_1d = hash('sha256', $key, true);
$iv_1d = substr(hash('sha256', $key . 'iv', true), 0, 16);
$enc_1d = openssl_encrypt($plaintext, 'aes-256-cbc', $derivedkey_1d, OPENSSL_RAW_DATA, $iv_1d);
$b64_1d = base64_encode($enc_1d);
echo "1D. SHA256 direct: $b64_1d\n";
echo "    Match? " . ($b64_1d === $target_encrypted ? "✅ YES!!!" : "❌ NO") . "\n\n";

// Test 2: Different cipher modes
echo "TEST 2: Different Cipher Modes\n";
echo "─────────────────────────────────────────────────────\n\n";

$key_simple = str_pad($key, 32, "\0");
$iv_simple = str_repeat("\0", 16);

$ciphers = ['aes-128-cbc', 'aes-192-cbc', 'aes-256-cbc', 'aes-256-ecb', 'aes-256-ctr'];
foreach ($ciphers as $cipher) {
    $enc_test = @openssl_encrypt($plaintext, $cipher, $key_simple, OPENSSL_RAW_DATA, $iv_simple);
    if ($enc_test !== false) {
        $b64_test = base64_encode($enc_test);
        echo "Mode: $cipher: $b64_test\n";
        echo "  Match? " . ($b64_test === $target_encrypted ? "✅ YES!!!" : "❌ NO") . "\n";
    }
}
echo "\n";

// Test 3: Maybe the website is using JavaScript/Node.js crypto
echo "TEST 3: JavaScript/Node.js Crypto Variations\n";
echo "─────────────────────────────────────────────────────\n\n";

// Variation: Key as UTF-8 bytes + different iterations
$derivedkey_3a = hash_pbkdf2('sha256', $key, '', 10000, 32, true);
$iv_3a = str_repeat("\0", 16);
$enc_3a = openssl_encrypt($plaintext, 'aes-256-cbc', $derivedkey_3a, OPENSSL_RAW_DATA, $iv_3a);
$b64_3a = base64_encode($enc_3a);
echo "3A. PBKDF2-SHA256 (10000 iterations): $b64_3a\n";
echo "    Match? " . ($b64_3a === $target_encrypted ? "✅ YES!!!" : "❌ NO") . "\n\n";

// Test 4: Check if maybe no padding is used
echo "TEST 4: Without PKCS7 Padding\n";
echo "─────────────────────────────────────────────────────\n\n";

// Pad plaintext to exactly 32 bytes (2 blocks)
$plaintext_padded = str_pad($plaintext, 32, "\0");
$enc_4a = openssl_encrypt($plaintext_padded, 'aes-256-cbc', $key_simple, OPENSSL_RAW_DATA, $iv_simple);
$b64_4a = base64_encode($enc_4a);
echo "4A. With null padding: $b64_4a\n";
echo "    Match? " . ($b64_4a === $target_encrypted ? "✅ YES!!!" : "❌ NO") . "\n\n";

echo "\n\n=== IMPORTANT ===\n";
echo "If none of these match, we need to:\n";
echo "1. Ask the admin what EXACT tool they used (URL)\n";
echo "2. Ask them to test encrypt something simple like 'test' or 'hello'\n";
echo "3. With test cases, we can reverse-engineer the exact algorithm\n\n";

echo "Check /src/proxy/test_plaintext_variations.php for plaintext variant tests\n";
?>
