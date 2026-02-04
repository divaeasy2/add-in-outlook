<?php
// test_decryption.php - Test AES-256-CBC decryption with known values

echo "=== AES-256-CBC Decryption Test ===\n\n";

// Known values from encode-decode.com website
$testKey = 'divaito2025';  // Key shown in screenshot
$encryptedFromWebsite = 'zbO20+c2Z9JOxfvtFhPD9rwvC+ZQD7jhGzNmtHSVZic=';
$expectedPlaintext = 'sCsKRW8ro0RfG4t,JeIUZ_yg';

echo "Test 1: Using key from website (divaito2025)\n";
echo "─────────────────────────────────────────────\n";
echo "Key: $testKey\n";
echo "Encrypted: $encryptedFromWebsite\n";
echo "Expected plaintext: $expectedPlaintext\n\n";

testDecryption($encryptedFromWebsite, $testKey, $expectedPlaintext);

// Also test with divalto2025
$testKey2 = 'divalto2025';  // Key in code
echo "\n\nTest 2: Using key from code (divalto2025)\n";
echo "─────────────────────────────────────────────\n";
echo "Key: $testKey2\n";
echo "Encrypted: $encryptedFromWebsite\n";
echo "Expected plaintext: $expectedPlaintext\n\n";

testDecryption($encryptedFromWebsite, $testKey2, $expectedPlaintext);

function testDecryption($encrypted, $key, $expected) {
    echo "Step 1: EVP_BytesToKey derivation\n";
    $salt = '';
    $hash = '';
    $d = $dI = '';
    while (strlen($hash) < 48) {
        $dI = md5($d . $key . $salt, true);
        $d .= $dI;
        $hash .= $dI;
    }
    $derivedKey = substr($hash, 0, 32);
    $derivedIv = substr($hash, 32, 16);
    
    echo "   Derived Key (hex): " . bin2hex($derivedKey) . "\n";
    echo "   Derived IV (hex): " . bin2hex($derivedIv) . "\n\n";
    
    echo "Step 2: Base64 decode\n";
    $ciphertext = base64_decode($encrypted, true);
    if ($ciphertext === false) {
        echo "   ❌ Base64 decode FAILED\n";
        return;
    }
    echo "   ✅ Base64 decode successful\n";
    echo "   Ciphertext (hex): " . bin2hex($ciphertext) . "\n";
    echo "   Ciphertext length: " . strlen($ciphertext) . " bytes\n\n";
    
    echo "Step 3: Decrypt - Method 1 (EVP_BytesToKey IV)\n";
    $decrypted = openssl_decrypt($ciphertext, 'aes-256-cbc', $derivedKey, OPENSSL_RAW_DATA, $derivedIv);
    if ($decrypted === false) {
        echo "   ❌ Decryption failed: " . openssl_error_string() . "\n";
        
        echo "\nStep 4: Decrypt - Method 2 (Zero IV)\n";
        $zeroIv = str_repeat("\0", 16);
        $decrypted = openssl_decrypt($ciphertext, 'aes-256-cbc', $derivedKey, OPENSSL_RAW_DATA, $zeroIv);
        if ($decrypted === false) {
            echo "   ❌ Decryption failed: " . openssl_error_string() . "\n";
            return;
        }
        echo "   ✅ Decryption succeeded with zero IV\n";
    } else {
        echo "   ✅ Decryption succeeded with EVP_BytesToKey IV\n";
    }
    
    $plaintext = trim($decrypted);
    echo "\nStep 5: Result\n";
    echo "   Decrypted plaintext: " . htmlspecialchars($plaintext) . "\n";
    echo "   Expected plaintext:  " . htmlspecialchars($expected) . "\n";
    
    if ($plaintext === $expected) {
        echo "   ✅ MATCH! Decryption is correct!\n";
    } else {
        echo "   ❌ NO MATCH! Values don't match.\n";
        echo "   Decrypted length: " . strlen($plaintext) . " chars\n";
        echo "   Expected length:  " . strlen($expected) . " chars\n";
    }
}

echo "\n\n=== Summary ===\n";
echo "This test helps identify:\n";
echo "1. Which encryption key is correct (divaito2025 vs divalto2025)\n";
echo "2. Which decryption method works (EVP_BytesToKey IV vs Zero IV)\n";
echo "3. Whether the algorithm matches the website's implementation\n";
?>
