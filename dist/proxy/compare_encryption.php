<?php
// compare_encryption.php - Compare database vs website encryption

echo "=== ENCRYPTION COMPARISON ===\n\n";

// From the website screenshot
$website_key = 'divalto2025';  // Confirmed from your logs
$website_encrypted = 'zbO20+c2Z9JOxfvtFhPD9rwvC+ZQD7jhGzNmtHSVZic=';
$website_plaintext = 'sCsKRW8ro0RfG4t;JeIUZ_yg';

echo "STEP 1: Understand the website encryption\n";
echo "─────────────────────────────────────────\n";
echo "Website used:\n";
echo "  Key: $website_key\n";
echo "  Plaintext: $website_plaintext\n";
echo "  Encrypted: $website_encrypted\n\n";

// The critical question: Did the admin encrypt the SAME password that's in the database?
echo "STEP 2: Critical Question\n";
echo "─────────────────────────────────────────\n";
echo "⚠️  The real issue is likely:\n";
echo "   The password shown on the website screenshot may NOT be\n";
echo "   the same password that was encrypted and stored in the database.\n\n";

echo "STEP 3: What you need to check\n";
echo "─────────────────────────────────────────\n";
echo "1. Check your database:\n";
echo "   SELECT password FROM 'auth-add-in' LIMIT 1;\n";
echo "   → Copy the exact encrypted value\n\n";

echo "2. Compare with database:\n";
echo "   Is the database password = $website_encrypted?\n";
echo "   If NO: The admin encrypted a different password\n";
echo "   If YES: Then we have a key/algorithm problem\n\n";

echo "3. Ask the admin:\n";
echo "   - What password did you encrypt on the website?\n";
echo "   - Is it the same as what's in the database?\n";
echo "   - Can you encrypt 'test' on the website and show me the result?\n\n";

echo "STEP 4: Testing approach\n";
echo "─────────────────────────────────────────\n";
echo "Run this PHP file at: /src/proxy/compare_encryption.php\n";
echo "Then check: /src/proxy/diagnostic_encryption.php\n";
echo "And check: /src/proxy/test_decryption.php\n\n";

echo "STEP 5: If database has DIFFERENT encrypted value\n";
echo "─────────────────────────────────────────\n";
echo "We need to:\n";
echo "1. Know what plaintext password was stored\n";
echo "2. Try different key combinations\n";
echo "3. Check if the admin used a different tool/method\n\n";

// Create a test encryption with the website key to show what SHOULD happen
echo "TEST: Encrypting plaintext with website key\n";
echo "─────────────────────────────────────────\n";

$test_plaintext = "sCsKRW8ro0RfG4t;Je!UZ_yg";
$test_key = "divalto2025";

// Standard EVP_BytesToKey derivation
$salt = '';
$hash = '';
$d = '';
while (strlen($hash) < 48) {
    $d = md5($d . $test_key . $salt, true);
    $hash .= $d;
}
$key = substr($hash, 0, 32);
$iv = substr($hash, 32, 16);

$encrypted = openssl_encrypt($test_plaintext, 'aes-256-cbc', $key, OPENSSL_RAW_DATA, $iv);
$encrypted_b64 = base64_encode($encrypted);

echo "Encrypting: $test_plaintext\n";
echo "Key: $test_key\n";
echo "Derived Key (hex): " . bin2hex($key) . "\n";
echo "Derived IV (hex): " . bin2hex($iv) . "\n";
echo "Result: $encrypted_b64\n";
echo "Matches website? " . ($encrypted_b64 === $website_encrypted ? "✅ YES" : "❌ NO") . "\n";

if ($encrypted_b64 !== $website_encrypted) {
    echo "\n⚠️  MISMATCH! The encrypted value doesn't match what we expected.\n";
    echo "This means one of:\n";
    echo "  1. The key is wrong\n";
    echo "  2. The algorithm is different than EVP_BytesToKey\n";
    echo "  3. The plaintext is wrong\n";
}
?>
