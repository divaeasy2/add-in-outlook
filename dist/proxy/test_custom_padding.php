<?php
// test_custom_padding.php - Test non-standard padding methods

echo "=== TESTING CUSTOM PADDING METHODS ===\n\n";

$key = 'divalto2025';
$key_padded = str_pad($key, 32, "\0");
$iv_zero = str_repeat("\0", 16);
$plaintext = 'sCsKRW8ro0RfG4t;JeIUZ_yg';
$expected_b64 = 'zbO20+c2Z9JOxfvtFhPD9rwvC+ZQD7jhGzNmtHSVZic=';
$expected_hex = bin2hex(base64_decode($expected_b64));

echo "Plaintext: '$plaintext' (" . strlen($plaintext) . " chars)\n";
echo "Expected (hex): $expected_hex\n";
echo "Expected (b64): $expected_b64\n\n";

echo "TESTING DIFFERENT PADDING METHODS:\n";
echo "─────────────────────────────────────────\n\n";

// Test 1: No padding, just pad with null bytes to 32
echo "Test 1: Null-byte padding (0x00)\n";
$padded_plaintext = $plaintext . str_repeat("\0", 32 - strlen($plaintext));
$encrypted = openssl_encrypt($padded_plaintext, 'aes-256-cbc', $key_padded, OPENSSL_RAW_DATA, $iv_zero);
$encrypted_b64 = base64_encode($encrypted);
$encrypted_hex = bin2hex($encrypted);
echo "  Padded length: " . strlen($padded_plaintext) . "\n";
echo "  Result (hex): $encrypted_hex\n";
echo "  Match? " . ($encrypted_hex === $expected_hex ? "✅ YES!!!" : "❌ NO") . "\n\n";

// Test 2: Pad with space characters
echo "Test 2: Space padding (0x20)\n";
$padded_plaintext = $plaintext . str_repeat(" ", 32 - strlen($plaintext));
$encrypted = openssl_encrypt($padded_plaintext, 'aes-256-cbc', $key_padded, OPENSSL_RAW_DATA, $iv_zero);
$encrypted_b64 = base64_encode($encrypted);
$encrypted_hex = bin2hex($encrypted);
echo "  Result (hex): $encrypted_hex\n";
echo "  Match? " . ($encrypted_hex === $expected_hex ? "✅ YES!!!" : "❌ NO") . "\n\n";

// Test 3: No padding at all (just take first 32 bytes)
echo "Test 3: No padding (truncate/pad to 32)\n";
$padded_plaintext = substr(($plaintext . str_repeat("\0", 100)), 0, 32);
$encrypted = openssl_encrypt($padded_plaintext, 'aes-256-cbc', $key_padded, OPENSSL_RAW_DATA, $iv_zero);
$encrypted_b64 = base64_encode($encrypted);
$encrypted_hex = bin2hex($encrypted);
echo "  Result (hex): $encrypted_hex\n";
echo "  Match? " . ($encrypted_hex === $expected_hex ? "✅ YES!!!" : "❌ NO") . "\n\n";

// Test 4: Pad with repetition of padding length (PKCS7 but different)
echo "Test 4: PKCS7-style with length byte\n";
$padding_len = 32 - strlen($plaintext);
$padded_plaintext = $plaintext . str_repeat(chr($padding_len), $padding_len);
$encrypted = openssl_encrypt($padded_plaintext, 'aes-256-cbc', $key_padded, OPENSSL_RAW_DATA, $iv_zero);
$encrypted_b64 = base64_encode($encrypted);
$encrypted_hex = bin2hex($encrypted);
echo "  Padding length: $padding_len\n";
echo "  Padding char: 0x" . dechex($padding_len) . "\n";
echo "  Result (hex): $encrypted_hex\n";
echo "  Match? " . ($encrypted_hex === $expected_hex ? "✅ YES!!!" : "❌ NO") . "\n\n";

// Analyze what the actual padding is in the expected result
echo "ANALYZING EXPECTED PADDING:\n";
echo "─────────────────────────────────────────\n\n";
$expected_ct = base64_decode($expected_b64);
$plaintext_part = substr($expected_ct, 0, 16); // First block is encrypted plaintext
$second_block = substr($expected_ct, 16, 16);

echo "First block (encrypted): " . bin2hex($plaintext_part) . "\n";
echo "Second block (encrypted): " . bin2hex($second_block) . "\n\n";

// Try to decrypt the second block to see what plaintext was encrypted
$second_block_decrypted = openssl_decrypt($second_block, 'aes-256-cbc', $key_padded, OPENSSL_RAW_DATA, substr($second_block, 0, 16)); // wrong but let's try
// Actually, the IV for the second block in CBC is the ciphertext of the first block
$second_block_decrypted = openssl_decrypt($second_block, 'aes-256-cbc', $key_padded, OPENSSL_RAW_DATA, $plaintext_part);
echo "Second block decrypted (with IV=first ciphertext): " . bin2hex($second_block_decrypted) . "\n";
echo "As ASCII: " . $second_block_decrypted . "\n";
echo "As hex bytes: ";
for ($i = 0; $i < strlen($second_block_decrypted); $i++) {
    echo dechex(ord($second_block_decrypted[$i])) . " ";
}
echo "\n\n";

// The second block should be: last 8 bytes of plaintext + 8 bytes of padding
echo "Expected second block plaintext: " . substr($plaintext, 16) . " + [8 bytes padding]\n";
echo "Plaintext hex: " . bin2hex(substr($plaintext, 16)) . "\n";
echo "Expected in second block to encrypt: " . bin2hex(substr($plaintext, 16) . str_repeat(chr(8), 8)) . "\n";

?>
