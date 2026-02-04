<?php
// final_padding_test.php - Test different padding scenarios

echo "=== FINAL PADDING TEST ===\n\n";

$key = 'divalto2025';
$key_padded = str_pad($key, 32, "\0");
$iv_zero = str_repeat("\0", 16);
$plaintext = 'sCsKRW8ro0RfG4t;JeIUZ_yg';
$expected_b64 = 'zbO20+c2Z9JOxfvtFhPD9rwvC+ZQD7jhGzNmtHSVZic=';

echo "Plaintext: '$plaintext' (" . strlen($plaintext) . " chars)\n";
echo "Expected: $expected_b64\n";
echo "Expected (hex): " . bin2hex(base64_decode($expected_b64)) . "\n\n";

echo "Test 1: With OPENSSL_RAW_DATA (no automatic padding)\n";
$encrypted = openssl_encrypt($plaintext, 'aes-256-cbc', $key_padded, OPENSSL_RAW_DATA, $iv_zero);
$encrypted_b64 = base64_encode($encrypted);
echo "  Result: $encrypted_b64\n";
echo "  Result (hex): " . bin2hex($encrypted) . "\n";
echo "  Match? " . ($encrypted_b64 === $expected_b64 ? "✅ YES" : "❌ NO") . "\n\n";

echo "Test 2: Without OPENSSL_RAW_DATA (auto PKCS7 padding)\n";
$encrypted = openssl_encrypt($plaintext, 'aes-256-cbc', $key_padded, false, $iv_zero);
$encrypted_b64 = base64_encode($encrypted);
echo "  Result: $encrypted_b64\n";
echo "  Result (hex): " . bin2hex($encrypted) . "\n";
echo "  Match? " . ($encrypted_b64 === $expected_b64 ? "✅ YES" : "❌ NO") . "\n\n";

echo "Test 3: Manual PKCS7 padding + OPENSSL_RAW_DATA\n";
$padding_len = 16 - (strlen($plaintext) % 16);
$padded_plaintext = $plaintext . str_repeat(chr($padding_len), $padding_len);
echo "  Plaintext length: " . strlen($plaintext) . "\n";
echo "  Padding length: $padding_len\n";
echo "  Padded plaintext length: " . strlen($padded_plaintext) . "\n";
$encrypted = openssl_encrypt($padded_plaintext, 'aes-256-cbc', $key_padded, OPENSSL_RAW_DATA, $iv_zero);
$encrypted_b64 = base64_encode($encrypted);
echo "  Result: $encrypted_b64\n";
echo "  Result (hex): " . bin2hex($encrypted) . "\n";
echo "  Match? " . ($encrypted_b64 === $expected_b64 ? "✅ YES" : "❌ NO") . "\n\n";

echo "Test 4: Check if the first block is consistent\n";
$ct_expected = base64_decode($expected_b64);
$ct_got = base64_decode($encrypted_b64);
$first_block_expected = substr($ct_expected, 0, 16);
$first_block_got = substr($ct_got, 0, 16);
echo "  First 16 bytes expected (hex): " . bin2hex($first_block_expected) . "\n";
echo "  First 16 bytes got (hex):      " . bin2hex($first_block_got) . "\n";
echo "  Match? " . ($first_block_expected === $first_block_got ? "✅ YES" : "❌ NO") . "\n\n";

if ($first_block_expected === $first_block_got) {
    echo "✅ First block matches - key and IV are correct!\n";
    echo "❌ Padding block differs - this is a padding issue or the admin's tool uses different padding\n\n";
    
    // Check what padding the expected result has
    echo "Analysis of padding in expected result:\n";
    $last_byte_expected = ord(substr($ct_expected, -1));
    echo "  Last byte value: " . $last_byte_expected . " (0x" . dechex($last_byte_expected) . ")\n";
    echo "  If PKCS7, should have " . $last_byte_expected . " bytes of 0x" . dechex($last_byte_expected) . "\n";
    
    // Extract what appears to be the padding
    $suspected_padding = substr($ct_expected, -$last_byte_expected);
    $is_pkcs7 = strlen(array_unique(str_split($suspected_padding))) === 1;
    echo "  Appears to be PKCS7? " . ($is_pkcs7 ? "✅ YES" : "❌ NO") . "\n";
}

?>
