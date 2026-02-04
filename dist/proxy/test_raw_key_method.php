<?php
// test_raw_key_method.php - Test the "raw key padded" method in detail

echo "=== TESTING RAW KEY PADDED METHOD IN DETAIL ===\n\n";

$key = 'divalto2025';
$key_padded = str_pad($key, 32, "\0");
$iv_zero = str_repeat("\0", 16);

$tests = [
    'test' => 'uifBXK7PiJNphwEsNlkdvA==',
    'hello' => '+CSXULqfeewZP1D4RUMyOQ==',
    'a' => 'j10Unc80rVrODPNBU/4sVg==',
    'password123' => 'AxhgWhE5ZrbN9SgLXWLBuA==',
    'sCsKRW8ro0RfG4t;JeIUZ_yg' => 'zbO20+c2Z9JOxfvtFhPD9rwvC+ZQD7jhGzNmtHSVZic=',
];

echo "Key (raw): $key\n";
echo "Key (hex padded): " . bin2hex($key_padded) . "\n";
echo "IV (hex): " . bin2hex($iv_zero) . "\n\n";

foreach ($tests as $plaintext => $expected_b64) {
    echo "Testing: '$plaintext' (" . strlen($plaintext) . " chars)\n";
    
    $encrypted = openssl_encrypt($plaintext, 'aes-256-cbc', $key_padded, OPENSSL_RAW_DATA, $iv_zero);
    $encrypted_b64 = base64_encode($encrypted);
    
    echo "  Expected: $expected_b64\n";
    echo "  Got:      $encrypted_b64\n";
    
    if ($encrypted_b64 === $expected_b64) {
        echo "  ✅ MATCH!\n";
    } else {
        echo "  ❌ NO MATCH\n";
        echo "  Expected (hex): " . bin2hex(base64_decode($expected_b64)) . "\n";
        echo "  Got (hex):      " . bin2hex($encrypted) . "\n";
    }
    echo "\n";
}

echo "\n" . str_repeat("=", 80) . "\n";
echo "ANALYSIS\n";
echo str_repeat("=", 80) . "\n\n";

// Check if maybe it's using a DIFFERENT IV per encryption
echo "Testing with different IVs:\n\n";

// Maybe IV is derived from plaintext?
foreach ($tests as $plaintext => $expected_b64) {
    // Try IV as hash of plaintext
    $iv_hash = substr(hash('md5', $plaintext, true), 0, 16);
    $encrypted = openssl_encrypt($plaintext, 'aes-256-cbc', $key_padded, OPENSSL_RAW_DATA, $iv_hash);
    $encrypted_b64 = base64_encode($encrypted);
    
    if ($encrypted_b64 === $expected_b64) {
        echo "IV from MD5(plaintext) works for: '$plaintext'\n";
    }
}

echo "\n";

// Maybe IV is derived from key?
foreach ($tests as $plaintext => $expected_b64) {
    // Try IV as hash of key
    $iv_hash = substr(hash('md5', $key, true), 0, 16);
    $encrypted = openssl_encrypt($plaintext, 'aes-256-cbc', $key_padded, OPENSSL_RAW_DATA, $iv_hash);
    $encrypted_b64 = base64_encode($encrypted);
    
    if ($encrypted_b64 === $expected_b64) {
        echo "IV from MD5(key) works for: '$plaintext'\n";
    }
}

echo "\n";

// Try different padding modes
echo "Testing with different padding:\n\n";

// Maybe it's using OPENSSL_RAW_DATA vs automatic padding
foreach ($tests as $plaintext => $expected_b64) {
    // Try with automatic PKCS7 padding (openssl_encrypt default)
    $encrypted = openssl_encrypt($plaintext, 'aes-256-cbc', $key_padded, false, $iv_zero);
    $encrypted_b64 = base64_encode($encrypted);
    
    if ($encrypted_b64 === $expected_b64) {
        echo "✅ Auto padding works for: '$plaintext'\n";
    }
}

?>
