<?php
// analyze_test_results.php - Analyze the test encryptions to find the key derivation

echo "=== ANALYZING TEST ENCRYPTIONS ===\n\n";

$key = 'divalto2025';

// Test cases from admin
$tests = [
    'test' => 'uifBXK7PiJNphwEsNlkdvA==',
    'hello' => '+CSXULqfeewZP1D4RUMyOQ==',
    'a' => 'j10Unc80rVrODPNBU/4sVg==',
    'password123' => 'AxhgWhE5ZrbN9SgLXWLBuA==',
    'sCsKRW8ro0RfG4t;JeIUZ_yg' => 'zbO20+c2Z9JOxfvtFhPD9rwvC+ZQD7jhGzNmtHSVZic=',
];

echo "Analyzing ciphertext patterns...\n";
echo "Key: $key\n\n";

foreach ($tests as $plaintext => $ciphertext_b64) {
    $ct = base64_decode($ciphertext_b64);
    echo "Plaintext: '$plaintext' (" . strlen($plaintext) . " chars)\n";
    echo "  Ciphertext: $ciphertext_b64\n";
    echo "  Ciphertext (hex): " . bin2hex($ct) . "\n";
    echo "  Ciphertext length: " . strlen($ct) . " bytes\n";
    echo "  Expected with PKCS7: " . (ceil((strlen($plaintext)) / 16) * 16) . " bytes\n\n";
}

echo "\n" . str_repeat("=", 80) . "\n";
echo "KEY DERIVATION SEARCH\n";
echo str_repeat("=", 80) . "\n\n";

// Try different key derivations
$key_derivations = [
    'raw_key_zero_iv' => function($k) {
        return [
            'key' => str_pad($k, 32, "\0"),
            'iv' => str_repeat("\0", 16),
            'name' => 'Raw key padded with zeros, zero IV'
        ];
    },
    'md5_key_md5_iv' => function($k) {
        return [
            'key' => str_pad(hash('md5', $k, true) . hash('md5', $k . 'extra', true), 32),
            'iv' => hash('md5', $k, true),
            'name' => 'MD5 hash as key, MD5 hash as IV'
        ];
    },
    'sha256_key_zero_iv' => function($k) {
        return [
            'key' => hash('sha256', $k, true),
            'iv' => str_repeat("\0", 16),
            'name' => 'SHA256 hash as key, zero IV'
        ];
    },
    'sha1_derived' => function($k) {
        $h = hash('sha1', $k, true);
        return [
            'key' => $h . substr(hash('sha1', $h . $k, true), 0, 12),
            'iv' => str_repeat("\0", 16),
            'name' => 'SHA1 derived, zero IV'
        ];
    },
    'sha256_both' => function($k) {
        return [
            'key' => hash('sha256', $k, true),
            'iv' => substr(hash('sha256', $k . 'iv', true), 0, 16),
            'name' => 'SHA256 for key, SHA256 for IV'
        ];
    },
    'evp_bytestokey' => function($k) {
        $salt = '';
        $hash = '';
        $d = '';
        while (strlen($hash) < 48) {
            $d = md5($d . $k . $salt, true);
            $hash .= $d;
        }
        return [
            'key' => substr($hash, 0, 32),
            'iv' => substr($hash, 32, 16),
            'name' => 'EVP_BytesToKey (MD5, no salt)'
        ];
    },
];

$found = false;

foreach ($key_derivations as $method_name => $derivation_fn) {
    $params = $derivation_fn($key);
    $key_derived = $params['key'];
    $iv_derived = $params['iv'];
    $name = $params['name'];
    
    echo "Testing: $name\n";
    echo "  Key (hex): " . bin2hex($key_derived) . "\n";
    echo "  IV (hex): " . bin2hex($iv_derived) . "\n";
    
    $matches = 0;
    $total = count($tests);
    
    foreach ($tests as $plaintext => $expected_b64) {
        $encrypted = openssl_encrypt($plaintext, 'aes-256-cbc', $key_derived, OPENSSL_RAW_DATA, $iv_derived);
        $encrypted_b64 = base64_encode($encrypted);
        
        if ($encrypted_b64 === $expected_b64) {
            $matches++;
        }
    }
    
    echo "  Result: $matches/$total matches\n";
    if ($matches === $total) {
        echo "  ✅✅✅ FOUND IT! THIS IS THE CORRECT KEY DERIVATION! ✅✅✅\n";
        $found = true;
    }
    echo "\n";
}

if (!$found) {
    echo "\n⚠️  No exact match found. Trying more variations...\n\n";
    
    // Try PBKDF2 with different configurations
    echo "Trying PBKDF2 variations:\n\n";
    
    $pbkdf2_configs = [
        ['sha1', 1000],
        ['sha1', 10000],
        ['sha256', 1000],
        ['sha256', 10000],
        ['sha512', 1000],
    ];
    
    foreach ($pbkdf2_configs as [$algo, $iterations]) {
        $key_derived = hash_pbkdf2($algo, $key, '', $iterations, 32, true);
        $iv_derived = str_repeat("\0", 16);
        
        echo "PBKDF2-$algo ($iterations iterations):\n";
        echo "  Key (hex): " . bin2hex($key_derived) . "\n";
        
        $matches = 0;
        foreach ($tests as $plaintext => $expected_b64) {
            $encrypted = openssl_encrypt($plaintext, 'aes-256-cbc', $key_derived, OPENSSL_RAW_DATA, $iv_derived);
            $encrypted_b64 = base64_encode($encrypted);
            if ($encrypted_b64 === $expected_b64) {
                $matches++;
            }
        }
        
        echo "  Result: $matches/" . count($tests) . " matches\n";
        if ($matches === count($tests)) {
            echo "  ✅✅✅ FOUND IT! ✅✅✅\n";
            $found = true;
        }
        echo "\n";
    }
}

if ($found) {
    echo "\n🎉 SUCCESS! Update decryptPassword() function with the found derivation!\n";
} else {
    echo "\n⚠️  Still searching. May need custom JavaScript/Node.js investigation.\n";
}
?>
