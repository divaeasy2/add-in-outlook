<?php
/**
 * Password Encryption/Decryption Tool
 * 
 * Admin uses this to encrypt database passwords
 * Application uses the same method to decrypt them
 * 
 * Standard: AES-256-CBC with documented key derivation
 */

// Configuration
define('ENCRYPTION_KEY', 'divalto2025');
define('ENCRYPTION_METHOD', 'aes-256-cbc');

// Mode: 'encrypt' or 'decrypt'
$mode = $_GET['mode'] ?? 'encrypt';
$password = $_POST['password'] ?? $_GET['password'] ?? '';
$result = '';
$error = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST' || !empty($password)) {
    try {
        if ($mode === 'encrypt') {
            $result = encryptPassword($password);
        } elseif ($mode === 'decrypt') {
            $result = decryptPassword($password);
        } else {
            $error = 'Invalid mode. Use "encrypt" or "decrypt"';
        }
    } catch (Exception $e) {
        $error = 'Error: ' . $e->getMessage();
    }
}

/**
 * Encrypt password using AES-256-CBC
 * Standard method suitable for database storage
 */
function encryptPassword($plaintext) {
    // Prepare key: pad to 32 bytes with null bytes
    $key = str_pad(ENCRYPTION_KEY, 32, "\0");
    
    // IV: 16 zero bytes (deterministic for consistency)
    $iv = str_repeat("\0", 16);
    
    // Encrypt with automatic PKCS7 padding
    $encrypted = openssl_encrypt(
        $plaintext,
        ENCRYPTION_METHOD,
        $key,
        0,  // 0 = with automatic PKCS7 padding, 1 = raw data
        $iv
    );
    
    if ($encrypted === false) {
        throw new Exception('Encryption failed: ' . openssl_error_string());
    }
    
    return $encrypted;
}

/**
 * Decrypt password using AES-256-CBC
 * Application uses this to retrieve original password
 */
function decryptPassword($ciphertext) {
    // Prepare key: same as encryption
    $key = str_pad(ENCRYPTION_KEY, 32, "\0");
    
    // IV: same as encryption
    $iv = str_repeat("\0", 16);
    
    // Decrypt with automatic PKCS7 padding
    $decrypted = openssl_decrypt(
        $ciphertext,
        ENCRYPTION_METHOD,
        $key,
        0,  // 0 = with automatic PKCS7 padding, 1 = raw data
        $iv
    );
    
    if ($decrypted === false) {
        throw new Exception('Decryption failed: ' . openssl_error_string());
    }
    
    return $decrypted;
}

?>
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Password Encryption Tool</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
        }
        
        .container {
            background: white;
            border-radius: 10px;
            box-shadow: 0 20px 60px rgba(0, 0, 0, 0.3);
            max-width: 600px;
            width: 100%;
            padding: 40px;
        }
        
        h1 {
            color: #333;
            margin-bottom: 10px;
            text-align: center;
        }
        
        .subtitle {
            color: #666;
            text-align: center;
            margin-bottom: 30px;
            font-size: 14px;
        }
        
        .tabs {
            display: flex;
            gap: 10px;
            margin-bottom: 30px;
            border-bottom: 2px solid #eee;
        }
        
        .tab-btn {
            padding: 12px 20px;
            border: none;
            background: none;
            cursor: pointer;
            font-size: 16px;
            color: #666;
            border-bottom: 3px solid transparent;
            transition: all 0.3s ease;
        }
        
        .tab-btn.active {
            color: #667eea;
            border-bottom-color: #667eea;
        }
        
        .tab-btn:hover {
            color: #667eea;
        }
        
        .form-group {
            margin-bottom: 20px;
        }
        
        label {
            display: block;
            margin-bottom: 8px;
            color: #333;
            font-weight: 600;
            font-size: 14px;
        }
        
        textarea, input[type="text"], input[type="password"] {
            width: 100%;
            padding: 12px;
            border: 2px solid #eee;
            border-radius: 5px;
            font-family: 'Monaco', 'Courier New', monospace;
            font-size: 14px;
            transition: border-color 0.3s ease;
        }
        
        textarea:focus, input:focus {
            outline: none;
            border-color: #667eea;
        }
        
        textarea {
            min-height: 100px;
            resize: vertical;
        }
        
        .button-group {
            display: flex;
            gap: 10px;
            margin-top: 25px;
        }
        
        button {
            flex: 1;
            padding: 12px 20px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
            font-weight: 600;
            cursor: pointer;
            transition: all 0.3s ease;
        }
        
        .btn-primary {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            color: white;
        }
        
        .btn-primary:hover {
            transform: translateY(-2px);
            box-shadow: 0 10px 20px rgba(102, 126, 234, 0.4);
        }
        
        .btn-secondary {
            background: #f0f0f0;
            color: #333;
        }
        
        .btn-secondary:hover {
            background: #e0e0e0;
        }
        
        .result-box {
            background: #f8f9fa;
            border: 2px solid #eee;
            border-radius: 5px;
            padding: 15px;
            margin-top: 20px;
            display: none;
        }
        
        .result-box.show {
            display: block;
        }
        
        .result-box.success {
            border-color: #4caf50;
            background: #f1f8f4;
        }
        
        .result-box.error {
            border-color: #f44336;
            background: #fdeaea;
        }
        
        .result-label {
            font-size: 12px;
            color: #666;
            text-transform: uppercase;
            letter-spacing: 1px;
            margin-bottom: 8px;
        }
        
        .result-text {
            word-break: break-all;
            font-family: 'Monaco', 'Courier New', monospace;
            padding: 10px;
            background: white;
            border-radius: 3px;
            font-size: 13px;
            color: #333;
        }
        
        .error-text {
            color: #f44336;
            font-weight: 600;
        }
        
        .copy-btn {
            margin-top: 10px;
            padding: 8px 15px;
            font-size: 13px;
            background: #667eea;
            color: white;
            border: none;
            border-radius: 3px;
            cursor: pointer;
        }
        
        .copy-btn:hover {
            background: #5568d3;
        }
        
        .info-box {
            background: #e3f2fd;
            border-left: 4px solid #2196f3;
            padding: 15px;
            margin-top: 20px;
            border-radius: 3px;
            font-size: 13px;
            color: #1976d2;
        }
        
        .info-box strong {
            display: block;
            margin-bottom: 5px;
        }
        
        .hidden {
            display: none;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>🔐 Password Encryption Tool</h1>
        <p class="subtitle">Encrypt/Decrypt passwords for database storage</p>
        
        <div class="tabs">
            <button class="tab-btn active" onclick="switchMode('encrypt')">🔒 Encrypt</button>
            <button class="tab-btn" onclick="switchMode('decrypt')">🔓 Decrypt</button>
        </div>
        
        <form method="POST" id="encryptForm">
            <div class="form-group">
                <label for="password">Password to Encrypt:</label>
                <textarea id="password" name="password" placeholder="Enter the password to encrypt..." required></textarea>
            </div>
            
            <div class="button-group">
                <button type="submit" class="btn-primary">Encrypt</button>
                <button type="reset" class="btn-secondary">Clear</button>
            </div>
        </form>
        
        <form method="POST" id="decryptForm" class="hidden">
            <div class="form-group">
                <label for="ciphertext">Encrypted Password (Base64):</label>
                <textarea id="ciphertext" name="password" placeholder="Paste the encrypted password here..." required></textarea>
            </div>
            
            <div class="button-group">
                <button type="submit" class="btn-primary">Decrypt</button>
                <button type="reset" class="btn-secondary">Clear</button>
            </div>
        </form>
        
        <?php if ($result || $error): ?>
        <div class="result-box <?php echo $error ? 'error show' : 'success show'; ?>">
            <div class="result-label">
                <?php echo $error ? '❌ Error' : '✅ Result'; ?>
            </div>
            <?php if ($error): ?>
                <div class="error-text"><?php echo htmlspecialchars($error); ?></div>
            <?php else: ?>
                <div class="result-text"><?php echo htmlspecialchars($result); ?></div>
                <button class="copy-btn" onclick="copyToClipboard('<?php echo addslashes($result); ?>')">📋 Copy to Clipboard</button>
            <?php endif; ?>
        </div>
        <?php endif; ?>
        
        <div class="info-box">
            <strong>ℹ️ How to Use:</strong>
            <div>
                <strong>For Admin:</strong> Use the "Encrypt" tab to encrypt database passwords. Copy the result and store it in the database.
            </div>
            <div style="margin-top: 10px;">
                <strong>For App:</strong> The application automatically decrypts passwords using the same method.
            </div>
            <div style="margin-top: 10px;">
                <strong>Method:</strong> AES-256-CBC with standard PKCS7 padding
            </div>
        </div>
    </div>
    
    <script>
        function switchMode(mode) {
            const encryptForm = document.getElementById('encryptForm');
            const decryptForm = document.getElementById('decryptForm');
            const tabs = document.querySelectorAll('.tab-btn');
            
            tabs.forEach(btn => btn.classList.remove('active'));
            
            if (mode === 'encrypt') {
                encryptForm.classList.remove('hidden');
                decryptForm.classList.add('hidden');
                tabs[0].classList.add('active');
                document.location.hash = '#encrypt';
            } else {
                encryptForm.classList.add('hidden');
                decryptForm.classList.remove('hidden');
                tabs[1].classList.add('active');
                document.location.hash = '#decrypt';
            }
        }
        
        function copyToClipboard(text) {
            navigator.clipboard.writeText(text).then(() => {
                alert('Copied to clipboard!');
            }).catch(() => {
                alert('Failed to copy');
            });
        }
        
        // Handle form submissions
        document.getElementById('encryptForm').addEventListener('submit', function(e) {
            if (!document.getElementById('password').value) {
                e.preventDefault();
                alert('Please enter a password to encrypt');
            }
        });
        
        document.getElementById('decryptForm').addEventListener('submit', function(e) {
            if (!document.getElementById('ciphertext').value) {
                e.preventDefault();
                alert('Please enter encrypted password');
            }
        });
    </script>
</body>
</html>
