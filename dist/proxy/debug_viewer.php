<?php
// debug_viewer.php - View proxy debug logs via web

header("Content-Type: text/html; charset=utf-8");

// Get parameters
$logType = $_GET['log'] ?? 'proxy';  // 'proxy' or 'proxy_child'
$filter = $_GET['filter'] ?? 'decrypt'; // 'decrypt', 'all'
$lines_count = intval($_GET['lines'] ?? 100);

// Determine log file
$logFile = ($logType === 'proxy_child') 
    ? __DIR__ . '/proxy_child_debug.log'
    : __DIR__ . '/proxy_debug.log';

if (!file_exists($logFile)) {
    http_response_code(404);
    echo "<!DOCTYPE html><html><head><meta charset='utf-8'><style>body{font-family:monospace;padding:20px;}</style></head><body>";
    echo "<h2>❌ Debug log file not found</h2>";
    echo "<p>File: $logFile</p>";
    echo "<p><a href='?log=proxy&filter=decrypt'>View Proxy Logs</a> | <a href='?log=proxy_child&filter=decrypt'>View Proxy Child Logs</a></p>";
    echo "</body></html>";
    exit;
}

if (!is_readable($logFile)) {
    http_response_code(403);
    echo "<!DOCTYPE html><html><head><meta charset='utf-8'><style>body{font-family:monospace;padding:20px;}</style></head><body>";
    echo "<h2>❌ Debug log file is not readable</h2>";
    echo "<p>Check file permissions for: $logFile</p>";
    echo "</body></html>";
    exit;
}

$fileSize = filesize($logFile);
$allLines = file($logFile, FILE_IGNORE_NEW_LINES);

// Filter lines if needed
$displayLines = $allLines;
if ($filter === 'decrypt' && $allLines) {
    $displayLines = array_filter($allLines, function($line) {
        return strpos($line, '🔍 DECRYPTION') !== false || 
               strpos($line, 'DECRYPTION') !== false ||
               strpos($line, 'Encryption Key') !== false ||
               strpos($line, 'Key (hex)') !== false ||
               strpos($line, 'IV (hex)') !== false ||
               strpos($line, 'Base64') !== false ||
               strpos($line, 'Decryption Attempt') !== false ||
               strpos($line, 'DECRYPTION SUCCESSFUL') !== false ||
               strpos($line, 'DECRYPTION FAILED') !== false ||
               strpos($line, 'Decrypted password') !== false ||
               strpos($line, 'Method 1') !== false ||
               strpos($line, 'Method 2') !== false ||
               strpos($line, '✅') !== false ||
               strpos($line, '❌') !== false ||
               strpos($line, '═') !== false ||
               preg_match('/^\s*$/', $line); // Include empty lines
    });
}

// Get last N lines
$lastLines = array_slice($displayLines, -$lines_count);

?>
<!DOCTYPE html>
<html>
<head>
    <meta charset='utf-8'>
    <title>Debug Log Viewer</title>
    <style>
        body {
            font-family: 'Courier New', monospace;
            background-color: #1e1e1e;
            color: #e0e0e0;
            padding: 20px;
            margin: 0;
        }
        .header {
            background-color: #2d2d2d;
            border: 1px solid #444;
            padding: 15px;
            margin-bottom: 20px;
            border-radius: 5px;
        }
        .info {
            margin: 10px 0;
            font-size: 14px;
        }
        .log-content {
            background-color: #1e1e1e;
            border: 1px solid #444;
            padding: 15px;
            border-radius: 5px;
            line-height: 1.6;
            white-space: pre-wrap;
            word-wrap: break-word;
            max-height: 80vh;
            overflow-y: auto;
        }
        .controls {
            margin-bottom: 20px;
        }
        .button {
            background-color: #0066cc;
            color: white;
            padding: 8px 15px;
            margin: 5px 5px 5px 0;
            border: none;
            border-radius: 3px;
            cursor: pointer;
            font-family: monospace;
            text-decoration: none;
            display: inline-block;
        }
        .button:hover {
            background-color: #0052a3;
        }
        .button.active {
            background-color: #00aa00;
        }
        .success { color: #00ff00; }
        .error { color: #ff6b6b; }
        .info-text { color: #64b5f6; }
        h1 { margin: 0 0 20px 0; }
        h2 { margin-top: 0; }
    </style>
</head>
<body>
    <h1>🔍 Proxy Debug Log Viewer</h1>
    
    <div class="header">
        <div class="info"><strong>Log File:</strong> <?php echo basename($logFile); ?></div>
        <div class="info"><strong>File Size:</strong> <?php echo number_format($fileSize) . ' bytes'; ?></div>
        <div class="info"><strong>Last Modified:</strong> <?php echo date('Y-m-d H:i:s', filemtime($logFile)); ?></div>
        <div class="info"><strong>Total Lines:</strong> <?php echo count($allLines); ?></div>
        <div class="info"><strong>Filtered Lines:</strong> <?php echo count($displayLines); ?></div>
    </div>

    <div class="controls">
        <strong>Log File:</strong>
        <a href="?log=proxy&filter=<?php echo $filter; ?>&lines=<?php echo $lines_count; ?>" 
           class="button <?php echo ($logType === 'proxy') ? 'active' : ''; ?>">📋 Proxy</a>
        <a href="?log=proxy_child&filter=<?php echo $filter; ?>&lines=<?php echo $lines_count; ?>" 
           class="button <?php echo ($logType === 'proxy_child') ? 'active' : ''; ?>">👶 Proxy Child</a>
        
        <br><strong>Filter:</strong>
        <a href="?log=<?php echo $logType; ?>&filter=decrypt&lines=<?php echo $lines_count; ?>" 
           class="button <?php echo ($filter === 'decrypt') ? 'active' : ''; ?>">🔐 Decryption Only</a>
        <a href="?log=<?php echo $logType; ?>&filter=all&lines=<?php echo $lines_count; ?>" 
           class="button <?php echo ($filter === 'all') ? 'active' : ''; ?>">📊 All Logs</a>
        
        <br><strong>Lines:</strong>
        <a href="?log=<?php echo $logType; ?>&filter=<?php echo $filter; ?>&lines=50" 
           class="button <?php echo ($lines_count === 50) ? 'active' : ''; ?>">50</a>
        <a href="?log=<?php echo $logType; ?>&filter=<?php echo $filter; ?>&lines=100" 
           class="button <?php echo ($lines_count === 100) ? 'active' : ''; ?>">100</a>
        <a href="?log=<?php echo $logType; ?>&filter=<?php echo $filter; ?>&lines=500" 
           class="button <?php echo ($lines_count === 500) ? 'active' : ''; ?>">500</a>
        <a href="?log=<?php echo $logType; ?>&filter=<?php echo $filter; ?>&lines=9999" 
           class="button <?php echo ($lines_count === 9999) ? 'active' : ''; ?>">All</a>
        
        <br><a href="javascript:location.reload()" class="button">🔄 Refresh</a>
    </div>

    <div class="log-content">
        <?php 
        if (empty($lastLines)) {
            echo "No log entries found.\n";
        } else {
            $logText = implode("\n", $lastLines);
            // Highlight important sections
            $logText = preg_replace('/✅/', '<span class="success">✅</span>', $logText);
            $logText = preg_replace('/❌/', '<span class="error">❌</span>', $logText);
            $logText = preg_replace('/(🔍|🔐|🔑|📝|📊|📨|📤|🔙|⚡|🔗)/', '<span class="info-text">$1</span>', $logText);
            echo $logText;
        }
        ?>
    </div>
</body>
</html>
