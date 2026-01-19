<?php
/**
 * MySQL Connection Diagnostic - Try Multiple Connection Methods
 * Tests different ways to connect to MySQL
 */

header('Content-Type: application/json');

$results = array();

// Test 1: localhost (socket - usually fails on remote)
$results['test_1_localhost_socket'] = testConnection('localhost', 3306, 'DivyAddIN', 'DivyAddIN', 'i$aKtRG48hjffr0?');

// Test 2: 127.0.0.1 (TCP loopback)
$results['test_2_127.0.0.1_tcp'] = testConnection('127.0.0.1', 3306, 'DivyAddIN', 'DivyAddIN', 'i$aKtRG48hjffr0?');

// Test 3: Check if PrestaShop config has the host
$results['test_3_prestashop_config'] = getPrestashopDbConfig();

// Test 4: Try to find MySQL socket location
$results['test_4_socket_locations'] = findMySQLSocket();

// Test 5: List common OVH MySQL hosts
$results['test_5_ovh_hosts'] = tryOVHCommonHosts('DivyAddIN', 'DivyAddIN', 'i$aKtRG48hjffr0?');

echo json_encode($results, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);

function testConnection($host, $port, $db, $user, $pass) {
    $result = array(
        'host' => $host,
        'port' => $port,
        'success' => false,
        'error' => ''
    );
    
    $conn = @mysqli_connect($host, $user, $pass, $db, $port);
    if ($conn === false) {
        $result['error'] = mysqli_connect_error();
    } else {
        $result['success'] = true;
        $result['message'] = 'Connected successfully!';
        mysqli_close($conn);
    }
    
    return $result;
}

function getPrestashopDbConfig() {
    $result = array(
        'status' => 'Not found',
        'config' => null
    );
    
    // Check common PrestaShop paths
    $paths = array(
        '/home/maisogv/www/config/settings.inc.php',
        '/home/maisogv/www/config/db.inc.php',
        '../../../config/settings.inc.php',
        '../../../config/db.inc.php'
    );
    
    foreach ($paths as $path) {
        if (file_exists($path)) {
            $result['status'] = 'Found at: ' . $path;
            // Try to extract DB_SERVER from PrestaShop config
            $content = file_get_contents($path);
            if (preg_match("/define\('_DB_SERVER_',\s*'([^']+)'\)/", $content, $matches)) {
                $result['DB_SERVER'] = $matches[1];
            }
            if (preg_match("/define\('_DB_NAME_',\s*'([^']+)'\)/", $content, $matches)) {
                $result['DB_NAME'] = $matches[1];
            }
            if (preg_match("/define\('_DB_USER_',\s*'([^']+)'\)/", $content, $matches)) {
                $result['DB_USER'] = $matches[1];
            }
            break;
        }
    }
    
    return $result;
}

function findMySQLSocket() {
    $result = array('sockets_found' => array());
    
    $commonPaths = array(
        '/var/run/mysqld/mysqld.sock',
        '/tmp/mysql.sock',
        '/var/mysql/mysql.sock',
        '/opt/local/var/run/mysqld/mysqld.sock'
    );
    
    foreach ($commonPaths as $path) {
        $exists = file_exists($path);
        $result['sockets_found'][$path] = $exists ? '✅ Exists' : '❌ Not found';
    }
    
    return $result;
}

function tryOVHCommonHosts($db, $user, $pass) {
    $result = array();
    
    // Common OVH MySQL hosts
    $hosts = array(
        'localhost',
        '127.0.0.1',
        'sql.ovh.net',
        'db.ovh.net',
        'mysql.ovh.net',
        // Generic OVH hosts
        'xxx.mysql.db',  // where xxx is account name
    );
    
    foreach ($hosts as $host) {
        $result[$host] = testConnection($host, 3306, $db, $user, $pass);
    }
    
    return $result;
}
?>
