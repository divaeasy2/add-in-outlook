<?php
/**
 * Database Connection Test - Using Static Values
 * Tests database connection with hardcoded credentials
 */

header('Content-Type: application/json');

$response = array(
    'success' => false,
    'message' => '',
    'config' => array(),
    'error' => null,
    'timestamp' => date('Y-m-d H:i:s')
);

// Static database configuration
$dbHost = 'localhost';
$dbPort = 3306;
$dbName = 'DivyADDIN';
$dbUser = 'DivyAddIN';
$dbPass = 'i$aKtRG48hjffr0?2';
$dbType = 'mysql';

// Store config (without password for security)
$response['config'] = array(
    'type' => $dbType,
    'host' => $dbHost,
    'port' => $dbPort,
    'database' => $dbName,
    'user' => $dbUser,
    'password' => $dbPass ? '***' : 'NOT SET'
);

// Validate required variables
$errors = array();
if (!$dbHost) $errors[] = "DB_HOST is not set";
if (!$dbName) $errors[] = "DB_NAME is not set";
if (!$dbUser) $errors[] = "DB_USER is not set";

if (!empty($errors)) {
    $response['error'] = "Configuration Errors: " . implode(", ", $errors);
    echo json_encode($response);
    exit;
}

// Test connection based on database type
try {
    switch (strtolower($dbType)) {
        case 'mysql':
        case 'mariadb':
            if (!extension_loaded('mysqli')) {
                throw new Exception("mysqli extension not loaded");
            }
            $conn = @mysqli_connect($dbHost, $dbUser, $dbPass, $dbName, $dbPort);
            if ($conn === false) {
                throw new Exception("MySQL Error: " . mysqli_connect_error());
            }
            
            // Test query
            $result = mysqli_query($conn, "SELECT 1 as test");
            if (!$result) {
                throw new Exception("Query failed: " . mysqli_error($conn));
            }
            
            // Get version
            $versionResult = mysqli_query($conn, "SELECT VERSION() as version");
            $versionRow = mysqli_fetch_assoc($versionResult);
            
            $response['success'] = true;
            $response['message'] = "Successfully connected to MySQL database";
            $response['version'] = $versionRow['version'] ?? 'Unknown';
            mysqli_close($conn);
            break;
        
        case 'pgsql':
        case 'postgresql':
            if (!extension_loaded('pgsql')) {
                throw new Exception("pgsql extension not loaded");
            }
            $connString = "host=$dbHost port=$dbPort dbname=$dbName user=$dbUser password=$dbPass";
            $conn = @pg_connect($connString);
            if ($conn === false) {
                throw new Exception("PostgreSQL Error: Connection failed");
            }
            
            // Test query
            $result = @pg_query($conn, "SELECT 1 as test");
            if (!$result) {
                throw new Exception("Query failed: " . pg_last_error($conn));
            }
            
            // Get version
            $versionResult = @pg_query($conn, "SELECT version() as version");
            $versionRow = @pg_fetch_assoc($versionResult);
            
            $response['success'] = true;
            $response['message'] = "Successfully connected to PostgreSQL database";
            $response['version'] = $versionRow['version'] ?? 'Unknown';
            pg_close($conn);
            break;
        
        case 'sqlsrv':
        case 'mssql':
            if (!extension_loaded('sqlsrv')) {
                throw new Exception("sqlsrv extension not loaded");
            }
            $serverName = $dbHost . "," . $dbPort;
            $connectionInfo = array(
                "Database" => $dbName,
                "Uid" => $dbUser,
                "PWD" => $dbPass,
                "ReturnDatesAsStrings" => true
            );
            $conn = @sqlsrv_connect($serverName, $connectionInfo);
            if ($conn === false) {
                throw new Exception("SQL Server Error: " . print_r(sqlsrv_errors(), true));
            }
            
            // Test query
            $result = @sqlsrv_query($conn, "SELECT 1 as test");
            if ($result === false) {
                throw new Exception("Query failed: " . print_r(sqlsrv_errors(), true));
            }
            
            $response['success'] = true;
            $response['message'] = "Successfully connected to SQL Server database";
            sqlsrv_close($conn);
            break;
        
        default:
            throw new Exception("Unsupported database type: $dbType");
    }
} catch (Exception $e) {
    $response['error'] = $e->getMessage();
}

echo json_encode($response, JSON_PRETTY_PRINT | JSON_UNESCAPED_SLASHES);
?>
