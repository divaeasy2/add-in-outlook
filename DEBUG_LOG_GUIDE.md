# Debug Log Guide - Database Connection Verification

## 📍 Where to Find Debug Logs

The proxy files create debug logs in the same directory as the PHP scripts:

### Main Proxy (Token & Actions)
- **Location**: `src/proxy/proxy_debug.log`
- **Purpose**: Logs token retrieval and API action requests

### Child Proxy (Linked Events)
- **Location**: `src/proxy/proxy_child_debug.log`
- **Purpose**: Logs linked events fetching with RealiseOk filter

## 🔍 What to Look For

### 1. .env File Detection
Look for this section in the logs:

```
[2026-01-21 10:30:45] 🔍 Looking for .env file...
[2026-01-21 10:30:45] 📍 Current script location: /var/www/html/src/proxy
[2026-01-21 10:30:45]   Checking: /var/www/html/src/proxy/../../.env (exists: YES)
[2026-01-21 10:30:45]   ✅ Found .env at: /var/www/html/src/proxy/../../.env
[2026-01-21 10:30:45] 📖 Reading .env file: /var/www/html/src/proxy/../../.env
[2026-01-21 10:30:45] 📏 File size: 145 bytes
```

**✅ If you see "YES" and "Found .env at" → .env file is being found correctly**

### 2. Environment Variables Loading
Look for this section:

```
[2026-01-21 10:30:45]   ✅ Set DB_HOST = ****(24 chars)
[2026-01-21 10:30:45]   ✅ Set DB_USER = ****(10 chars)
[2026-01-21 10:30:45]   ✅ Set DB_PASS = ****(15 chars)
[2026-01-21 10:30:45]   ✅ Set DB_NAME = ****(10 chars)
[2026-01-21 10:30:45]   ✅ Set ENCRYPTION_KEY = ****(11 chars)
[2026-01-21 10:30:45] ✅ .env file loaded successfully - 4 new variables loaded
```

**✅ If you see "Set DB_HOST", "Set DB_USER", etc. → Variables are loaded from .env**

### 3. Database Connection Variables
Look for this section:

```
[2026-01-21 10:30:45] ════════════════════════════════════════════
[2026-01-21 10:30:45] 🔗 DATABASE CONNECTION VARIABLES
[2026-01-21 10:30:45] ════════════════════════════════════════════
[2026-01-21 10:30:45] DB_HOST: ✅ SET = 'maisogv978.mysql.db'
[2026-01-21 10:30:45] DB_USER: ✅ SET = 'maisogv978'
[2026-01-21 10:30:45] DB_PASS: ✅ SET (15 chars)
[2026-01-21 10:30:45] DB_NAME: ✅ SET = 'maisogv978'
[2026-01-21 10:30:45] ════════════════════════════════════════════
[2026-01-21 10:30:45] ✅ All DB variables loaded from .env file successfully!
```

**✅ If you see "✅ SET" for all variables → Everything is loaded correctly from .env**

## ⚠️ Troubleshooting

### If you see:
```
[2026-01-21 10:30:45]   Checking: /var/www/html/src/proxy/../../.env (exists: NO)
[2026-01-21 10:30:45]   Checking: /var/www/html/src/proxy/../../../.env (exists: NO)
[2026-01-21 10:30:45] ❌ .env file NOT found in any of the checked paths!
[2026-01-21 10:30:45] ⚠️  Relying on system environment variables
```

**Action**: 
1. Check if `.env` file exists in project root: `ls -la /var/www/html/.env`
2. Verify file permissions: `chmod 644 /var/www/html/.env`
3. Check .env content: `cat /var/www/html/.env`

### If you see:
```
[2026-01-21 10:30:45] DB_HOST: ❌ NOT SET = ''
[2026-01-21 10:30:45] DB_USER: ❌ NOT SET = ''
[2026-01-21 10:30:45] ⚠️  DB_HOST not found in .env, using FALLBACK credentials
[2026-01-21 10:30:45] ⚠️  FALLBACK CONNECTION DETAILS:
```

**Action**:
1. Verify .env file has correct variables
2. Check for typos in variable names (must be exactly: `DB_HOST`, `DB_USER`, `DB_PASS`, `DB_NAME`)
3. Check for BOM (Byte Order Mark) in .env file - should be UTF-8 without BOM
4. Verify no leading/trailing spaces in variable names

### If you see:
```
[2026-01-21 10:30:45] ⏭️  Skipped DB_HOST (already set in system environment)
```

**Action**: 
- This is OK if system environment variables are already set
- But .env values will NOT override system environment variables
- To use .env values, unset system environment variables first

## 📊 How to Check Logs via SSH/Terminal

### View last 20 lines:
```bash
tail -20 /var/www/html/src/proxy/proxy_debug.log
tail -20 /var/www/html/src/proxy/proxy_child_debug.log
```

### View entire log:
```bash
cat /var/www/html/src/proxy/proxy_debug.log
cat /var/www/html/src/proxy/proxy_child_debug.log
```

### Monitor logs in real-time (follow new entries):
```bash
tail -f /var/www/html/src/proxy/proxy_debug.log
```

### Search for specific content:
```bash
grep "DB_HOST" /var/www/html/src/proxy/proxy_debug.log
grep "✅ .env file loaded" /var/www/html/src/proxy/proxy_debug.log
```

## 🎯 Success Indicators

You should see these markers in the logs when everything works correctly:

✅ `Found .env at:` - .env file was located
✅ `Set DB_HOST`, `Set DB_USER`, `Set DB_PASS`, `Set DB_NAME` - All DB variables loaded
✅ `.env file loaded successfully - 4 new variables loaded` - All 4 DB variables from .env
✅ `✅ All DB variables loaded from .env file successfully!` - DB connection using .env
✅ `✅ Database connected` - Successfully connected to database
✅ `✅ Credentials fetched and password decrypted` - Successfully got credentials from DB

## 🚀 Complete Workflow Log Example

Here's what a successful complete flow looks like:

```
[2026-01-21 10:30:45] 📨 Incoming request: {"token_request":true}
[2026-01-21 10:30:45] 🔍 Looking for .env file...
[2026-01-21 10:30:45] 📍 Current script location: /var/www/html/src/proxy
[2026-01-21 10:30:45]   Checking: /var/www/html/src/proxy/../../.env (exists: YES)
[2026-01-21 10:30:45]   ✅ Found .env at: /var/www/html/src/proxy/../../.env
[2026-01-21 10:30:45] 📖 Reading .env file: /var/www/html/src/proxy/../../.env
[2026-01-21 10:30:45] 📏 File size: 145 bytes
[2026-01-21 10:30:45]   ✅ Set DB_HOST = ****(24 chars)
[2026-01-21 10:30:45]   ✅ Set DB_USER = ****(10 chars)
[2026-01-21 10:30:45]   ✅ Set DB_PASS = ****(15 chars)
[2026-01-21 10:30:45]   ✅ Set DB_NAME = ****(10 chars)
[2026-01-21 10:30:45] ✅ .env file loaded successfully - 4 new variables loaded
[2026-01-21 10:30:45] 🔐 Token retrieval request detected
[2026-01-21 10:30:45] ════════════════════════════════════════════
[2026-01-21 10:30:45] 🔗 DATABASE CONNECTION VARIABLES
[2026-01-21 10:30:45] ════════════════════════════════════════════
[2026-01-21 10:30:45] DB_HOST: ✅ SET = 'maisogv978.mysql.db'
[2026-01-21 10:30:45] DB_USER: ✅ SET = 'maisogv978'
[2026-01-21 10:30:45] DB_PASS: ✅ SET (15 chars)
[2026-01-21 10:30:45] DB_NAME: ✅ SET = 'maisogv978'
[2026-01-21 10:30:45] ════════════════════════════════════════════
[2026-01-21 10:30:45] ✅ All DB variables loaded from .env file successfully!
[2026-01-21 10:30:45] 🔗 Attempting DB connection - Host: maisogv978.mysql.db, User: maisogv978, DB: maisogv978
[2026-01-21 10:30:46] ✅ Database connected
[2026-01-21 10:30:46] ✅ Credentials fetched and password decrypted
[2026-01-21 10:30:46] 📤 Calling Auth API: https://remote.divy-si.fr:8443/...
[2026-01-21 10:30:47] 🔙 Auth API Response (HTTP 200): {"error":0,"access_token":"eyJ0eXAi...
[2026-01-21 10:30:47] ✅ Token obtained successfully
```

This confirms that:
1. ✅ .env file was found
2. ✅ All DB credentials loaded from .env
3. ✅ Database connection successful
4. ✅ Credentials retrieved from database
5. ✅ API token obtained
