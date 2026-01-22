# Linked Events API Implementation - Summary

## 🎯 Overview

The linked events feature now uses the **same execution API** as normal actions, with the distinction being the action name used in the swinfinity parameter.

## 📋 Implementation Details

### Flow Comparison

#### Normal Action (dv_creation_evt)
```
1. JavaScript fetches fresh token
2. Calls WebService/Execute API with:
   - action: "WEB_SERVICE_INFINITY"
   - swinfinity: "dv_creation_evt"
   - param: { evenement data }
3. API creates new event
```

#### Linked Events (dv_lister_evt)
```
1. JavaScript fetches fresh token
2. Calls proxy_child.php with event data
3. proxy_child.php:
   - Fetches fresh token from Auth API
   - Calls WebService/Execute API with:
     - action: "WEB_SERVICE_INFINITY"
     - swinfinity: "dv_lister_evt"
     - param: { utilisateur, tiers, RealiseOk }
4. API returns list of linked events
```

## 🔧 Technical Changes

### 1. proxy_child.php Updated
**File**: `src/proxy/proxy_child.php`

**Key Changes**:
- ✅ Loads environment variables from `.env` file
- ✅ Connects to database to get credentials
- ✅ Fetches fresh token from Auth API
- ✅ Calls WebService/Execute endpoint (NOT webhook)
- ✅ Uses action `dv_lister_evt` for listing events
- ✅ Supports `RealiseOk` parameter (0 = available, 1 = all events)
- ✅ Enhanced logging for debugging

**Flow**:
```php
// Step 1: Load .env and connect to database
// Step 2: Fetch credentials from database
// Step 3: Get authentication token
$token = callAuthApi($creds);

// Step 4: Build REST payload with dv_lister_evt
$payload = [
    "action" => "WEB_SERVICE_INFINITY",
    "access_token" => $token,
    "param" => json_encode([
        "action" => ["swinfinity" => "dv_lister_evt"],
        "data" => [
            "evenement" => [
                "utilisateur" => $user,
                "tiers" => $tiers,
                "RealiseOk" => $realiseOk
            ]
        ]
    ])
];

// Step 5: Call WebService/Execute API
echo callExecutionApi($payload);
```

### 2. taskpane.js Updated
**File**: `src/taskpane/taskpane.js`

**Key Changes**:
- ✅ `loadChildEventsForSav()` now fetches fresh token before API call
- ✅ Token passed implicitly (proxy_child.php gets it independently)
- ✅ `RealiseOk` parameter included in request
- ✅ Checkbox toggle maintains fresh token
- ✅ Better debug logging

**Code Example**:
```javascript
async function loadChildEventsForSav() {
  // Fetch fresh token
  cachedToken = await fetchTokenFromProxy();
  
  // Build payload
  const payload = {
    evenement: {
      utilisateur: cachedPayload.evenement.utilisateur,
      tiers: cachedPayload.evenement.tiers,
      RealiseOk: realiseOkFilter // 0 or 1
    }
  };
  
  // Call proxy_child.php
  // proxy_child.php will fetch its own token and call execution API
  const response = await fetch(PROXY_CHILD_URL, {
    method: "POST",
    body: JSON.stringify(payload)
  });
}
```

## 📊 API Call Comparison

### Normal Action
```
JavaScript
    ↓
Token API (proxy.php)
    ↓
Execution API
    ↓
Create Event (dv_creation_evt)
```

### Linked Events
```
JavaScript
    ↓
proxy_child.php (new flow)
    ├─ Get credentials from DB
    ├─ Call Token API
    └─ Call Execution API
        ↓
    List Events (dv_lister_evt)
```

## 🔐 Security Features

1. **Database Credentials**
   - All sensitive data stored encrypted in database
   - `.env` file used only for DB connection
   - No hardcoded API credentials

2. **Token Management**
   - Fresh token fetched for each operation
   - Token expires in 30 minutes
   - Automatic refresh on each action

3. **Password Encryption**
   - AES-256-CBC encryption
   - Passwords decrypted only when needed
   - Secure key management via `.env`

## 📝 Debug Logging

### Log Files
- **proxy.php**: `src/proxy/proxy_debug.log`
- **proxy_child.php**: `src/proxy/proxy_child_debug.log`

### What's Logged
✅ .env file detection and loading
✅ Database connection details
✅ Token fetching status
✅ API call details
✅ Response parsing
✅ Error details

### Example Log Entry
```
[2026-01-21 10:30:45] 🔐 Linked events request detected
[2026-01-21 10:30:45] ✅ Database connected
[2026-01-21 10:30:45] ✅ Credentials fetched and password decrypted
[2026-01-21 10:30:45] 📤 Calling Auth API
[2026-01-21 10:30:46] ✅ Token obtained successfully
[2026-01-21 10:30:46] 📤 Calling WebService/Execute API
[2026-01-21 10:30:46] 🔙 WebService Response (HTTP 200): {...}
```

## ✅ Testing Checklist

- [ ] Build project: `npm run build:dev`
- [ ] Test "Évènement lié" button
- [ ] Verify checkbox "Afficher tous les évènements" works
- [ ] Check debug logs for successful API calls
- [ ] Verify linked events load correctly (RealiseOk=0)
- [ ] Verify all events load when checkbox checked (RealiseOk=1)
- [ ] Verify events list when checkbox unchecked
- [ ] Select an event and confirm
- [ ] Verify event creation succeeds

## 🚀 Summary

✅ **Complete Implementation**
- Linked events use execution API (not webhook)
- Action uses `dv_lister_evt` for listing
- Token fetched fresh for each operation
- `RealiseOk` parameter controls filtering
- Database credentials used throughout
- Comprehensive logging for debugging

✅ **Consistent with Normal Actions**
- Same authentication flow
- Same execution API endpoint
- Same error handling
- Same security practices
