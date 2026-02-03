/* global Office */

let cachedPayload = null;
const MAX_EMAIL_SIZE = 10 * 1024 * 1024; // 10 MB
let statusTimeoutId = null;
let allEvents = []; // Store all events for filtering
let cachedToken = null; // Cache for API token
let tokenFetchInProgress = false; // Prevent multiple token requests
let realiseOkFilter = 0; // Track the RealiseOk filter state (0 = available only, 1 = all events)

Office.onReady(() => {
  const displayName = Office.context.mailbox.userProfile.displayName;
  document.getElementById("user").innerText = displayName;

  document.getElementById("userEmail").innerText =
    Office.context.mailbox.userProfile.emailAddress;
    
  // Set avatar with user initials
  setAvatarInitials(displayName);
    
  // SAV event - show options modal
  document.getElementById("btnSav").onclick = () => showSavOptions();
  
  // Options modal buttons
  document.getElementById("btnSavNew").onclick = () => sendSavNew();
  document.getElementById("btnSavLinked").onclick = () => loadChildEventsForSav();
  document.getElementById("btnSavCancel").onclick = () => hideSavOptions();
  
  // Other event types
  document.getElementById("btnComm").onclick = () => send("2");
  document.getElementById("btnDDP").onclick = () => send("3");
  document.getElementById("btnCDE").onclick = () => send("4");
  document.getElementById("btnDDI").onclick = () => send("5");

  // Child events modal buttons
  document.getElementById("confirmEvt").onclick = () => confirmLinkedEvent();
  document.getElementById("AnnuleEvt").onclick = () => returnFromLinkedEvents();
  
  // Search input for events
  document.getElementById("eventSearchInput").addEventListener("input", (e) => filterEvents(e.target.value));
  
  // Show all events checkbox
  document.getElementById("showAllEventsCheckbox").addEventListener("change", (e) => handleShowAllEventsToggle(e.target.checked));
  
  // Confirmation modal buttons
  document.getElementById("btnConfirmLinked").onclick = () => sendWithLinkedEvent();
  document.getElementById("btnCancelLinked").onclick = () => returnFromConfirmation();

  // Initialize: Just prepare email, token will be fetched on demand when user clicks an action
  debugLog("🚀 Initializing add-in...");
  try {
    enableButtons(); // Enable buttons from the start
    showStatus("✅ Add-in prêt - Cliquez sur une action pour continuer", "success");
    prepareEmail();
    debugLog("✅ Module complémentaire initialisé. Le jeton sera récupéré à la première action.");
  } catch (error) {
    debugLog(`❌ Échec de l'initialisation: ${error}`);
    showStatus(`❌ Erreur d'initialisation: ${error}`, "error");
    disableButtons();
  }
});

/* ======================
   STATUS
====================== */

function setAvatarInitials(displayName) {
  const avatar = document.getElementById("avatar");
  if (!displayName) {
    avatar.innerText = "?";
    return;
  }
  
  const names = displayName.trim().split(/\s+/);
  let initials = "";
  
  if (names.length >= 2) {
    // Get first letter of first name and last name
    initials = (names[0][0] + names[names.length - 1][0]).toUpperCase();
  } else if (names.length === 1) {
    // Single name: show first letter twice or just once
    initials = names[0][0].toUpperCase();
  }
  
  avatar.innerText = initials;
}

function showStatus(msg, type = "info") {
  const el = document.getElementById("status");
  el.className = `status ${type}`;
  el.innerText = msg;
  el.style.display = "block";
  
  // Clear previous timeout if exists
  if (statusTimeoutId) {
    clearTimeout(statusTimeoutId);
  }
  
  // Auto-hide status message after 3.5 seconds
  statusTimeoutId = setTimeout(() => {
    el.style.display = "none";
    statusTimeoutId = null;
  }, 3500);
}

function showLoading() {
  const overlay = document.getElementById("loadingOverlay");
  if (overlay) {
    overlay.style.display = "flex";
    // Disable all action buttons (including popup buttons) to prevent multiple clicks
    document.querySelectorAll(".action-btn").forEach(btn => {
      btn.disabled = true;
      btn.style.opacity = "0.6";
    });
  }
}

function hideLoading() {
  const overlay = document.getElementById("loadingOverlay");
  if (overlay) {
    overlay.style.display = "none";
    // Re-enable all action buttons (including popup buttons)
    document.querySelectorAll(".action-btn").forEach(btn => {
      btn.disabled = false;
      btn.style.opacity = "1";
    });
  }
}

function updateProgress(percent, stepText) {
  const progressFill = document.getElementById("progressFill");
  const progressStep = document.getElementById("progressStep");
  const loadingText = document.getElementById("loadingText");
  
  if (progressFill) {
    progressFill.style.width = percent + "%";
  }
  if (progressStep) {
    progressStep.innerText = stepText || "";
  }
  if (loadingText) {
    loadingText.innerText = "Traitement en cours...";
  }
}

/* Child Events Loading - Compact Spinner */
function showChildLoading() {
  const loadingState = document.getElementById("childLoadingState");
  const contentState = document.getElementById("childContentState");
  if (loadingState) {
    loadingState.style.display = "flex";
  }
  if (contentState) {
    contentState.classList.add("hidden");
  }
}

function hideChildLoading() {
  const loadingState = document.getElementById("childLoadingState");
  const contentState = document.getElementById("childContentState");
  if (loadingState) {
    loadingState.style.display = "none";
  }
  if (contentState) {
    contentState.classList.remove("hidden");
  }
}

function showChildHint(msg = "") {
  const hint = document.getElementById("savHint");
  if (!hint) return;
  if (!msg) {
    hint.style.display = "none";
    hint.innerText = "";
  } else {
    hint.innerText = msg;
    hint.style.display = "block";
  }
}

/* ======================
   PREPARE EMAIL
====================== */

function prepareEmail() {
  const item = Office.context.mailbox.item;
  const user = Office.context.mailbox.userProfile.emailAddress;

  item.body.getAsync(Office.CoercionType.Text, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) {
      showStatus("❌ Impossible de lire l'email", "error");
      return;
    }

    cachedPayload = {
      evenement: {
        type: "",
        utilisateur: user,
        tiers: item.from?.emailAddress || "",
        lib: item.subject || "",
        messageId: "",
        pj: "",
        evt_lie: ""
      }
    };

    debugLog("Payload initialized with messageId: " + item.itemId);

    document.getElementById("btnSav").disabled = false;
    document.getElementById("btnComm").disabled = false;

    showStatus("✅ Email prêt pour traitement", "info");
  });
}

/* ======================
   EMAIL WITH OFFICE API
====================== */

async function getEmailWithOfficeApi() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    
    debugLog("📧 MessageId: " + item.itemId);
    
    // Try to get callback token (for OAuth scenarios)
    if (item.getCallbackTokenAsync) {
      item.getCallbackTokenAsync({ isRest: true }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          debugLog("✅ Jeton obtenu");
          resolve({
            success: true,
            token: result.value,
            messageId: item.itemId,
            emailAddress: Office.context.mailbox.userProfile.emailAddress
          });
        } else {
          debugLog("⚠️ Erreur de jeton: " + result.error.message);
          resolve({
            success: false,
            messageId: item.itemId
          });
        }
      });
    } else {
      debugLog("⚠️ Erreur de jeton: t.getCallbackTokenAsync n'est pas une fonction");
      resolve({
        success: false,
        messageId: item.itemId
      });
    }
  });
}

/* ======================
   EMAIL RETRIEVAL WITH ATTACHMENTS
====================== */

async function getEmailContent() {
  return new Promise((resolve) => {
    const item = Office.context.mailbox.item;
    let emailData = {
      subject: item.subject || "",
      from: item.from?.emailAddress || "",
      to: Office.context.mailbox.userProfile.emailAddress || "",
      date: new Date().toUTCString(),
      body: "",
      bodyType: "text",
      attachments: []
    };

    let bodyLoaded = false;
    let attachmentsLoaded = false;

    function checkComplete() {
      if (bodyLoaded && attachmentsLoaded) {
        resolve(emailData);
      }
    }

    // Try to get HTML body first (preserves images and formatting)
    item.body.getAsync(Office.CoercionType.Html, (res) => {
      if (res.status === Office.AsyncResultStatus.Succeeded && res.value) {
        emailData.body = res.value;
        emailData.bodyType = "html";
        debugLog("✅ Corps de l'email récupéré (HTML)");
        bodyLoaded = true;
      } else {
        // Fallback to plain text
        item.body.getAsync(Office.CoercionType.Text, (textRes) => {
          if (textRes.status === Office.AsyncResultStatus.Succeeded) {
            emailData.body = textRes.value;
            emailData.bodyType = "text";
            debugLog("✅ Corps de l'email récupéré (Texte - alternative)");
          } else {
            debugLog("⚠️ Échec de la récupération du corps");
            emailData.body = "[Impossible de récupérer le corps]";
          }
          bodyLoaded = true;
          checkComplete();
        });
        return;
      }
      bodyLoaded = true;
      checkComplete();
    });

    // Get attachments
    if (item.attachments && item.attachments.length > 0) {
      debugLog(`📎 ${item.attachments.length} pièce(s) jointe(s) trouvée(s)`);
      let loadedCount = 0;
      
      item.attachments.forEach((att, idx) => {
        try {
          // Get attachment data
          item.getAttachmentContentAsync(att.id, (result) => {
            debugLog(`📥 Traitement de la pièce jointe: ${att.name}`);
            
            if (result.status === Office.AsyncResultStatus.Succeeded) {
              try {
                let binaryData = result.value;
                
                if (!binaryData) {
                  debugLog(`❌ No data in result.value for ${att.name}`);
                  loadedCount++;
                  if (loadedCount === item.attachments.length) {
                    attachmentsLoaded = true;
                    checkComplete();
                  }
                  return;
                }
                
                // Detect attachment type
                const isEmailAttachment = att.contentType && 
                  (att.contentType.includes('message/rfc822') || 
                   att.contentType.includes('application/vnd.ms-outlook') ||
                   att.name?.endsWith('.eml') ||
                   att.name?.endsWith('.msg'));
                
                if (isEmailAttachment) {
                  debugLog(`   - Type: EMAIL ATTACHMENT (${att.contentType})`);
                }
                
                debugLog(`   - Data type: ${typeof binaryData} (${binaryData instanceof ArrayBuffer ? 'ArrayBuffer' : 'other'})`);
                
                let base64Data = '';
                let dataLen = 0;
                
                // Convert binary to base64
                if (binaryData instanceof ArrayBuffer) {
                  dataLen = binaryData.byteLength;
                  debugLog(`   - Size: ${dataLen} bytes`);
                  const uint8 = new Uint8Array(binaryData);
                  let binaryStr = '';
                  for (let i = 0; i < uint8.length; i++) {
                    binaryStr += String.fromCharCode(uint8[i]);
                  }
                  base64Data = btoa(binaryStr);
                  debugLog(`✅ Attachment loaded: ${att.name} (${dataLen} bytes → ${base64Data.length} base64)`);
                } else if (typeof binaryData === 'string') {
                  // Already base64
                  base64Data = binaryData;
                  dataLen = binaryData.length;
                  debugLog(`✅ Attachment loaded: ${att.name} (${dataLen} chars base64)`);
                } else if (typeof binaryData === 'object') {
                  // Outlook Desktop might return object with nested data
                  debugLog(`   - Checking object properties...`);
                  
                  // Check for common data properties - priority order for email attachments
                  let extractedData = null;
                  if (isEmailAttachment && binaryData.content) {
                    extractedData = binaryData.content;
                    debugLog(`   - Found .content property (email attachment)`);
                  } else if (binaryData.content) {
                    extractedData = binaryData.content;
                    debugLog(`   - Found .content property`);
                  } else if (binaryData.data) {
                    extractedData = binaryData.data;
                    debugLog(`   - Found .data property`);
                  } else if (binaryData.value) {
                    extractedData = binaryData.value;
                    debugLog(`   - Found .value property`);
                  } else if (binaryData.bytes) {
                    extractedData = binaryData.bytes;
                    debugLog(`   - Found .bytes property`);
                  } else {
                    extractedData = JSON.stringify(binaryData);
                    debugLog(`   - Using stringified object`);
                  }
                  
                  if (extractedData instanceof ArrayBuffer) {
                    dataLen = extractedData.byteLength;
                    const uint8 = new Uint8Array(extractedData);
                    let binaryStr = '';
                    for (let i = 0; i < uint8.length; i++) {
                      binaryStr += String.fromCharCode(uint8[i]);
                    }
                    base64Data = btoa(binaryStr);
                    debugLog(`✅ Attachment loaded: ${att.name} (${dataLen} bytes → ${base64Data.length} base64)`);
                  } else if (typeof extractedData === 'string') {
                    // For email attachments, validate it's proper base64
                    if (isEmailAttachment && extractedData.length > 0) {
                      // Validate base64 for email attachments
                      const base64Regex = /^[A-Za-z0-9+/=]*$/;
                      if (base64Regex.test(extractedData)) {
                        base64Data = extractedData;
                        dataLen = extractedData.length;
                        debugLog(`✅ Attachment loaded: ${att.name} (string: ${dataLen} chars, email format)`);
                      } else {
                        // Raw text email content - encode to base64
                        base64Data = btoa(extractedData);
                        dataLen = extractedData.length;
                        debugLog(`✅ Attachment loaded: ${att.name} (raw text → base64: ${dataLen} chars)`);
                      }
                    } else {
                      base64Data = extractedData;
                      dataLen = extractedData.length;
                      debugLog(`✅ Attachment loaded: ${att.name} (string: ${dataLen} chars)`);
                    }
                  } else if (extractedData instanceof Uint8Array) {
                    // Handle Uint8Array from email attachments
                    dataLen = extractedData.length;
                    let binaryStr = '';
                    for (let i = 0; i < extractedData.length; i++) {
                      binaryStr += String.fromCharCode(extractedData[i]);
                    }
                    base64Data = btoa(binaryStr);
                    debugLog(`✅ Attachment loaded: ${att.name} (Uint8Array: ${dataLen} bytes → base64)`);
                  }
                }
                
                if (base64Data && base64Data.length > 0) {
                  // For email attachments, ensure correct MIME type
                  let finalContentType = att.contentType || "application/octet-stream";
                  if (isEmailAttachment && !finalContentType.includes('message')) {
                    finalContentType = "message/rfc822";
                  }
                  
                  const attData = {
                    name: att.name || `attachment_${idx}`,
                    contentType: finalContentType,
                    data: base64Data
                  };
                  emailData.attachments.push(attData);
                  debugLog(`✅ Attachment added to email data (type: ${finalContentType})`);
                } else {
                  debugLog(`⚠️ Base64 data is empty for ${att.name}`);
                }
              } catch (encodeErr) {
                debugLog(`❌ Échec de l'encodage de la pièce jointe ${att.name}: ${encodeErr.message}`);
              }
            } else {
              debugLog(`⚠️ Échec du chargement de la pièce jointe: ${att.name}`);
              if (result.error) debugLog(`   Error: ${result.error.message}`);
            }
            
            loadedCount++;
            if (loadedCount === item.attachments.length) {
              // Deduplicate attachments by name + content to reduce payload size
              const seenAttachments = new Map();
              const uniqueAttachments = [];
              
              emailData.attachments.forEach(att => {
                // Create a hash key from name and first 100 chars of data
                const key = att.name + '_' + att.data.substring(0, 100);
                
                if (!seenAttachments.has(key)) {
                  seenAttachments.set(key, true);
                  uniqueAttachments.push(att);
                  debugLog(`✅ Keeping: ${att.name}`);
                } else {
                  debugLog(`⏭️ Skipping duplicate: ${att.name}`);
                }
              });
              
              const removed = emailData.attachments.length - uniqueAttachments.length;
              if (removed > 0) {
                debugLog(`✅ Deduplication: Removed ${removed} duplicate(s), kept ${uniqueAttachments.length}`);
              }
              
              emailData.attachments = uniqueAttachments;
              attachmentsLoaded = true;
              checkComplete();
            }
          });
        } catch (e) {
          debugLog(`❌ Error processing attachment ${att.name}: ${e.message}`);
          loadedCount++;
          if (loadedCount === item.attachments.length) {
            attachmentsLoaded = true;
            checkComplete();
          }
        }
      });
    } else {
      debugLog("📎 No attachments");
      attachmentsLoaded = true;
      checkComplete();
    }
  });
}

/* ======================
   UTILITY: Wrap base64 at 76 chars per line (RFC 2045)
====================== */
function wrapBase64(base64String) {
  const lines = [];
  for (let i = 0; i < base64String.length; i += 76) {
    lines.push(base64String.substring(i, i + 76));
  }
  return lines.join('\r\n') + '\r\n';
}

/* ======================
   BUILD EMAIL (.eml) WITH ATTACHMENTS
====================== */

function buildEmailBase64(item, emailContent) {
  const isHtml = emailContent.bodyType === "html";
  const boundary = "----=_Part_" + Math.random().toString(36).substring(2, 15);
  
  // Use array to collect parts - much faster than string concatenation for large payloads
  const parts = [];
  
  // Headers
  parts.push(`From: ${item.from?.emailAddress || ""}`);
  parts.push(`To: ${Office.context.mailbox.userProfile.emailAddress}`);
  parts.push(`Subject: ${emailContent.subject || ""}`);
  parts.push(`Date: ${emailContent.date}`);
  parts.push(`MIME-Version: 1.0`);

  // If there are attachments, use multipart format
  if (emailContent.attachments && emailContent.attachments.length > 0) {
    parts.push(`Content-Type: multipart/mixed; boundary="${boundary}"`);
    parts.push('');
    parts.push(`--${boundary}`);
    parts.push(`Content-Type: ${isHtml ? 'text/html' : 'text/plain'}; charset=UTF-8`);
    parts.push(`Content-Transfer-Encoding: quoted-printable`);
    parts.push('');
    parts.push(emailContent.body);
    
    // Add each attachment
    emailContent.attachments.forEach((att) => {
      parts.push('');
      parts.push(`--${boundary}`);
      parts.push(`Content-Type: ${att.contentType}`);
      parts.push(`Content-Transfer-Encoding: base64`);
      parts.push(`Content-Disposition: attachment; filename="${att.name}"`);
      parts.push('');
      
      // att.data should already be base64 string from our encoding above
      if (typeof att.data === 'string' && att.data.length > 0) {
        // Data is already base64, wrap it at 76 chars per line
        const wrapped = wrapBase64(att.data).trimEnd(); // Remove trailing CRLF that wrapBase64 adds
        parts.push(wrapped);
      } else {
        debugLog(`⚠️ Attachment ${att.name} has no data`);
      }
    });
    
    parts.push('');
    parts.push(`--${boundary}--`);
  } else {
    // Simple single-part email
    parts.push(`Content-Type: ${isHtml ? 'text/html' : 'text/plain'}; charset=UTF-8`);
    parts.push(`Content-Transfer-Encoding: quoted-printable`);
    parts.push('');
    parts.push(emailContent.body);
  }
  
  // Join all parts with CRLF line endings
  const eml = parts.join('\r\n');

  const size = new Blob([eml]).size;
  if (size > MAX_EMAIL_SIZE) {
    debugLog(`⚠️ Email size: ${(size / 1024 / 1024).toFixed(2)}MB exceeds limit`);
    return null;
  }

  try {
    return btoa(unescape(encodeURIComponent(eml)));
  } catch (e) {
    debugLog(`❌ Error encoding email: ${e.message}`);
    return null;
  }
}


/* ======================
   PARSER API
====================== */

function parseWeirdApiResponse(raw) {
  let n1;
  try { 
    n1 = JSON.parse(raw); 
  } catch {
    return { ok:false, error:"N1 n'est pas JSON", raw };
  }

  let n2 = n1.raw || n1.response || raw;

  let cleaned = n2
    .replace(/\\"/g, '"')
    .replace(/"{/g, '{')
    .replace(/}"/g, '}')
    .replace(/""result":/g, '"result":')
    .replace(/"result":"result":/g, '"result":')
    .replace(/"result":""/g, '"result":')
    .replace(/"result":\s*"({)/g, '"result":$1')
    .trim();

  debugLog("🔧 Nettoyé:\n" + cleaned);

  let n3;
  try { n3 = JSON.parse(cleaned); }
  catch {
    return {
      ok:false,
      error:"❌ Impossible de parser N2 → JSON",
      cleaned
    };
  }

  const events =
    n3.Evenements ||
    n3.evenements ||
    (n3.response && n3.response.Evenements);

  if (!events) return { ok:false, error:"❌ Aucun évènement trouvé", json:n3 };

  return { ok: true, count: events.length, events };
}


function debugLog(msg){
  const box = document.getElementById("debug");
  // box.style.display = "block";
  box.innerText += "\n" + msg;
}

/* ======================
   SAV EVENT WORKFLOW
====================== */

function showSavOptions() {
  document.getElementById("savOptionsModal").style.display = "block";
  disablePrimaryButtons();
}

function hideSavOptions() {
  document.getElementById("savOptionsModal").style.display = "none";
  enablePrimaryButtons();
  showStatus("");
  document.getElementById("status").style.display = "none";
}

async function sendSavNew() {
  hideSavOptions();
  await send("1");
}

async function loadChildEventsForSav() {
  document.getElementById("status").style.display = "block";
  if (!cachedPayload) return showStatus("⚠️ Aucun email prêt", "error");

  showStatus("⏳ Vérification des évènements ...", "info");
  
  // Show child popup with loading spinner
  document.getElementById("childPopup").style.display = "block";
  document.getElementById("savOptionsModal").style.display = "none";
  showChildLoading();

  const payload = {
    evenement: {
      utilisateur: cachedPayload.evenement.utilisateur,
      tiers: cachedPayload.evenement.tiers,
      RealiseOk: realiseOkFilter // Include the RealiseOk filter (0 = available, 1 = all)
    }
  };

  debugLog("📤 Sending RealiseOk=" + realiseOkFilter + " to API");
  debugLog("📤 Full Payload: " + JSON.stringify(payload));

  const res = await fetch("https://addin-divalto.divy-si.fr/ASFLUID/outlook/proxy/proxy_child.php", {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=UTF-8" },
    body: JSON.stringify(payload)
  }).then(r => r.text()).catch(() => null);

  if (!res) {
    hideChildLoading(); // Hide child loading animation
    document.getElementById("childPopup").style.display = "none";
    return showStatus("❌ Erreur réseau", "error");
  }

  const parsed = parseWeirdApiResponse(res);
  const popup = document.getElementById("childPopup");
  const select = document.getElementById("childSelect");
  const evtCount = document.getElementById("evtCount");
  const searchInput = document.getElementById("eventSearchInput");
  const showAllCheckbox = document.getElementById("showAllEventsCheckbox");

  if (!parsed.ok) {
    hideChildLoading(); // Hide child loading animation
    popup.style.display = "none";
    return showStatus("🔴 " + parsed.error, "error");
  }

  // Store events for filtering
  allEvents = parsed.events;

  // Reset checkbox state
  showAllCheckbox.checked = realiseOkFilter === 1;
  
  // Clear and rebuild select with proper encoding
  select.innerHTML = `<option value="">-- Choisissez un évènement --</option>`;
  parsed.events.forEach(evt => {
    const opt = document.createElement("option");
    opt.value = evt.evtNo;
    // Ensure proper encoding of event text
    const eventCode = String(evt.evtNo || "");
    const eventLib = String(evt.lib || "(sans lib)");
    opt.textContent = `${eventCode} - ${eventLib}`;
    select.appendChild(opt);
  });

  evtCount.innerText = `${parsed.count} évènements trouvés`;
  
  // Show search input if events exist
  if (parsed.count > 0) {
    searchInput.style.display = "block";
    searchInput.value = "";
  } else {
    searchInput.style.display = "none";
  }
  
  hideChildLoading(); // Hide child loading animation and show content
  disablePrimaryButtons();
  
  const filterStatus = realiseOkFilter === 1 ? "tous les évènements" : "évènements disponibles";
  showStatus(`🟢 ${parsed.count} ${filterStatus} récupérés`, "success");
}

function filterEvents(searchTerm) {
  const select = document.getElementById("childSelect");
  const searchValue = searchTerm.toLowerCase().trim();
  
  // Clear current options (except placeholder)
  select.innerHTML = `<option value="">-- Choisissez un évènement --</option>`;
  
  // Filter events
  const filteredEvents = allEvents.filter(evt => {
    const eventCode = String(evt.evtNo || "").toLowerCase();
    const eventLib = String(evt.lib || "").toLowerCase();
    return eventCode.includes(searchValue) || eventLib.includes(searchValue);
  });
  
  // Add filtered events to select
  filteredEvents.forEach(evt => {
    const opt = document.createElement("option");
    opt.value = evt.evtNo;
    const eventCode = String(evt.evtNo || "");
    const eventLib = String(evt.lib || "(sans lib)");
    opt.textContent = `${eventCode} - ${eventLib}`;
    select.appendChild(opt);
  });
  
  // Update count
  const evtCount = document.getElementById("evtCount");
  if (searchValue) {
    evtCount.innerText = `${filteredEvents.length} évènements trouvés`;
  } else {
    evtCount.innerText = `${allEvents.length} évènements trouvés`;
  }
}

async function handleShowAllEventsToggle(isChecked) {
  debugLog(`📋 Show all events toggle: ${isChecked ? 'ON (RealiseOk=1)' : 'OFF (RealiseOk=0)'}`);
  
  // Update the filter state
  realiseOkFilter = isChecked ? 1 : 0;
  
  // Reload events with new filter
  showStatus(`⏳ Chargement des ${isChecked ? 'tous les' : 'évènements disponibles'}...`, "info");
  await loadChildEventsForSav();
}

function confirmLinkedEvent() {
  const select = document.getElementById("childSelect");
  const chosen = select.value;
  
  if (!chosen) {
    return showStatus("⚠️ Sélectionnez un évènement", "error");
  }

  // Store the selected event
  cachedPayload.evenement.evt_lie = chosen;
  
  // Get the event details for display
  const selectedOption = select.options[select.selectedIndex];
  const eventText = selectedOption.textContent;
  
  // Hide child events popup
  document.getElementById("childPopup").style.display = "none";
  
  // Show confirmation modal
  document.getElementById("confirmLinkedText").innerText = 
    `Confirmer l'évènement:\n${eventText}`;
  document.getElementById("confirmLinkedModal").style.display = "block";
  disablePrimaryButtons();
}

async function sendWithLinkedEvent() {
  document.getElementById("confirmLinkedModal").style.display = "none";
  await send("1");
}

function returnFromConfirmation() {
  document.getElementById("confirmLinkedModal").style.display = "none";
  cachedPayload.evenement.evt_lie = "";
  
  // Show child events popup again
  document.getElementById("childPopup").style.display = "block";
  disablePrimaryButtons();
  showStatus("⏳ Retour à la sélection...", "info");
}

function returnFromLinkedEvents() {
  document.getElementById("childPopup").style.display = "none";
  enablePrimaryButtons();
  cachedPayload.evenement.evt_lie = "";
  
  // Show options modal again
  document.getElementById("savOptionsModal").style.display = "block";
  disablePrimaryButtons();
  showStatus("");
  document.getElementById("status").style.display = "none";
}

/* ======================
   TOKEN MANAGEMENT
====================== */

async function fetchTokenFromProxy() {
  if (tokenFetchInProgress) {
    debugLog("⏳ Token fetch already in progress...");
    // Wait for existing fetch to complete
    return new Promise((resolve, reject) => {
      const checkInterval = setInterval(() => {
        if (cachedToken && !tokenFetchInProgress) {
          clearInterval(checkInterval);
          resolve(cachedToken);
        }
      }, 100);
    });
  }

  tokenFetchInProgress = true;
  debugLog("🔐 Fetching authentication token from proxy...");
  
  try {
    const response = await fetch("https://addin-divalto.divy-si.fr/ASFLUID/outlook/proxy/proxy.php", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ token_request: true }) // Explicit token request signal
    });

    debugLog(`📊 Token response status: ${response.status}`);

    if (!response.ok) {
      const errorText = await response.text();
      debugLog(`❌ Token fetch failed: HTTP ${response.status}`);
      debugLog(`Error details: ${errorText.substring(0, 200)}`);
      throw new Error(`HTTP ${response.status}: ${errorText}`);
    }

    const data = await response.json();
    
    if (!data.ok) {
      debugLog(`❌ Token API error: ${data.error}`);
      throw new Error(data.error || "Token retrieval failed");
    }

    if (!data.token) {
      debugLog(`❌ No token in response`);
      throw new Error("No token provided in response");
    }

    debugLog(`✅ Token obtained successfully`);
    debugLog(`🔑 Token length: ${data.token.length} chars`);
    
    tokenFetchInProgress = false;
    return data.token;

  } catch (error) {
    tokenFetchInProgress = false;
    debugLog(`❌ Token fetch error: ${error.message}`);
    throw error;
  }
}

function enableButtons() {
  document.querySelectorAll(".action-btn").forEach((btn) => {
    btn.disabled = false;
    btn.style.opacity = "1";
  });
  debugLog("✅ All action buttons enabled");
}

function disableButtons() {
  document.querySelectorAll(".action-btn").forEach((btn) => {
    btn.disabled = true;
    btn.style.opacity = "0.5";
  });
  debugLog("⛔ All action buttons disabled");
}

function disablePrimaryButtons() {
  // Disable only the main action buttons (not popup buttons)
  const primaryBtns = ["btnSav", "btnComm", "btnDDP", "btnCDE", "btnDDI"];
  primaryBtns.forEach((id) => {
    const btn = document.getElementById(id);
    if (btn) {
      btn.disabled = true;
      btn.style.opacity = "0.5";
    }
  });
}

function enablePrimaryButtons() {
  // Enable only the main action buttons (not popup buttons)
  const primaryBtns = ["btnSav", "btnComm", "btnDDP", "btnCDE", "btnDDI"];
  primaryBtns.forEach((id) => {
    const btn = document.getElementById(id);
    if (btn) {
      btn.disabled = false;
      btn.style.opacity = "1";
    }
  });
}

/* ======================
   SEND
====================== */

async function send(type) {
  document.getElementById("status").style.display = "block";
  if (!cachedPayload) return showStatus("⚠️ Aucun email prêt", "error");
  
  showLoading(); // Start loading animation
  updateProgress(5, "Récupération du token d'authentification...");
  
  // Fetch fresh token before each action (tokens expire in 30 minutes)
  try {
    debugLog("🔐 Fetching fresh authentication token for this action...");
    cachedToken = await fetchTokenFromProxy();
    debugLog("✅ Token fetched successfully");
  } catch (error) {
    debugLog(`❌ Failed to fetch token: ${error.message}`);
    showStatus(`❌ Erreur d'authentification: ${error.message}`, "error");
    hideLoading();
    return;
  }

  updateProgress(10, "Initialisation...");
  try {
    const item = Office.context.mailbox.item;

    showStatus("⌛ Récupération des données...", "info");

    // Step 1: Get Office API messageId
    debugLog("📝 Étape 1: Récupération de l'ID du message API Office...");
    updateProgress(15, "Récupération de l'ID du message...");
    const tokenData = await getEmailWithOfficeApi();
    
    if (tokenData.success) {
      debugLog(`✅ Jeton Office obtenu pour: ${tokenData.emailAddress}`);
      debugLog(`📧 Message ID: ${tokenData.messageId}`);
      cachedPayload.evenement.messageId = tokenData.messageId;
    } else {
      debugLog(`⚠️ Jeton Office non disponible, utilisation de l'ID alternatif`);
      cachedPayload.evenement.messageId = item.itemId;
    }

    // Step 2: Get email content with attachments
    debugLog("📧 Étape 2: Récupération du contenu de l'email et des pièces jointes...");
    updateProgress(25, "Lecture du contenu de l'email...");
    const emailContent = await getEmailContent();
    
    debugLog(`✅ Email sender: ${emailContent.from}`);
    debugLog(`📧 Subject: ${emailContent.subject}`);
    debugLog(`📄 Body type: ${emailContent.bodyType} (${emailContent.bodyType === "html" ? "HTML with formatting" : "Plain text"})`);
    if (emailContent.attachments && emailContent.attachments.length > 0) {
      debugLog(`📎 ${emailContent.attachments.length} piece(s) jointe(s):`);
      emailContent.attachments.forEach(att => {
        debugLog(`   - ${att.name} (${att.contentType})`);
      });
    } else {
      debugLog(`📎 Aucune pièce jointe`);
    }
    
    // Step 3: Build email in base64 format
    debugLog("🔐 Étape 3: Encodage de l'email en base64...");
    updateProgress(50, "Encodage de l'email...");
    const emailBase64 = buildEmailBase64(item, emailContent);
    
    if (!emailBase64) {
      debugLog("⚠️ Email trop volumineux, envoi sans pièces jointes");
      cachedPayload.evenement.pj = "";
    } else {
      debugLog(`✅ Email encodé: ${emailBase64.length} octets`);
      const sizeMB = (emailBase64.length / 1024 / 1024).toFixed(2);
      if (emailBase64.length > 1000000) {
        debugLog(`⚠️ Payload volumineux détecté: ${sizeMB}MB`);
      }
      cachedPayload.evenement.pj = emailBase64;
    }

    cachedPayload.evenement.type = type;
    cachedPayload.evenement.evt_lie = cachedPayload.evenement.evt_lie || "";

    // Step 4: Prepare and send REST API request
    debugLog(`🚀 Étape 4: Préparation de l'appel API REST (type: ${type})...`);
    updateProgress(70, "Préparation du payload...");
    showStatus("🚀 Envoi vers l'API...", "info");

    // Build REST API payload
    const restPayload = {
      action: "WEB_SERVICE_INFINITY",
      access_token: cachedToken,
      param: JSON.stringify({
        action: { swinfinity: "dv_creation_evt" },
        data: { evenement: cachedPayload.evenement }
      })
    };

    debugLog(`📋 Structure du payload: action=${restPayload.action}, token=${restPayload.access_token.substring(0, 20)}...`);
    debugLog(`📄 Paramètres d'événement: type=${type}, utilisateur=${cachedPayload.evenement.utilisateur}`);

    updateProgress(75, "Transmission vers le serveur...");
    
    // Send to REST API
    const startFetch = performance.now();
    const res = await fetch("https://addin-divalto.divy-si.fr/ASFLUID/outlook/proxy/proxy.php", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(restPayload)
    });
    
    const totalFetchTime = performance.now() - startFetch;
    debugLog(`⏱️ Requête API effectuée en ${totalFetchTime.toFixed(2)}ms`);
    debugLog(`📊 Statut de la réponse: ${res.status}`);
    
    // Step 5: Parse response
    updateProgress(90, "Traitement de la réponse...");
    const text = await res.text();
    debugLog(`📥 Taille de la réponse: ${text.length} octets`);

    if (!res.ok) {
      debugLog(`❌ Erreur serveur: HTTP ${res.status}`);
      debugLog(`Response: ${text.substring(0, 300)}`);
      showStatus(`❌ Erreur serveur (${res.status})`, "error");
      hideLoading();
      return;
    }

    let parsed;
    try {
      parsed = JSON.parse(text);
      debugLog(`✅ Réponse analysée avec succès`);
    } catch (e) {
      debugLog(`❌ Erreur d'analyse JSON: ${e.message}`);
      debugLog(`Raw response: ${text.substring(0, 200)}`);
      showStatus("❌ Réponse invalide du serveur", "error");
      hideLoading();
      return;
    }

    // Step 6: Extract and validate response
    updateProgress(95, "Finalisation...");
    debugLog(`🔍 Analyse de la réponse: ${JSON.stringify(parsed)}`);

    // Parse the response structure - it may contain nested JSON strings
    let resultCode = null;
    let eventNo = null;
    let errorMessage = null;

    // Check for error code at top level
    if (parsed.error !== undefined && parsed.error !== 0) {
      debugLog(`❌ L'API a renvoyé le code d'erreur: ${parsed.error}`);
      showStatus(`❌ Erreur serveur (${parsed.error})`, "error");
      hideLoading();
      return;
    }

    // Parse the result field if it exists
    if (parsed.result) {
      try {
        // The result might be a JSON string within a string
        let resultStr = parsed.result;
        debugLog(`📋 Chaîne de résultat brute: ${resultStr.substring(0, 100)}...`);

        // Try to extract resultcode and EvtNo from the result string
        const resultcodeMatch = resultStr.match(/"resultcode"\s*:\s*"([^"]*)"/) || 
                               resultStr.match(/"resultcode"\s*:\s*(\d+)/);
        const evtNoMatch = resultStr.match(/"EvtNo"\s*:\s*"([^"]*)"/) || 
                          resultStr.match(/"EvtNo"\s*:\s*(\d+)/);
        const errorMatch = resultStr.match(/"errormessage"\s*:\s*"([^"]*)"/);

        if (resultcodeMatch) {
          resultCode = resultcodeMatch[1];
          debugLog(`✅ Code de résultat extrait: ${resultCode}`);
        }

        if (evtNoMatch) {
          eventNo = evtNoMatch[1];
          debugLog(`✅ Événement extrait: ${eventNo}`);
        }

        if (errorMatch) {
          errorMessage = errorMatch[1];
          debugLog(`📝 Error message: ${errorMessage}`);
        }
      } catch (e) {
        debugLog(`⚠️ Erreur lors de l'analyse du champ de résultat: ${e.message}`);
      }
    }

    // Determine success/failure - prioritize extracted resultCode over parsed.error
    if (resultCode) {
      // We extracted a resultCode, use it to determine success/failure
      if (resultCode === "0") {
        // Success case
        debugLog(`✅ SUCCÈS - Événement créé avec succès`);
        if (eventNo) {
          debugLog(`📌 Numéro d'événement: ${eventNo}`);
          updateProgress(100, "✅ Succès!");
          hideLoading();
          showStatus(`🎉 Succès - Évènement: ${eventNo}`, "success");
        } else if (errorMessage) {
          // If we have additional info in errormessage, show it as part of success
          debugLog(`📝 Informations: ${errorMessage}`);
          updateProgress(100, "✅ Succès!");
          hideLoading();
          showStatus(`🎉 Succès - ${errorMessage}`, "success");
        } else {
          updateProgress(100, "✅ Succès!");
          hideLoading();
          showStatus(`🎉 Succès - Votre demande a été traitée`, "success");
        }
      } else {
        // Error case - resultcode is not 0
        debugLog(`❌ ERREUR - L'API a renvoyé le code d'erreur: ${resultCode}`);
        let errorMsg = errorMessage || `Code ${resultCode}`;
        updateProgress(100, "❌ Erreur");
        hideLoading();
        showStatus(`❌ ${errorMsg}`, "error");
        return;
      }
    } else if (parsed.error === 0) {
      // No resultCode extracted, but parsed.error is 0, treat as success
      debugLog(`✅ SUCCÈS - Demande traitée avec succès`);
      updateProgress(100, "✅ Succès!");
      hideLoading();
      showStatus(`🎉 Succès - Votre demande a été traitée`, "success");
    } else if (parsed.error !== 0) {
      // API returned error code at top level
      debugLog(`❌ ERREUR - L'API a renvoyé le code d'erreur: ${parsed.error}`);
      updateProgress(100, "❌ Erreur");
      hideLoading();
      showStatus(`❌ Erreur serveur (${parsed.error})`, "error");
      return;
    } else if (errorMessage) {
      // Error indicated by errormessage field when resultcode is not extracted
      debugLog(`❌ ERREUR - Le serveur a renvoyé un message d'erreur: ${errorMessage}`);
      updateProgress(100, "❌ Erreur");
      hideLoading();
      showStatus(`❌ ${errorMessage}`, "error");
      return;
    } else {
      debugLog(`⚠️ Impossible d'extraire le résultat ou l'événement de la réponse`);
      debugLog(`📊 Réponse complète: ${JSON.stringify(parsed)}`);
      showStatus(`✅ Demande traitée`, "success");
      hideLoading();
    }

  } catch (err) {
    debugLog("❌ Erreur de récupération: " + err.message);
    updateProgress(100, "❌ Erreur");
    hideLoading();
    showStatus("❌ Erreur de communication", "error");
  }
}

/* ======================
   ACTIONS
====================== */

async function performAction(actionType) {
  if (!cachedToken) {
    debugLog("❌ No authentication token - cannot perform action");
    showStatus("❌ Token non disponible", "error");
    return;
  }

  showLoading();
  updateProgress(10, "Préparation de l'action...");

  try {
    const user = Office.context.mailbox.userProfile.emailAddress;
    const item = Office.context.mailbox.item;

    // Build REST API payload for WebService/Execute
    const actionPayload = {
      action: "WEB_SERVICE_INFINITY",
      access_token: cachedToken,
      param: JSON.stringify({
        action: { swinfinity: "dv_creation_evt" },
        data: {
          evenement: {
            type: actionType,
            utilisateur: user,
            tiers: item.from?.emailAddress || "",
            lib: item.subject || "",
            pj: cachedPayload?.evenement?.pj || "",
            evt_lie: cachedPayload?.evenement?.evt_lie || ""
          }
        }
      })
    };

    debugLog(`📋 Action payload prepared:`);
    debugLog(`  - Action type: ${actionType}`);
    debugLog(`  - User: ${user}`);
    debugLog(`  - Token: ${cachedToken.substring(0, 20)}...`);

    updateProgress(50, "Envoi de l'action...");
    debugLog(`🚀 Sending action to WebService/Execute...`);

    const response = await fetch("https://addin-divalto.divy-si.fr/ASFLUID/outlook/proxy/proxy.php", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(actionPayload)
    });

    updateProgress(80, "Traitement de la réponse...");

    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const result = await response.json();
    debugLog(`✅ Action response: ${JSON.stringify(result).substring(0, 200)}`);

    if (result.ok === true || result.token) {
      debugLog(`✅ Action completed successfully`);
      updateProgress(100, "✅ Succès!");
      hideLoading();
      showStatus(`🎉 Action réussie`, "success");
    } else {
      debugLog(`⚠️ Unexpected response: ${JSON.stringify(result)}`);
      updateProgress(100, "✅ Complété");
      hideLoading();
      showStatus(`✅ Action traitée`, "success");
    }

  } catch (error) {
    debugLog(`❌ Action error: ${error.message}`);
    updateProgress(100, "❌ Erreur");
    hideLoading();
    showStatus(`❌ Erreur lors de l'action: ${error.message}`, "error");
  }
}

