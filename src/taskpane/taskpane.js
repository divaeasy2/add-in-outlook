/* global Office */

let cachedPayload = null;
const MAX_EMAIL_SIZE = 10 * 1024 * 1024; // 10 MB
let statusTimeoutId = null;
let allEvents = []; // Store all events for filtering

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
  
  // Confirmation modal buttons
  document.getElementById("btnConfirmLinked").onclick = () => sendWithLinkedEvent();
  document.getElementById("btnCancelLinked").onclick = () => returnFromConfirmation();

  prepareEmail();
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
          debugLog("✅ Token obtained");
          resolve({
            success: true,
            token: result.value,
            messageId: item.itemId,
            emailAddress: Office.context.mailbox.userProfile.emailAddress
          });
        } else {
          debugLog("⚠️ Token error: " + result.error.message);
          resolve({
            success: false,
            messageId: item.itemId
          });
        }
      });
    } else {
      debugLog("⚠️ Token error: t.getCallbackTokenAsync is not a function");
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
        debugLog("✅ Email body retrieved (HTML)");
        bodyLoaded = true;
      } else {
        // Fallback to plain text
        item.body.getAsync(Office.CoercionType.Text, (textRes) => {
          if (textRes.status === Office.AsyncResultStatus.Succeeded) {
            emailData.body = textRes.value;
            emailData.bodyType = "text";
            debugLog("✅ Email body retrieved (Text - fallback)");
          } else {
            debugLog("⚠️ Body retrieval failed");
            emailData.body = "[Unable to retrieve body]";
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
      debugLog(`📎 Found ${item.attachments.length} attachment(s)`);
      let loadedCount = 0;
      
      item.attachments.forEach((att, idx) => {
        try {
          // Get attachment data
          item.getAttachmentContentAsync(att.id, (result) => {
            debugLog(`📥 Processing attachment: ${att.name}`);
            
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
                debugLog(`❌ Failed to encode attachment ${att.name}: ${encodeErr.message}`);
              }
            } else {
              debugLog(`⚠️ Failed to load attachment: ${att.name}`);
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
}

function hideSavOptions() {
  document.getElementById("savOptionsModal").style.display = "none";
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

  const payload = {
    evenement: {
      utilisateur: cachedPayload.evenement.utilisateur,
      tiers: cachedPayload.evenement.tiers
    }
  };

  const res = await fetch("https://maisondelarose.org/proxy/proxy_child.php", {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=UTF-8" },
    body: JSON.stringify(payload)
  }).then(r => r.text()).catch(() => null);

  if (!res) return showStatus("❌ Erreur réseau", "error");

  const parsed = parseWeirdApiResponse(res);
  const popup = document.getElementById("childPopup");
  const select = document.getElementById("childSelect");
  const evtCount = document.getElementById("evtCount");
  const searchInput = document.getElementById("eventSearchInput");

  if (!parsed.ok) {
    popup.style.display = "none";
    return showStatus("🔴 " + parsed.error, "error");
  }

  // Store events for filtering
  allEvents = parsed.events;

  // Hide options modal and show child events popup
  document.getElementById("savOptionsModal").style.display = "none";
  popup.style.display = "block";
  
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
  
  showStatus(`🟢 ${parsed.count} évènements récupérés`, "success");
}

function filterEvents(searchTerm) {
  const select = document.getElementById("childSelect");
  const searchValue = searchTerm.toLowerCase().trim();
  
  // Clear current options (except placeh older)
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
  showStatus("⏳ Retour à la sélection...", "info");
}

function returnFromLinkedEvents() {
  document.getElementById("childPopup").style.display = "none";
  cachedPayload.evenement.evt_lie = "";
  
  // Show options modal again
  document.getElementById("savOptionsModal").style.display = "block";
  showStatus("");
  document.getElementById("status").style.display = "none";
}


/* ======================
   SEND
====================== */

async function send(type) {
  document.getElementById("status").style.display = "block";
  if (!cachedPayload) return showStatus("⚠️ Aucun email prêt", "error");

  showLoading(); // Start loading animation
  updateProgress(10, "Initialisation...");
  try {
    const item = Office.context.mailbox.item;

    showStatus("⌛ Récupération des données...", "info");

    // Step 1: Get Office API messageId
    debugLog("Step 1: Getting callback token...");
    updateProgress(15, "Récupération du token...");
    const tokenData = await getEmailWithOfficeApi();
    
    if (tokenData.success) {
      debugLog(`✅ Token obtained for user: ${tokenData.emailAddress}`);
      debugLog(`Message ID: ${tokenData.messageId}`);
      cachedPayload.evenement.messageId = tokenData.messageId;
    } else {
      debugLog(`⚠️ No token, using Office API fallback`);
      cachedPayload.evenement.messageId = item.itemId;
    }

    // Step 2: Get email content with attachments
    debugLog("Step 2: Getting email content...");
    updateProgress(25, "Lecture du contenu de l'email et des pièces jointes...");
    const emailContent = await getEmailContent();
    
    debugLog(`✅ Email from: ${emailContent.from}`);
    debugLog(`📧 Subject: ${emailContent.subject}`);
    debugLog(`📄 Body type: ${emailContent.bodyType} (${emailContent.bodyType === "html" ? "images/formatting preserved" : "text only"})`);
    if (emailContent.attachments && emailContent.attachments.length > 0) {
      debugLog(`📎 ${emailContent.attachments.length} attachment(s):`);
      emailContent.attachments.forEach(att => {
        debugLog(`   - ${att.name} (${att.contentType})`);
      });
    } else {
      debugLog(`📎 No attachments`);
    }
    updateProgress(50, "Encodage de l'email en base64...");

    // Step 3: Build email in base64
    showStatus("⌛ Encodage du message...", "info");
    const emailBase64 = buildEmailBase64(item, emailContent);
    
    if (!emailBase64) {
      debugLog("⚠️ Email too large, will send without attachment");
      cachedPayload.evenement.pj = "";
    } else {
      debugLog(`✅ Email encoded: ${emailBase64.length} bytes`);
      const sizeMB = (emailBase64.length / 1024 / 1024).toFixed(2);
      if (emailBase64.length > 1000000) {
        debugLog(`⚠️ Large payload: ${sizeMB}MB - may take longer to process`);
      }
      cachedPayload.evenement.pj = emailBase64;
    }

    cachedPayload.evenement.type = type;
    cachedPayload.evenement.evt_lie = cachedPayload.evenement.evt_lie || "";

    // Step 4: Send to proxy
    debugLog(`📤 Sending to Divalto (type: ${type})...`);
    updateProgress(70, "Préparation du payload...");
    showStatus("🚀 Envoi... (cela peut prendre du temps pour les gros fichiers)", "info");

    // Measure JSON stringify performance
    const startStringify = performance.now();
    const jsonPayload = JSON.stringify(cachedPayload);
    const stringifyTime = performance.now() - startStringify;
    const payloadSizeKB = (jsonPayload.length / 1024).toFixed(2);
    debugLog(`⏱️ JSON.stringify took ${stringifyTime.toFixed(2)}ms for ${payloadSizeKB}KB`);
    
    updateProgress(75, `Transmission vers le serveur (${payloadSizeKB}KB)...`);
    
    // Measure fetch performance - split into request and response times
    const startFetch = performance.now();
    let requestTime = 0;
    
    const res = await fetch("https://maisondelarose.org/proxy/proxy.php", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: jsonPayload
    }).then(response => {
      requestTime = performance.now() - startFetch;
      debugLog(`⏱️ Request transmission took ${requestTime.toFixed(2)}ms`);
      return response;
    });
    
    updateProgress(82, "Attente de la réponse du serveur...");
    const startResponse = performance.now();

    // Step 5: Parse response
    updateProgress(85, "Traitement par le serveur...");
    debugLog(`Response status: ${res.status}`);
    const text = await res.text();
    const responseTime = performance.now() - startResponse;
    debugLog(`⏱️ Response received in ${responseTime.toFixed(2)}ms`);
    const totalFetchTime = performance.now() - startFetch;
    debugLog(`⏱️ Total fetch (request + response) took ${totalFetchTime.toFixed(2)}ms`);
    debugLog(`Response text length: ${text.length} bytes`);
    debugLog(`First 500 chars: ${text.substring(0, 500)}`);
    
    updateProgress(95, "Finalisation...");
    const startFinalization = performance.now();
    
    let parsed;
    try {
      const parseStart = performance.now();
      parsed = JSON.parse(text);
      const parseTime = performance.now() - parseStart;
      debugLog(`⏱️ JSON.parse took ${parseTime.toFixed(2)}ms`);
      debugLog(`✅ JSON parsed successfully`);
    } catch (e) {
      debugLog(`❌ JSON parse error: ${e.message}`);
      showStatus("⚠️ Format de réponse inattendu", "warning");
      return;
    }

    debugLog(`Full response: ${JSON.stringify(parsed)}`);
    
    // Parse Divalto response structure
    let code = null;
    let evt = null;
    
    try {
      let resultStr = parsed?.json?.result || "";
      debugLog(`Raw result string: ${resultStr}`);
      
      if (!resultStr) {
        throw new Error("No result string found");
      }
      
      // Unescape the string
      let unescaped = resultStr.replace(/\\"/g, '"').replace(/\\\\/g, '\\');
      debugLog(`Unescaped: ${unescaped}`);
      
      // Extract using regex
      const codeMatch = unescaped.match(/"resultcode"\s*:\s*"(\d+)"/);
      const evtMatch = unescaped.match(/"EvtNo"\s*:\s*"([^"]+)"/);
      
      code = codeMatch ? codeMatch[1] : null;
      evt = evtMatch ? evtMatch[1].trim() : null;
      
      debugLog(`Extracted - Code: ${code}, Event: ${evt}`);
      
    } catch (e) {
      debugLog(`❌ Parse error: ${e.message}`);
    }

    const finalizationTime = performance.now() - startFinalization;
    debugLog(`⏱️ Finalisation took ${finalizationTime.toFixed(2)}ms`);

    if (code === "0" && evt) {
      debugLog(`✅ SUCCESS - Event: ${evt}`);
      updateProgress(100, "✅ Succès!");
      hideLoading();
      showStatus(`🎉 Succès — Évènement: ${evt}`, "success");
    } else if (code && code !== "0") {
      debugLog(`❌ API Error - Code: ${code}`);
      updateProgress(100, "❌ Erreur");
      hideLoading();
      showStatus(`❌ Erreur API (Code: ${code})`, "error");
    } else if (res.status === 502) {
      // Handle 502 Bad Gateway (timeout from proxy)
      updateProgress(100, "❌ Timeout");
      hideLoading();
      const errorDetails = parsed?.details || "";
      if (errorDetails.includes("timeout")) {
        debugLog(`⚠️ TIMEOUT: Divalto API took too long to process (large file)`);
        showStatus(`⚠️ Timeout - le fichier est trop volumineux pour Divalto`, "warning");
      } else {
        debugLog(`❌ Erreur 502 - Proxy error: ${errorDetails}`);
        showStatus(`❌ Erreur 502 - Problème de communication`, "error");
      }
    } else {
      debugLog(`⚠️ Could not extract code or event`);
      debugLog(`Response: ${JSON.stringify(parsed)}`);
      updateProgress(100, "⚠️ Réponse invalide");
      hideLoading();
      showStatus(`⚠️ Réponse inattendue du serveur`, "warning");
    }

  } catch (err) {
    debugLog("❌ Fetch error: " + err.message);
    updateProgress(100, "❌ Erreur");
    hideLoading();
    showStatus("❌ Erreur de communication", "error");
  }
}


