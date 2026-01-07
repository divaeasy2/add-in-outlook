/* global Office */

let cachedPayload = null;
const MAX_EMAIL_SIZE = 500 * 1024; // 500 KB
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
        pj: "",
        evt_lie: ""
      }
    };

    document.getElementById("btnSav").disabled = false;
    document.getElementById("btnComm").disabled = false;

    showStatus("✅ Email prêt pour traitement", "info");
  });
}

/* ======================
   BUILD EMAIL (.eml)
====================== */

function buildEmailBase64(item, bodyText) {
  const bodyBase64 = btoa(unescape(encodeURIComponent(bodyText)));

  const eml =
`From: ${item.from?.emailAddress || ""}
To: ${Office.context.mailbox.userProfile.emailAddress}
Subject: ${item.subject || ""}
Date: ${new Date().toUTCString()}
MIME-Version: 1.0
Content-Type: text/plain; charset=UTF-8
Content-Transfer-Encoding: base64

${bodyBase64}`;

  const size = new Blob([eml]).size;
  if (size > MAX_EMAIL_SIZE) return null;

  return btoa(unescape(encodeURIComponent(eml)));
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

  try {
    const item = Office.context.mailbox.item;

    showStatus("⌛ Lecture...", "info");

    const body = await new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, r => {
        r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject();
      });
    });

    cachedPayload.evenement.type = type;
    cachedPayload.evenement.evt_lie = cachedPayload.evenement.evt_lie || "";

    const emailBase64 = buildEmailBase64(item, body);
    cachedPayload.evenement.pj = emailBase64 || "";

    showStatus("🚀 Envoi...", "info");

    const res = await fetch("https://maisondelarose.org/proxy/proxy.php", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(cachedPayload)
    });

    const text = await res.text();
    const parsed = JSON.parse(text);
    const resultStr = parsed?.json?.result || "";

    const code = resultStr.match(/"resultcode"\s*:\s*"(\d+)"/)?.[1];
    const evt = resultStr.match(/"EvtNo"\s*:\s*"([^"]+)"/)?.[1]?.trim();

    if (code === "0") {
      showStatus(`🎉 SUCCESS — Code ${evt}`, "success");
    } else {
      showStatus(`❌ Erreur`, "error");
    }

  } catch {
    showStatus("❌ Erreur de communication", "error");
  }
}




