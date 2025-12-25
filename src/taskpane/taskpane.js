
// Office.onReady(() => {
//   const user = document.getElementById("user").innerText = Office.context.mailbox.userProfile.displayName
//   const userEmail = document.getElementById("userEmail").innerText = Office.context.mailbox.userProfile.emailAddress
//   // Hide sideload message
//   const sideload = document.getElementById("sideload-msg");
//   if (sideload) sideload.style.display = "none";

//   // Show main app body
//   const appBody = document.getElementById("app-body");
//   if (appBody) appBody.style.display = "block";
//   document.getElementById("btnTest").onclick = testEmail;
// });

// function testEmail() {
//   const item = Office.context.mailbox.item;
//   const subject = item.subject;
//   const from = item.from?.emailAddress;

//   const body = document.getElementById("body");
//   if (body) body.style.display = "block";

//   const resultJson = document.getElementById("resultJson");
//   if (resultJson) resultJson.style.display = "block";

//   const getCompanyName = () => {
//     return from.split('@')[1].split('.')[0]
//     // return from.slice(from.indexOf('@')+1, from.indexOf('.'))
//   }

//   item.body.getAsync(
//     Office.CoercionType.Text,
//     function (result){
//       if(result.status===Office.AsyncResultStatus.Succeeded) {
//         body.innerText = result.value

//         const resultJSON = {
//           evenement:{
//             codeevt : 1102,
//             tiers : item.sender.displayName,
//             company : getCompanyName(),
//             contact : from,
//             lib: subject,
//           }
//         }

//      fetch('https://remote.divy-si.fr:8443/DhsDivaltoServiceDivaApiRest/api/v1/Webhook/5DED7C6421BE4694A7D992BE08D93D2F0278797F',{
//           method: 'POST',
//           headers: {
//             "Content-Type": "Application/json"
//           },
//           body: JSON.stringify(resultJSON)
//         }
//         )

//         resultJson.innerText = JSON.stringify(resultJSON, null, 2)

//       } else{
//         body.innerText = "Cannot read the content"
//       }
//     }
//   )

//   const result = document.getElementById("result");
//   if (result) result.style.display = "block";
//   result.textContent =
//     `Subject: ${subject}\nFrom: ${from}`;
// }


/* global Office */

let cachedPayload = null;
const MAX_EMAIL_SIZE = 500 * 1024; // 500 KB

Office.onReady(() => {
  document.getElementById("user").innerText =
    Office.context.mailbox.userProfile.displayName;

  document.getElementById("userEmail").innerText =
    Office.context.mailbox.userProfile.emailAddress;
    
  document.getElementById("btnSav").onclick = () => send("1"); // SAV = toujours type 1
  document.getElementById("btnComm").onclick = () => send("2");
  document.getElementById("btnDDP").onclick = () => send("3");
  document.getElementById("btnCDE").onclick = () => send("4");
  document.getElementById("btnDDI").onclick = () => send("5");
  document.getElementById("btnChild").onclick = loadChildEvents;

  document.getElementById("btnSav").disabled = true;
  document.getElementById("btnComm").disabled = true;

  prepareEmail();
});

/* ======================
   STATUS
====================== */

function showStatus(msg, type = "info") {
  const el = document.getElementById("status");
  el.className = `status ${type}`;
  el.innerText = msg;
  el.style.display = "block";
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
      showStatus("❌ Impossible de lire l’email", "error");
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

  // 🧽 تنظيف و تصحيح جميع الأنماط المعطوبة
  let cleaned = n2
    .replace(/\\"/g, '"')
    .replace(/"{/g, '{')
    .replace(/}"/g, '}')
    // ⬇️ إصلاح كل الأنواع ديال "result" المعطوبة
    .replace(/""result":/g, '"result":')
    .replace(/"result":"result":/g, '"result":')
    .replace(/"result":""/g, '"result":')
    // ⬇️ إصلاح duplication آخر محتمل
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

  if (!events) return { ok:false, error:"❌ Aucun événement trouvé", json:n3 };

  return { ok: true, count: events.length, events };
}


function debugLog(msg){
  const box = document.getElementById("debug");
  box.style.display = "block";
  box.innerText += "\n" + msg;
}

/* ======================
   LOAD CHILD EVENTS
====================== */

async function loadChildEvents() {
  if (!cachedPayload) return showStatus("⚠️ Aucun email prêt", "error");

  showStatus("⏳ Vérification...", "info");

  const payload = {
    evenement: {
      utilisateur: cachedPayload.evenement.utilisateur,
      tiers: cachedPayload.evenement.tiers
    }
  };

  const res = await fetch("https://maisondelarose.org/proxy/proxy_child.php", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  }).then(r => r.text()).catch(() => null);

  if (!res) return showStatus("❌ Erreur réseau", "error");

  const parsed = parseWeirdApiResponse(res);

  if (parsed.ok) {
    showStatus(`🟢 ${parsed.count} événements récupérés`, "success");

    const list = document.getElementById("childList");
    list.innerHTML = `<option value="">--- Choisissez ---</option>`;
    list.style.display = "block";

    parsed.events.forEach(evt => {
      const opt = document.createElement("option");
      opt.value = evt.evtNo;
      opt.innerText = `${evt.evtNo} - ${evt.lib || "(sans lib)"}`;
      list.appendChild(opt);
    });

    list.onchange = () => {
      cachedPayload.evenement.evt_lie = list.value || ""; // 👉 هادي المعتمدة دابا

      if (list.value) {
        showChildHint("⚠️ Événement lié sélectionné — cliquez sur Événement SAV pour l’envoyer");
      } else {
        showChildHint("");
      }

      showStatus(`📌 Sélectionné: ${list.value}`, "info");
    };

    return;
  }

  showStatus("🔴 " + parsed.error, "error");
}


/* ======================
   SEND
====================== */

async function send(type) {
  if (!cachedPayload) return showStatus("⚠️ Aucun email prêt", "error");

  if (type === "1") showChildHint(""); // logique SAV

  try {
    const item = Office.context.mailbox.item;

    showStatus("⌛ Lecture...", "info");

    const body = await new Promise((resolve, reject) => {
      item.body.getAsync(Office.CoercionType.Text, r => {
        r.status === Office.AsyncResultStatus.Succeeded ? resolve(r.value) : reject();
      });
    });

    cachedPayload.evenement.type = type;          // 👍 type = 1 pour SAV
    cachedPayload.evenement.evt_lie = cachedPayload.evenement.evt_lie || ""; // 👍 enfant si كاين

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




