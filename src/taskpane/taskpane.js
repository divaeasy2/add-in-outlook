
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

/* ======================
   GLOBAL
====================== */

let cachedPayload = null;

/* ======================
   OFFICE READY
====================== */

Office.onReady(() => {
  document.getElementById("user").innerText =
    Office.context.mailbox.userProfile.displayName;

  document.getElementById("userEmail").innerText =
    Office.context.mailbox.userProfile.emailAddress;

  document.getElementById("btnSav").onclick = () => send("1");
  document.getElementById("btnComm").onclick = () => send("2");

  document.getElementById("btnSav").disabled = true;
  document.getElementById("btnComm").disabled = true;

  prepareEmail();
});

/* ======================
   STATUS (NOTIFICATION)
====================== */

function showStatus(msg, type = "info") {
  const el = document.getElementById("status");
  el.className = `status ${type}`;
  el.innerText = msg;
  el.style.display = "block";
}


/* ======================
   PREPARE EMAIL (NO UI)
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
        pj: ""
      }
    };

    document.getElementById("btnSav").disabled = false;
    document.getElementById("btnComm").disabled = false;

    showStatus("📩 Email prêt pour traitement", "info");
  });
}

/* ======================
   SEND (SAV / COMM)
====================== */

function send(type) {
  if (!cachedPayload) {
    showStatus("❌ Aucun email prêt", "error");
    return;
  }

  cachedPayload.evenement.type = type;

  showStatus("⏳ Envoi en cours...", "info");

  fetch("https://maisondelarose.org/proxy/proxy.php", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(cachedPayload)
  })
    .then(async (res) => {
      const text = await res.text();
      if (!res.ok) throw new Error(`HTTP ${res.status}`);

      const parsed = JSON.parse(text);
      const resultStr = parsed?.json?.result || "";

      const code = resultStr.match(/"resultcode"\s*:\s*"(\d+)"/)?.[1];
      const evt = resultStr.match(/"EvtNo"\s*:\s*"([^"]+)"/)?.[1]?.trim();
      const err = resultStr.match(/"errormessage"\s*:\s*"([^"]*)"/)?.[1] || "";

      if (code === "0") {
        showStatus(`✅ Succès — Code ${evt}`, "success");
      } else {
        showStatus(`❌ ${err || "Erreur inconnue"}`, "error");
      }
    })
    .catch(() => {
      showStatus("❌ Erreur de communication avec le serveur", "error");
    });
}
