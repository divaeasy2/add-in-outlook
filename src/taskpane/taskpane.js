
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
   GLOBAL STATE
====================== */

let cachedPayload = null;

/* ======================
   OFFICE READY
====================== */

Office.onReady(() => {
  // User info
  document.getElementById("user").innerText =
    Office.context.mailbox.userProfile.displayName;

  document.getElementById("userEmail").innerText =
    Office.context.mailbox.userProfile.emailAddress;

  const sideload = document.getElementById("sideload-msg");
  if (sideload) sideload.style.display = "none";

  const appBody = document.getElementById("app-body");
  if (appBody) appBody.style.display = "block";

  document.getElementById("btnDetails").onclick = displayEmailDetails;
  document.getElementById("btnSav").onclick = () => sendToApi("1");
  document.getElementById("btnComm").onclick = () => sendToApi("2");

  document.getElementById("btnSav").disabled = true;
  document.getElementById("btnComm").disabled = true;
});

/* ======================
   UI HELPERS
====================== */

function showStatus(msg, isError = false) {
  const el = document.getElementById("status");
  el.innerText = msg;
  el.style.color = isError ? "red" : "green";
}

function showApiResponse(text) {
  document.getElementById("apiResponse").innerText = text;
}

/* ======================
   DISPLAY EMAIL (NO FETCH)
====================== */

function displayEmailDetails() {

  const body = document.getElementById("body");
  if (body) body.style.display = "block";

  const myresult = document.getElementById("result");
  if (myresult) myresult.style.display = "block";

  const resultJson = document.getElementById("resultJson");
  if (resultJson) resultJson.style.display = "block";

  const item = Office.context.mailbox.item;

  const subject = item.subject || "";
  const from = item.from?.emailAddress || "";
  const user = Office.context.mailbox.userProfile.emailAddress;

  item.body.getAsync(Office.CoercionType.Text, (result) => {
    if (result.status !== Office.AsyncResultStatus.Succeeded) {
      showStatus("❌ Impossible de lire l’email", true);
      return;
    }

    body.innerText = result.value;
    myresult.innerText =
      `Subject: ${subject}\nFrom: ${from}`;

    currentPayload = {
      evenement: {
        type: "",
        utilisateur: user,
        tiers: from,
        lib: subject,
        pj: ""
      }
    };

    resultJson.innerText =
      JSON.stringify(currentPayload, null, 2);

    document.getElementById("btnSav").disabled = false;
    document.getElementById("btnComm").disabled = false;

    showStatus("📩 Détails affichés");
  });
}


/* ======================
   SEND TO API (SAV / COMM)
====================== */

function sendToApi(type) {
  if (!currentPayload) {
    showStatus("❌ Aucun email chargé", true);
    return;
  }

  currentPayload.evenement.type = type;

  document.getElementById("resultJson").innerText =
    JSON.stringify(currentPayload, null, 2);

  showStatus("⏳ Envoi vers l’API...");

  callApiSafe(currentPayload);
}

/* ======================
   SAFE FETCH (PROXY)
====================== */

function callApiSafe(payload) {
  fetch("https://maisondelarose.org/proxy/proxy.php", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  })
    .then(async (res) => {
      const text = await res.text();
      showApiResponse(text);

      if (!res.ok) {
        showStatus(`❌ HTTP ${res.status}`, true);
        return;
      }

      try {
        const parsed = JSON.parse(text);

        document.getElementById("apiResponse").innerText =
          JSON.stringify(parsed, null, 2);

        if (!parsed.json || !parsed.json.result) {
          showStatus("❌ Structure API inattendue", true);
          return;
        }

        const resultStr = parsed.json.result;

        const codeMatch = resultStr.match(/"resultcode"\s*:\s*"(\d+)"/);
        const evtMatch = resultStr.match(/"EvtNo"\s*:\s*"([^"]+)"/);
        const errMatch = resultStr.match(/"errormessage"\s*:\s*"([^"]*)"/);

        const resultcode = codeMatch ? codeMatch[1] : null;
        const evtNo = evtMatch ? evtMatch[1].trim() : null;
        const errorMsg = errMatch ? errMatch[1] : null;

        if (resultcode === "0") {
          showStatus(`✅ Succès — EVTCODE : ${evtNo}`);
        } else {
          showStatus(
            `❌ Erreur API : ${errorMsg || "Erreur inconnue"}`,
            true
          );
        }

      } catch (e) {
        showStatus("⚠️ Erreur JS lors du parsing", true);
        showApiResponse(e.toString());
      }
    })
    .catch((err) => {
      showStatus("❌ Fetch bloqué / erreur réseau", true);
      showApiResponse(err.toString());
    });
}
