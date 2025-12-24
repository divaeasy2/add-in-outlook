
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

const MAX_ATTACHMENT_SIZE = 300 * 1024; // 300 KB

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
   STATUS
====================== */

function showStatus(msg, type = "info") {
  const el = document.getElementById("status");
  el.className = `status ${type}`;
  el.innerText = msg;
  el.style.display = "block";
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
        pj: ""
      }
    };

    document.getElementById("btnSav").disabled = false;
    document.getElementById("btnComm").disabled = false;

    showStatus("📩 Email prêt pour traitement", "info");
  });
}

/* ======================
   ATTACHMENT (SMART)
====================== */

function getAttachmentBase64(att) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.getAttachmentContentAsync(att.id, (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        reject("Lecture PJ échouée");
        return;
      }

      if (
        res.value.format !==
        Office.MailboxEnums.AttachmentContentFormat.Base64
      ) {
        reject("Format PJ non supporté");
        return;
      }

      resolve(res.value.content);
    });
  });
}

async function getSmartAttachment() {
  const item = Office.context.mailbox.item;

  if (!item.attachments || item.attachments.length === 0) {
    return "";
  }

  const att = item.attachments[0];

  if (att.size > MAX_ATTACHMENT_SIZE) {
    showStatus("⚠️ Pièce jointe ignorée (trop volumineuse)", "info");
    return "";
  }

  showStatus("📎 Lecture de la pièce jointe...", "info");

  return await getAttachmentBase64(att);
}

/* ======================
   SEND
====================== */

async function send(type) {
  if (!cachedPayload) {
    showStatus("❌ Aucun email prêt", "error");
    return;
  }

  try {
    cachedPayload.evenement.type = type;
    cachedPayload.evenement.pj = await getSmartAttachment();

    showStatus("⏳ Test Envoi en cours...", "info");

    const res = await fetch(
      "https://maisondelarose.org/proxy/proxy.php",
      {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify(cachedPayload)
      }
    );

    const text = await res.text();
    if (!res.ok) throw new Error();

    const parsed = JSON.parse(text);
    const resultStr = parsed?.json?.result || "";

    const code = resultStr.match(/"resultcode"\s*:\s*"(\d+)"/)?.[1];
    const evt = resultStr.match(/"EvtNo"\s*:\s*"([^"]+)"/)?.[1]?.trim();
    const err =
      resultStr.match(/"errormessage"\s*:\s*"([^"]*)"/)?.[1] || "";

    if (code === "0") {
      showStatus(`✅ Succès — Code ${evt}`, "success");
    } else {
      showStatus(`❌ ${err || "Erreur inconnue"}`, "error");
    }

  } catch {
    showStatus("❌ Erreur de communication avec le serveur !!!!", "error");
  }
}

