/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

// Office.onReady((info) => {
//   // Hide sideload message
//   const sideload = document.getElementById("sideload-msg");
//   if (sideload) sideload.style.display = "none";

//   // Show main app body
//   const appBody = document.getElementById("app-body");
//   if (appBody) appBody.style.display = "block";
// });

// export async function run() {
//   /**
//    * Insert your Outlook code here
//    */

//   const item = Office.context.mailbox.item;
//   let insertAt = document.getElementById("item-subject"); 
//   let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
//   insertAt.appendChild(label);
//   insertAt.appendChild(document.createElement("br"));
//   insertAt.appendChild(document.createTextNode(item.subject));
//   insertAt.appendChild(document.createElement("br"));
// }


Office.onReady(() => {
  const user = document.getElementById("user").innerText = Office.context.mailbox.userProfile.displayName
  const userEmail = document.getElementById("userEmail").innerText = Office.context.mailbox.userProfile.emailAddress
  // Hide sideload message
  const sideload = document.getElementById("sideload-msg");
  if (sideload) sideload.style.display = "none";

  // Show main app body
  const appBody = document.getElementById("app-body");
  if (appBody) appBody.style.display = "block";
  document.getElementById("btnTest").onclick = testEmail;
});

function testEmail() {
  const item = Office.context.mailbox.item;
  const subject = item.subject;
  const from = item.from?.emailAddress;

  const body = document.getElementById("body");
  if (body) body.style.display = "block";

  const resultJson = document.getElementById("resultJson");
  if (resultJson) resultJson.style.display = "block";

  const getCompanyName = () => {
    return from.split('@')[1].split('.')[0]
    // return from.slice(from.indexOf('@')+1, from.indexOf('.'))
  }

  item.body.getAsync(
    Office.CoercionType.Text,
    function (result){
      if(result.status===Office.AsyncResultStatus.Succeeded) {
        body.innerText = result.value

        const resultJSON = {
          evenement:{
            codeevt : 1102,
            tiers : item.sender.displayName,
            company : getCompanyName(),
            contact : from,
            lib: subject,
          }
        }

        fetch('https://remote.divy-si.fr:8443/DhsDivaltoServiceDivaApiRest/api/v1/Webhook/5DED7C6421BE4694A7D992BE08D93D2F0278797F',{
          method: 'POST',
          headers: {
            "Content-Type": "Application/json"
          },
          body: JSON.stringify(resultJSON)
        }
        )

        resultJson.innerText = JSON.stringify(resultJSON, null, 2)

      } else{
        body.innerText = "Cannot read the content"
      }
    }
  )

  const result = document.getElementById("result");
  if (result) result.style.display = "block";
  result.textContent =
    `Subject: ${subject}\nFrom: ${from}`;
}

