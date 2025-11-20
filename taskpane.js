// taskpane.js



// n8n PRODUCTION webhook URL

const N8N_WEBHOOK_URL =

  "https://gradhun.app.n8n.cloud/webhook/summarizer";



Office.onReady(function () {

  if (Office.context.host === Office.HostType.Outlook) {

    const btn = document.getElementById("analyzeButton");

    if (btn) {

      btn.onclick = analyzeEmail;

    }

  }

});



async function analyzeEmail() {

  const resultDiv = document.getElementById("result");

  const btn = document.getElementById("analyzeButton");



  resultDiv.textContent = "Analyzing email...";

  if (btn) btn.disabled = true;



  try {

    const item = Office.context.mailbox.item;



    const subject = item.subject || "(no subject)";

    const body = await getBodyAsync(item);



    const task = await callN8N(subject, body);



    resultDiv.textContent = task;

  } catch (err) {

    console.error("analyzeEmail error:", err);

    resultDiv.textContent = "Error: " + (err.message || String(err));

  } finally {

    if (btn) btn.disabled = false;

  }

}



function getBodyAsync(item) {

  return new Promise((resolve, reject) => {

    item.body.getAsync(Office.CoercionType.Text, (asyncResult) => {

      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {

        resolve(asyncResult.value || "");

      } else {

        reject(asyncResult.error);

      }

    });

  });

}



async function callN8N(subject, body) {

  const payload = { subject, body };



  const response = await fetch(N8N_WEBHOOK_URL, {

    method: "POST",

    headers: {

      "Content-Type": "application/json"

      // add auth header here if you secure the webhook

      // "x-api-key": "YOUR_KEY"

    },

    body: JSON.stringify(payload)

  });



  const text = await response.text();

  console.log("n8n status:", response.status);

  console.log("n8n raw response:", text);



  if (!response.ok) {

    throw new Error("n8n " + response.status + ": " + text);

  }



  try {

    const data = JSON.parse(text);

    return data.task || "No task returned.";

  } catch (e) {

    return text || "No response from n8n.";

  }

}


