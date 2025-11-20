const N8N_WEBHOOK_URL =
  "https://gradhun.app.n8n.cloud/webhook/summarizer";

Office.onReady(() => {
  if (Office.context.host === Office.HostType.Outlook) {
    const btn = document.getElementById("analyzeButton");
    if (btn) btn.onclick = analyzeEmail;
  }
});

async function analyzeEmail() {
  const result = document.getElementById("result");
  result.textContent = "Analyzing email...";

  try {
    const item = Office.context.mailbox.item;

    // SUBJECT
    const subject = item.subject || "(no subject)";

    // BODY
    const body = await getEmailBody(item);

    console.log("SUBJECT:", subject);
    console.log("BODY:", body);

    // SEND TO N8N
    const payload = { subject, body };

    const response = await fetch(N8N_WEBHOOK_URL, {
      method: "POST",
      headers: {
        "Content-Type": "application/json"
      },
      body: JSON.stringify(payload)
    });

    const text = await response.text();
    console.log("RAW n8n response:", text);

    let data;
    try {
      data = JSON.parse(text);
    } catch {
      data = { task: text };
    }

    result.textContent = data.task || "No task returned.";
  } catch (err) {
    console.error("Error:", err);
    result.textContent = "Error: " + err.message;
  }
}

function getEmailBody(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(
      Office.CoercionType.Text,
      (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
          resolve(asyncResult.value || "");
        } else {
          reject(asyncResult.error);
        }
      }
    );
  });
}
