// === N8N WEBHOOK URL ===
const N8N_WEBHOOK_URL =
  "https://gradhun.app.n8n.cloud/webhook/summarizer";

// === Initialize Add-in ===
Office.onReady(() => {
  if (Office.context.host === Office.HostType.Outlook) {
    const btn = document.getElementById("analyzeButton");
    if (btn) btn.onclick = analyzeEmail;
  }
});

// === MAIN ANALYZE FUNCTION ===
async function analyzeEmail() {
  const resultDiv = document.getElementById("result");
  resultDiv.textContent = "Analyzing email...";

  try {
    const item = Office.context.mailbox.item;

    // SUBJECT
    const subject = item.subject || "(no subject)";

    // BODY (HTML → text)
    const body = await getEmailBody(item);

    console.log("=== OUTLOOK EXTRACTED DATA ===");
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

    const rawText = await response.text();
    console.log("RAW n8n RESPONSE:", rawText);

    if (!response.ok)
      throw new Error("n8n error " + response.status + ": " + rawText);

    let data;
    try {
      data = JSON.parse(rawText);
    } catch {
      data = { task: rawText };
    }

    resultDiv.textContent = data.task || "No task returned.";
  } catch (err) {
    console.error("ANALYZE ERROR:", err);
    resultDiv.textContent = "Error: " + err.message;
  }
}

// === GET BODY AS HTML AND CLEAN TO TEXT ===
function getEmailBody(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(
      Office.CoercionType.Html,
      (asyncResult) => {
        if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
          return reject(asyncResult.error);
        }

        let html = asyncResult.value || "";

        // Remove scripts/styles
        html = html
          .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
          .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "");

        // Convert HTML → text
        let text = html
          .replace(/<br\s*\/?>/gi, "\n")
          .replace(/<\/p>/gi, "\n")
          .replace(/<\/div>/gi, "\n")
          .replace(/<[^>]+>/g, "")
          .replace(/\n\s*\n+/g, "\n\n")
          .trim();

        resolve(text);
      }
    );
  });
}
