const N8N_WEBHOOK_URL =
  "https://gradhun.app.n8n.cloud/webhook/summarizer";

Office.onReady(() => {
  if (Office.context.host === Office.HostType.Outlook) {
    document.getElementById("analyzeButton").onclick = analyzeEmail;
  }
});

async function analyzeEmail() {
  const res = document.getElementById("result");
  res.textContent = "Analyzing email...";

  try {
    const item = Office.context.mailbox.item;

    const subject = item.subject || "(no subject)";
    const body = await getEmailBody(item);

    console.log("SUBJECT:", subject);
    console.log("BODY:", body);

    const response = await fetch(N8N_WEBHOOK_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ subject, body })
    });

    const raw = await response.text();
    console.log("Raw n8n response:", raw);

    let data;
    try { data = JSON.parse(raw); }
    catch { data = { task: raw }; }

    res.textContent = data.task || "No task returned.";
  } catch (err) {
    res.textContent = "Error: " + err.message;
  }
}

function getEmailBody(item) {
  return new Promise((resolve, reject) => {
    item.body.getAsync(
      Office.CoercionType.Html,
      (r) => {
        if (r.status !== Office.AsyncResultStatus.Succeeded) {
          return reject(r.error);
        }

        let html = r.value || "";
        let text = html
          .replace(/<script[^>]*>[\s\S]*?<\/script>/gi, "")
          .replace(/<style[^>]*>[\s\S]*?<\/style>/gi, "")
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
