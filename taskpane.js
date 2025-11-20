async function sendToN8N(subject, body) {
  const payload = {
    subject: subject || "(no subject)",
    body: body || ""
  };

  const response = await fetch("https://gradhun.app.n8n.cloud/webhook/summarizer", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify(payload)
  });

  const text = await response.text();

  let data;
  try { data = JSON.parse(text); }
  catch { data = { task: text }; }

  return data.task;
}
