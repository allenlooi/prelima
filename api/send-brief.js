function escapeHtml(s) {
  return String(s || "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const apiKey = process.env.RESEND_API_KEY;
  if (!apiKey) {
    // Email isn't set up yet — no-op rather than break the intake flow.
    res.status(200).json({ skipped: true, reason: "RESEND_API_KEY not configured" });
    return;
  }

  const { to, name, projectName, freelancer, brief } = req.body || {};
  if (!to || !brief) {
    res.status(400).json({ error: "to and brief are required" });
    return;
  }

  const from = process.env.RESEND_FROM || "Prelima <onboarding@resend.dev>";
  const html = `
    <div style="font-family: sans-serif; max-width: 600px; margin: 0 auto; color: #17171C;">
      <h2 style="margin-bottom: 4px;">Your brief for ${escapeHtml(projectName || "your project")}</h2>
      <p style="color: #6E6E78; margin-top: 0;">Sent to ${escapeHtml(freelancer || "your freelancer")} — here's a copy for your records${name ? `, ${escapeHtml(name)}` : ""}.</p>
      <div style="white-space: pre-wrap; line-height: 1.6; border-top: 1px solid #E8E8E2; padding-top: 16px; margin-top: 16px;">${escapeHtml(brief)}</div>
    </div>
  `;

  try {
    const upstream = await fetch("https://api.resend.com/emails", {
      method: "POST",
      headers: { "Content-Type": "application/json", "Authorization": `Bearer ${apiKey}` },
      body: JSON.stringify({
        from,
        to: [to],
        subject: `Your brief — ${projectName || "New project"}`,
        html,
      }),
    });
    const data = await upstream.json();
    res.status(upstream.status).json(data);
  } catch {
    res.status(502).json({ error: "Failed to reach email service" });
  }
}
