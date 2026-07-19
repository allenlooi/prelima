// Trim raw HTML down to readable text and cap the length, so we only ever send
// the model a small, controlled amount of the page (keeps token cost predictable).
function extractText(html) {
  return String(html || "")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<[^>]+>/g, " ")
    .replace(/&[a-z#0-9]+;/gi, " ")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 3000);
}

async function fetchSiteText(url) {
  let u = String(url || "").trim();
  if (!u) return "";
  if (!/^https?:\/\//i.test(u)) u = "https://" + u;
  const controller = new AbortController();
  const timer = setTimeout(() => controller.abort(), 6000);
  try {
    const r = await fetch(u, {
      signal: controller.signal,
      headers: { "User-Agent": "Mozilla/5.0 (compatible; PrelimaBot/1.0)" },
    });
    const ct = r.headers.get("content-type") || "";
    if (!r.ok || !ct.includes("text/html")) return "";
    return extractText(await r.text());
  } catch {
    return "";
  } finally {
    clearTimeout(timer);
  }
}

// Cheaper model for extraction / short generation; Sonnet only when quality matters.
const MODEL_FAST = "claude-haiku-4-5-20251001";
const MODEL_QUALITY = "claude-sonnet-5";

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: "Server is missing ANTHROPIC_API_KEY" });
    return;
  }

  const { messages, fetchUrl, fast, maxTokens } = req.body || {};
  if (!Array.isArray(messages) || messages.length === 0) {
    res.status(400).json({ error: "messages is required" });
    return;
  }

  // If a URL is provided, fetch and trim it ourselves rather than paying for the
  // web-search tool — we already know the exact page we want.
  let outgoing = messages;
  if (fetchUrl) {
    const siteText = await fetchSiteText(fetchUrl);
    outgoing = messages.map((m, i) => {
      if (i !== messages.length - 1 || m.role !== "user") return m;
      const extra = siteText
        ? `\n\nWebsite text (already fetched for you — use only this):\n${siteText}`
        : `\n\n(Could not read the website — work only from what the user has typed.)`;
      return { ...m, content: `${m.content}${extra}` };
    });
  }

  const body = {
    model: fast ? MODEL_FAST : MODEL_QUALITY,
    max_tokens: Math.min(Math.max(Number(maxTokens) || 1000, 64), 1500),
    messages: outgoing,
  };

  try {
    const upstream = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify(body),
    });
    const data = await upstream.json();
    res.status(upstream.status).json(data);
  } catch {
    res.status(502).json({ error: "Failed to reach Anthropic API" });
  }
}
