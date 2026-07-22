/**
 * Integrated Campaign Intelligence — serverless generation endpoint
 * Developed by Carlos Eduardo - https://www.linkedin.com/in/carloseduardovf/
 *
 * The ANTHROPIC_API_KEY lives here (server-side) and never reaches the browser.
 * Set it in Vercel: Project → Settings → Environment Variables → ANTHROPIC_API_KEY
 */

const MODEL = "claude-sonnet-4-6";
const MAX_TOKENS = 4000;
const MAX_RETRIES = 3;

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

export default async function handler(req, res) {
  if (req.method !== "POST") {
    return res.status(405).json({ error: "Method not allowed. Use POST." });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return res.status(500).json({
      error: "Server is missing ANTHROPIC_API_KEY. Add it in Vercel → Settings → Environment Variables, then redeploy.",
    });
  }

  let system, user;
  try {
    const body = typeof req.body === "string" ? JSON.parse(req.body) : req.body;
    system = body && body.system;
    user = body && body.user;
  } catch (e) {
    return res.status(400).json({ error: "Invalid JSON body." });
  }

  if (!system || !user) {
    return res.status(400).json({ error: "Missing 'system' or 'user' in request body." });
  }

  let lastError = "Unknown error.";

  for (let attempt = 0; attempt < MAX_RETRIES; attempt++) {
    try {
      const controller = new AbortController();
      const timeout = setTimeout(() => controller.abort(), 50000);

      const upstream = await fetch("https://api.anthropic.com/v1/messages", {
        method: "POST",
        signal: controller.signal,
        headers: {
          "Content-Type": "application/json",
          "x-api-key": apiKey,
          "anthropic-version": "2023-06-01",
        },
        body: JSON.stringify({
          model: MODEL,
          max_tokens: MAX_TOKENS,
          system,
          messages: [{ role: "user", content: user }],
        }),
      });

      clearTimeout(timeout);

      // Transient upstream conditions: back off and retry.
      if (upstream.status === 429 || upstream.status === 529 || upstream.status >= 500) {
        lastError = "The model is busy (" + upstream.status + "). Please try again.";
        await sleep(700 * Math.pow(2, attempt));
        continue;
      }

      if (!upstream.ok) {
        let detail = "";
        try {
          const errJson = await upstream.json();
          detail = (errJson && errJson.error && errJson.error.message) || "";
        } catch (e) { /* no JSON body */ }

        if (upstream.status === 400 && /credit balance/i.test(detail)) {
          return res.status(402).json({
            error: "Anthropic credit balance is too low. Top up at console.anthropic.com → Billing.",
          });
        }
        if (upstream.status === 401) {
          return res.status(401).json({ error: "Invalid ANTHROPIC_API_KEY. Check the value in Vercel and redeploy." });
        }
        return res.status(upstream.status).json({ error: detail || "Upstream error (" + upstream.status + ")." });
      }

      const data = await upstream.json();
      const text = (data.content || [])
        .filter((b) => b.type === "text")
        .map((b) => b.text)
        .join("\n");

      return res.status(200).json({ text });
    } catch (e) {
      lastError = e && e.name === "AbortError"
        ? "The request timed out. Try again."
        : (e && e.message) || "Network error.";
      await sleep(700 * Math.pow(2, attempt));
    }
  }

  return res.status(503).json({ error: lastError });
}
