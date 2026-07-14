export const config = { runtime: "edge" };

// This route runs on Vercel's servers, never in the browser, so the API key
// stays private. The frontend calls /api/extract instead of api.anthropic.com.
export default async function handler(req: Request): Promise<Response> {
  if (req.method !== "POST") {
    return new Response(JSON.stringify({ error: "Method not allowed" }), { status: 405 });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return new Response(
      JSON.stringify({ error: "ANTHROPIC_API_KEY manquant sur le serveur (Vercel > Settings > Environment Variables)" }),
      { status: 500 }
    );
  }

  let body: { base64?: string; mimeType?: string; prompt?: string };
  try {
    body = await req.json();
  } catch {
    return new Response(JSON.stringify({ error: "Requête invalide" }), { status: 400 });
  }

  const { base64, mimeType, prompt } = body;
  if (!base64 || !mimeType || !prompt) {
    return new Response(JSON.stringify({ error: "Paramètres manquants" }), { status: 400 });
  }

  const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify({
      model: "claude-sonnet-5",
      max_tokens: 1500,
      messages: [
        {
          role: "user",
          content: [
            { type: "image", source: { type: "base64", media_type: mimeType, data: base64 } },
            { type: "text", text: prompt },
          ],
        },
      ],
    }),
  });

  const data = await anthropicRes.text();
  return new Response(data, {
    status: anthropicRes.status,
    headers: { "Content-Type": "application/json" },
  });
}
