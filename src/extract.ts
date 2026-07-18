// Fonction Node.js classique (pas "Edge") pour pouvoir monter la durée max à 60s :
// l'extraction d'un PDF de plusieurs pages peut prendre plus de temps qu'une simple image.
// La durée réelle autorisée est fixée dans vercel.json (functions -> api/extract.ts -> maxDuration).

export default async function handler(req: any, res: any) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Method not allowed" });
    return;
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    res.status(500).json({
      error: "ANTHROPIC_API_KEY manquant sur le serveur (Vercel > Settings > Environment Variables)",
    });
    return;
  }

  const { base64, mimeType, images, prompt } = req.body || {};
  if (!prompt) {
    res.status(400).json({ error: "Paramètres manquants" });
    return;
  }
  if ((base64 && !mimeType) || (!base64 && mimeType)) {
    res.status(400).json({ error: "Paramètres incohérents" });
    return;
  }

  const content: any[] = [];
  if (Array.isArray(images) && images.length > 0) {
    for (const img of images) {
      if (!img?.base64 || !img?.mimeType) continue;
      content.push({ type: "image", source: { type: "base64", media_type: img.mimeType, data: img.base64 } });
    }
  } else if (base64 && mimeType) {
    const isPdf = mimeType === "application/pdf";
    content.push(
      isPdf
        ? { type: "document", source: { type: "base64", media_type: "application/pdf", data: base64 } }
        : { type: "image", source: { type: "base64", media_type: mimeType, data: base64 } }
    );
  }
  content.push({ type: "text", text: prompt });

  try {
    const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      body: JSON.stringify({
        model: "claude-sonnet-5",
        max_tokens: 6000,
        messages: [
          {
            role: "user",
            content,
          },
        ],
      }),
    });

    const data = await anthropicRes.text();
    res.status(anthropicRes.status);
    res.setHeader("Content-Type", "application/json");
    res.send(data);
  } catch (err: any) {
    res.status(502).json({ error: err?.message || "Erreur lors de l'appel à l'API Anthropic" });
  }
}
