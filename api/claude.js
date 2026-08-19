// Proxy serverless para a API da Anthropic.
//
// Antes, o index.html chamava https://api.anthropic.com/v1/messages diretamente
// do browser: sem cabeçalho x-api-key (não funcionava) e, mesmo com uma chave,
// exposta a qualquer pessoa que inspecionasse o código-fonte. Este endpoint
// resolve os dois problemas — a chave fica só aqui, no servidor.
//
// Deploy: importar este repositório no Vercel e definir a variável de ambiente
// ANTHROPIC_API_KEY (Project Settings → Environment Variables). Nenhuma outra
// configuração é necessária — o Vercel deteta a pasta /api automaticamente.

export default async function handler(req, res) {
  if (req.method !== "POST") {
    res.status(405).json({ error: "Método não permitido" });
    return;
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    res.status(500).json({ error: "ANTHROPIC_API_KEY não configurada no servidor" });
    return;
  }

  try {
    const upstream = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01"
      },
      body: JSON.stringify(req.body)
    });

    const data = await upstream.json();
    res.status(upstream.status).json(data);
  } catch (err) {
    res.status(502).json({ error: "Falha ao contactar a API da Anthropic" });
  }
}
