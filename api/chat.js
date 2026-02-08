/**
 * Records chat API – conversation about the user's sessions and transactions.
 * Expects POST body: { messages: [{ role, content }], context: string }
 * Returns: { reply: string }
 * Set OPENAI_API_KEY in Vercel (or env) to enable.
 * CORS: set ALLOWED_ORIGINS (comma-separated), e.g. https://your-app.vercel.app,http://localhost,http://localhost:3000
 */

const getAllowedOrigins = () => {
  const raw = process.env.ALLOWED_ORIGINS || 'http://localhost,http://localhost:3000';
  return raw.split(',').map((o) => o.trim()).filter(Boolean);
};

export default async function handler(req, res) {
  const allowed = getAllowedOrigins();
  const origin = req.headers.origin;
  if (origin && allowed.includes(origin)) {
    res.setHeader('Access-Control-Allow-Origin', origin);
  }
  res.setHeader('Access-Control-Allow-Methods', 'POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type');

  if (req.method === 'OPTIONS') {
    return res.status(200).end();
  }

  if (req.method !== 'POST') {
    return res.status(405).json({ error: 'Method not allowed' });
  }

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    return res.status(503).json({
      reply: 'Records chat is not configured. Set OPENAI_API_KEY in your deployment environment (e.g. Vercel).',
    });
  }

  let body;
  try {
    body = typeof req.body === 'string' ? JSON.parse(req.body) : req.body || {};
  } catch {
    return res.status(400).json({ error: 'Invalid JSON body' });
  }

  const { messages = [], context = '' } = body;
  if (!Array.isArray(messages) || messages.length === 0) {
    return res.status(400).json({ error: 'messages array required' });
  }

  const systemContent =
    'You are a helpful assistant that answers questions about the user\'s records: sessions (meetings, reminders, tasks), notes, agenda items, stakeholders, and transactions (expenses/income). Answer only from the context provided. Be concise and factual. If the context does not contain enough information, say so.';

  const apiMessages = [
    { role: 'system', content: systemContent + '\n\n--- Context (user\'s records) ---\n' + (context || '(no context)') },
    ...messages.map((m) => ({ role: m.role, content: m.content })),
  ];

  try {
    const response = await fetch('https://api.openai.com/v1/chat/completions', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: 'gpt-4o-mini',
        messages: apiMessages,
        max_tokens: 1024,
        temperature: 0.3,
      }),
    });

    if (!response.ok) {
      const errText = await response.text();
      console.error('OpenAI error', response.status, errText);
      return res.status(502).json({
        reply: 'The assistant service returned an error. Try again or check OPENAI_API_KEY and quota.',
      });
    }

    const data = await response.json();
    const reply =
      data.choices?.[0]?.message?.content?.trim() ||
      'I didn’t get a reply. Please try again.';

    return res.status(200).json({ reply });
  } catch (err) {
    console.error('Chat API error', err);
    return res.status(500).json({
      reply: 'Something went wrong. Please try again.',
    });
  }
}
