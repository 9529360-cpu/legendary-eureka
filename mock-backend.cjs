/* Simple mock backend for AI replies on port 3002 */
const express = require('express');
const bodyParser = require('body-parser');
const app = express();
const port = process.env.PORT || 3002;

app.use(bodyParser.json());

app.get('/api/config/status', (req, res) => {
  res.json({
    success: true,
    configured: true,
    isValid: true,
    lastUpdated: new Date().toISOString(),
    maskedKey: 'sk-mock-xxxx',
    service: 'Mock Backend',
    version: '1.0.0',
    environment: 'development',
    features: ['Mock Reply', 'Excel Copilot Addin'],
    config: { rateLimitWindowMs: 900000, rateLimitMaxRequests: 100, logLevel: 'info' }
  });
});

app.get('/health', (req, res) => {
  res.json({ status: 'ok' });
});

app.post('/reply', (req, res) => {
  const text = (req.body && req.body.text) ? String(req.body.text) : '';
  // very small mock reply logic
  let reply = 'Received: ' + text;
  let payload = null;
  if (/perceive|recognize/i.test(text)) {
    reply = 'Perception detected. Select Sheet1 A1.';
    payload = { sheet: 'Sheet1', selection: 'A1' };
  } else if (/execute|run/i.test(text)) {
    reply = 'Execution successful. Wrote "Executed" to A1.';
    payload = { action: 'write A1', status: 'ok' };
  }
  res.json({ reply, payload });
});

app.listen(port, () => {
  console.log(`Mock backend listening on http://localhost:${port}`);
});


