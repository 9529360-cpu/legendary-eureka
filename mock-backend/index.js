const express = require('express');
const bodyParser = require('body-parser');
const app = express();
const port = process.env.PORT || 3001;
app.use(bodyParser.json());

app.get('/', (req, res) => {
  res.json({ status: 'ok', name: 'mock-deepseek-backend' });
});

app.post('/ai', (req, res) => {
  const { prompt } = req.body || {};
  // simple mock reply
  res.json({
    id: Math.random().toString(36).slice(2),
    response: `模拟响应：收到 prompt => ${prompt || '<empty>'}`,
    timestamp: Date.now(),
  });
});

app.listen(port, () => {
  console.log(`Mock AI backend listening on http://localhost:${port}`);
});
