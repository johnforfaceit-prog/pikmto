const express = require('express');
const router  = express.Router();
const fetch   = require('node-fetch');
const mammoth = require('mammoth');
const JSZip   = require('jszip');

// ── Проксируем запросы к Claude ──
// Решает CORS — браузер обращается к нашему серверу, сервер — к Anthropic
router.post('/claude', async (req, res) => {
  try {
    const apiKey = req.body.apiKey || process.env.ANTHROPIC_API_KEY;
    if (!apiKey) return res.status(400).json({ error: 'API ключ не указан' });

    const { messages, model, max_tokens, system } = req.body;

    const payload = { model: model || 'claude-sonnet-4-6', max_tokens: max_tokens || 1000, messages };
    if (system) payload.system = system;

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify(payload)
    });

    const data = await response.json();
    if (!response.ok) return res.status(response.status).json(data);
    res.json(data);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Парсинг DOCX — извлекаем текст и HTML ──
router.post('/parse-docx', async (req, res) => {
  try {
    const { fileBase64 } = req.body;
    if (!fileBase64) return res.status(400).json({ error: 'Файл не передан' });

    const buffer = Buffer.from(fileBase64, 'base64');

    // HTML для просмотра
    const htmlResult  = await mammoth.convertToHtml({ buffer });
    // Текст для парсинга реквизитов
    const textResult  = await mammoth.extractRawText({ buffer });

    res.json({ html: htmlResult.value, text: textResult.value });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Патч DOCX — подменяем значения в XML и отдаём новый файл ──
router.post('/patch-docx', async (req, res) => {
  try {
    const { fileBase64, patches } = req.body;
    // patches = [{ from: 'старый текст', to: 'новый текст' }, ...]

    const buffer = Buffer.from(fileBase64, 'base64');
    const zip = await JSZip.loadAsync(buffer);
    let xml = await zip.file('word/document.xml').async('string');

    for (const { from, to } of (patches || [])) {
      if (from && to && from !== to) xml = xml.split(from).join(to);
    }

    // Пишем обратно
    zip.file('word/document.xml', xml);
    const outBuffer = await zip.generateAsync({
      type: 'nodebuffer',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="zaklyuchenie.docx"');
    res.send(outBuffer);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Статус сервера ──
router.get('/status', (req, res) => {
  res.json({ ok: true, user: req.session.user?.login });
});

module.exports = router;
