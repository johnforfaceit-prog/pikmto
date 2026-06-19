const express = require('express');
const router  = express.Router();
const fetch   = require('node-fetch');
const mammoth = require('mammoth');
const JSZip   = require('jszip');
const fs      = require('fs');
const path    = require('path');

// ── Шаблон итогового заключения (один на все документы) ──
const TEMPLATE_PATH = path.join(__dirname, '..', 'templates', 'zaklyuchenie_template.docx');
const ZAK_TOKENS = [
  'ZAK_NUM','ZAK_DAY','ZAK_MONTH','CONTRACT_NUM','CONTRACT_DATE','SUBJECT','EXECUTOR','ADDRESS',
  'S_NAME','S_VOL_CONTRACT','S_VOL_PERIOD','S_VOL_FACT_PERIOD','S_VOL_FACT','S_COST_PERIOD','S_COST_FACT',
  'INV_NUM','INV_DATE','INV_DATE_PLAN','INV_DATE_FACT','UPD_NUM','UPD_DATE','UPD_DATE_PLAN','UPD_DATE_FACT',
  'LBL_EXECUTOR','LBL_ADDRESS','LBL_INV','LBL_UPD'
];
// Значения по умолчанию: подписи строк + дата заключения (всегда пустая — заполняется от руки)
const LABEL_DEFAULTS = {
  LBL_EXECUTOR: 'Наименование исполнителя',
  LBL_ADDRESS:  'Место нахождения, адрес',
  LBL_INV:      'Счёт на оплату',
  LBL_UPD:      'Документ о приёмке (функция СЧФДОП)*',
  ZAK_DAY:      '___',
  ZAK_MONTH:    '__________'
};
function escapeXml(s) {
  return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

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

// ── Сборка итогового заключения из встроенного шаблона ──
router.post('/build-zaklyuchenie', async (req, res) => {
  try {
    const fields = req.body.fields || {};
    if (!fs.existsSync(TEMPLATE_PATH)) return res.status(500).json({ error: 'Шаблон заключения не найден на сервере' });

    const zip = await JSZip.loadAsync(fs.readFileSync(TEMPLATE_PATH));
    let xml = await zip.file('word/document.xml').async('string');

    for (const t of ZAK_TOKENS) {
      let v = fields[t];
      if ((v == null || v === '') && LABEL_DEFAULTS[t] != null) v = LABEL_DEFAULTS[t];
      xml = xml.split('{{' + t + '}}').join(escapeXml(v));
    }

    zip.file('word/document.xml', xml);
    const out = await zip.generateAsync({
      type: 'nodebuffer',
      mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
    });

    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.setHeader('Content-Disposition', 'attachment; filename="zaklyuchenie.docx"');
    res.send(out);
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── Статус сервера ──
router.get('/status', (req, res) => {
  res.json({ ok: true, user: req.session.user?.login });
});

module.exports = router;
