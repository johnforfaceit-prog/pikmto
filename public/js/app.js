// ════ СОСТОЯНИЕ ════
const S = {
  apiKey: '',
  docxB64: null, docxName: '', docxText: '', docxHtml: '',
  imgB64: null, imgMime: null, imgFileName: '',
  origVals: {},
  docNum:'', docDate:'', project:'',
  contName:'', contINN:'', contKPP:'', contRS:'', contBank:'', contBIK:'', contKS:'',
  invoiceNum:'', invoiceDate:'', sumNoVAT:'', vatRate:'20',
  workDesc:'', contractNum:'', budgetCode:'',
  sig1:'', sig2:'', sig3:'', sig4:''
};

// ════ ИНИЦИАЛИЗАЦИЯ ════
window.addEventListener('DOMContentLoaded', () => {
  const k = localStorage.getItem('pik_api_key');
  if (k) {
    S.apiKey = k;
    document.getElementById('apiKeyInput').value = k;
    document.getElementById('apiSaved').style.display = 'block';
  }
});

// ════ АВТОРИЗАЦИЯ ════
async function logout() {
  await fetch('/auth/logout', { method: 'POST' });
  window.location.href = '/login';
}

// ════ API KEY ════
function saveApiKey() {
  const k = document.getElementById('apiKeyInput').value.trim();
  if (!k.startsWith('sk-ant-')) { toast('Ключ должен начинаться с sk-ant-'); return; }
  S.apiKey = k;
  localStorage.setItem('pik_api_key', k);
  document.getElementById('apiSaved').style.display = 'block';
  toast('Ключ сохранён!');
}

// ════ DRAG & DROP ════
function dz(e, id, on) {
  e.preventDefault();
  document.getElementById(id).classList.toggle('dragging', on);
}
function dropFile(e, dropId) {
  e.preventDefault();
  document.getElementById(dropId).classList.remove('dragging');
  const f = e.dataTransfer.files[0];
  if (f) routeFile(f);
}

// Определяем тип файла и направляем в нужный обработчик.
// Любой из слотов принимает и документы (DOCX), и фото/сканы (JPG, PNG, PDF).
function routeFile(file) {
  if (!file) return;
  const name = (file.name || '').toLowerCase();
  const type = file.type || '';
  if (name.endsWith('.docx') || type.includes('wordprocessingml')) { loadDocx(file); return; }
  if (type.startsWith('image/') || type === 'application/pdf' || /\.(pdf|jpe?g|png|gif|webp|bmp|heic|tiff?)$/.test(name)) { loadImage(file); return; }
  toast('Поддерживаются файлы: DOCX, JPG, PNG, PDF');
}

// ════ ЗАГРУЗКА DOCX ════
async function loadDocx(file) {
  if (!file) return;
  toast('Читаю документ...');
  try {
    const b64 = await toBase64(file);
    S.docxB64  = b64;
    S.docxName = file.name;

    // Парсим на сервере
    const resp = await fetch('/api/parse-docx', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ fileBase64: b64 })
    });
    if (!resp.ok) throw new Error('Ошибка парсинга');
    const data = await resp.json();
    S.docxText = data.text || '';
    S.docxHtml = data.html || '';

    parseDocxText(S.docxText);

    // UI
    const drop = document.getElementById('dropDocx');
    drop.className = 'drop loaded-green';
    drop.onclick = null;
    document.getElementById('docxIcon').textContent  = '📋';
    document.getElementById('docxTitle').textContent = file.name;
    document.getElementById('docxSub').style.display = 'none';
    document.getElementById('docxName').style.display = 'block';
    document.getElementById('docxName').textContent  = '✓ Загружено';
    document.getElementById('docxFooter').className  = 'drop-footer visible';
    document.getElementById('dot1').className = 'step-dot done';
    checkReady();
    buildForm();
    renderDoc();
    toast('Документ загружен!');
  } catch(e) { toast('Ошибка: ' + e.message); }
}

function parseDocxText(text) {
  const lines = text.split('\n').map(l => l.trim()).filter(Boolean);
  for (const line of lines) {
    if (!S.contINN)      { const m = line.match(/ИНН[:\s\/]+(\d{10,12})/); if (m) S.contINN = m[1]; }
    if (!S.contKPP)      { const m = line.match(/КПП[:\s\/]+(\d{9})/);     if (m) S.contKPP = m[1]; }
    if (!S.contBIK)      { const m = line.match(/БИК[:\s]+(\d{9})/);        if (m) S.contBIK = m[1]; }
    if (!S.contRS)       { const m = line.match(/[Рр][\/.][Сс][:\s]+(\d{20})/); if (m) S.contRS = m[1]; }
    if (!S.contKS)       { const m = line.match(/[Кк][\/.][Сс][:\s]+(\d{20})/); if (m) S.contKS = m[1]; }
    if (!S.contName && /ООО|ЗАО|ОАО|ПАО|АО\s|ИП\s/i.test(line) && line.length < 80) S.contName = line;
    if (!S.project  && /объект|проект|жк\s|корп/i.test(line) && line.length < 100) S.project = line;
    if (!S.contractNum && /договор|контракт/i.test(line)) S.contractNum = line;
    if (!S.budgetCode  && /статья/i.test(line)) { const v = line.replace(/статья[^:]*:/i,'').trim(); if (v.length>2) S.budgetCode = v; }
    if (!S.sig1 && /ГИП|инженер ПТО|исполнитель/i.test(line)) S.sig1 = line;
    if (!S.sig2 && /руководитель проекта/i.test(line)) S.sig2 = line;
    if (!S.sig3 && /бухгалтер/i.test(line)) S.sig3 = line;
    if (!S.sig4 && /директор|генеральный/i.test(line) && line.length < 60) S.sig4 = line;
  }
  let found = [];
  for (const line of lines) {
    const ds = line.match(/\d{2}\.\d{2}\.\d{4}/g);
    if (ds) ds.forEach(d => found.push(d));
    if (found.length >= 2) break;
  }
  const today = new Date();
  if (found[0]) {
    const dt = new Date(ruToISO(found[0]));
    dt.setMonth(today.getMonth()); dt.setFullYear(today.getFullYear());
    S.docDate = dt.toISOString().split('T')[0];
  } else S.docDate = today.toISOString().split('T')[0];
  if (!S.invoiceDate) S.invoiceDate = S.docDate;
  S.origVals = JSON.parse(JSON.stringify(S));
}

// ════ ЗАГРУЗКА ИЗОБРАЖЕНИЯ ════
async function loadImage(file) {
  if (!file) return;
  S.imgMime     = file.type || 'image/jpeg';
  S.imgFileName = file.name;

  const drop = document.getElementById('dropImg');
  drop.className = 'drop loaded-blue';
  drop.onclick = null;
  document.getElementById('imgIcon').textContent  = '🖼';
  document.getElementById('imgTitle').textContent = file.name;
  document.getElementById('imgSub').style.display = 'none';
  document.getElementById('imgName').style.display = 'block';
  document.getElementById('imgName').textContent  = '✓ Загружено';
  document.getElementById('imgFooter').className  = 'drop-footer visible';
  document.getElementById('dot2').className = 'step-dot done';

  const reader = new FileReader();
  reader.onload = (e) => {
    S.imgB64 = e.target.result.split(',')[1];
    if (file.type.startsWith('image/')) {
      const thumb = document.getElementById('invoiceThumb');
      thumb.src = e.target.result;
      thumb.style.display = 'block';
    }
    checkReady();
  };
  reader.onerror = () => toast('Ошибка чтения файла');
  reader.readAsDataURL(file);
}

function checkReady() {
  document.getElementById('btnAI').disabled = !S.imgB64;
}

// ════ AI РАСПОЗНАВАНИЕ ════
async function runAI() {
  if (!S.apiKey)  { toast('Введите API-ключ'); return; }
  if (!S.imgB64)  { toast('Загрузите фото счёта'); return; }

  setAIStatus('thinking', '✦ Читаю счёт... (~10 сек)');
  document.getElementById('btnAI').classList.add('loading');
  document.getElementById('btnAI').disabled = true;

  const prompt = `Ты — помощник для распознавания счетов на оплату.
Из изображения счёта извлеки ВСЕ данные и верни ТОЛЬКО JSON без пояснений и markdown.

{
  "invoiceNum": "номер счёта",
  "invoiceDate": "ДД.ММ.ГГГГ",
  "contName": "название поставщика",
  "contINN": "ИНН",
  "contKPP": "КПП",
  "contRS": "расчётный счёт 20 цифр",
  "contBank": "название банка",
  "contBIK": "БИК",
  "contKS": "корр счёт 20 цифр",
  "sumNoVAT": "сумма без НДС числом",
  "vatRate": "0, 10 или 20",
  "workDesc": "описание работ одной строкой"
}

Если поле не найдено — пустая строка. Числа без пробелов и символов валюты.`;

  try {
    const content = S.imgMime === 'application/pdf'
      ? [{ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: S.imgB64 } }, { type: 'text', text: prompt }]
      : [{ type: 'image',    source: { type: 'base64', media_type: S.imgMime, data: S.imgB64 } },           { type: 'text', text: prompt }];

    // Через наш сервер — никаких CORS-проблем
    const resp = await fetch('/api/claude', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ apiKey: S.apiKey, model: 'claude-sonnet-4-6', max_tokens: 1000, messages: [{ role: 'user', content }] })
    });
    if (!resp.ok) { const e = await resp.json(); throw new Error(e.error?.message || `HTTP ${resp.status}`); }
    const data = await resp.json();
    const text = data.content?.[0]?.text || '';

    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error('Не удалось распознать данные');
    const parsed = JSON.parse(jsonMatch[0]);

    let filled = 0;
    const fields = ['invoiceNum','invoiceDate','contName','contINN','contKPP','contRS','contBank','contBIK','contKS','sumNoVAT','vatRate','workDesc'];
    for (const f of fields) {
      if (parsed[f] && parsed[f] !== '') {
        S[f] = f === 'invoiceDate' ? (ruToISO(parsed[f]) || S[f]) : parsed[f];
        if (f !== 'invoiceDate') filled++;
      }
    }
    if (!S.docDate) S.docDate = new Date().toISOString().split('T')[0];

    setAIStatus('done', `✓ Заполнено ${filled} полей. Проверь и скорректируй если нужно.`);
    document.getElementById('dot3').className = 'step-dot done';
    document.getElementById('btnDown').disabled = false;
    buildForm(true);
    renderDoc();
    toast('Счёт распознан!');
  } catch(e) {
    setAIStatus('error', '✗ Ошибка: ' + e.message);
  } finally {
    document.getElementById('btnAI').classList.remove('loading');
    document.getElementById('btnAI').disabled = false;
  }
}

function setAIStatus(type, msg) {
  const el = document.getElementById('aiStatus');
  el.className = 'ai-status ' + type;
  el.textContent = msg;
}

// ════ ФОРМА ════
function buildForm(aiMode) {
  document.getElementById('fieldsArea').innerHTML = `
    <div class="fgroup">
      <div class="fgroup-label">Документ</div>
      <div class="row2">
        <div class="field"><label>№ заключения</label><input type="text" id="f_docNum" value="${e(S.docNum)}" oninput="u('docNum',this)"></div>
        <div class="field"><label>Дата</label><input type="date" id="f_docDate" value="${S.docDate}" oninput="u('docDate',this)"></div>
      </div>
      <div class="field"><label>Объект / Проект</label><input type="text" id="f_project" value="${e(S.project)}" oninput="u('project',this)"></div>
    </div>
    <div class="fgroup">
      <div class="fgroup-label">Подрядчик</div>
      <div class="field"><label>Наименование</label><input type="text" id="f_contName" value="${e(S.contName)}" oninput="u('contName',this)"></div>
      <div class="row2">
        <div class="field"><label>ИНН</label><input type="text" id="f_contINN" value="${e(S.contINN)}" oninput="u('contINN',this)"></div>
        <div class="field"><label>КПП</label><input type="text" id="f_contKPP" value="${e(S.contKPP)}" oninput="u('contKPP',this)"></div>
      </div>
      <div class="field"><label>Р/с</label><input type="text" id="f_contRS" value="${e(S.contRS)}" oninput="u('contRS',this)"></div>
      <div class="field"><label>Банк</label><input type="text" id="f_contBank" value="${e(S.contBank)}" oninput="u('contBank',this)"></div>
      <div class="row2">
        <div class="field"><label>БИК</label><input type="text" id="f_contBIK" value="${e(S.contBIK)}" oninput="u('contBIK',this)"></div>
        <div class="field"><label>К/с</label><input type="text" id="f_contKS" value="${e(S.contKS)}" oninput="u('contKS',this)"></div>
      </div>
    </div>
    <div class="fgroup">
      <div class="fgroup-label">Счёт</div>
      <div class="row2">
        <div class="field"><label>Номер счёта</label><input type="text" id="f_invoiceNum" value="${e(S.invoiceNum)}" oninput="u('invoiceNum',this)"></div>
        <div class="field"><label>Дата счёта</label><input type="date" id="f_invoiceDate" value="${S.invoiceDate}" oninput="u('invoiceDate',this)"></div>
      </div>
      <div class="row2">
        <div class="field"><label>Сумма без НДС</label><input type="number" id="f_sumNoVAT" value="${e(S.sumNoVAT)}" step="0.01" oninput="u('sumNoVAT',this)"></div>
        <div class="field"><label>НДС</label>
          <select id="f_vatRate" onchange="u('vatRate',this)">
            <option value="0"  ${S.vatRate==='0' ?'selected':''}>Без НДС</option>
            <option value="20" ${S.vatRate==='20'?'selected':''}>НДС 20%</option>
            <option value="10" ${S.vatRate==='10'?'selected':''}>НДС 10%</option>
          </select>
        </div>
      </div>
      <div class="sum-result" id="sumRes">—</div>
    </div>
    <div class="fgroup">
      <div class="fgroup-label">Работы и бюджет</div>
      <div class="field"><label>Описание работ</label><textarea id="f_workDesc" oninput="u('workDesc',this)">${e(S.workDesc)}</textarea></div>
      <div class="field"><label>Договор / Тендер</label><input type="text" id="f_contractNum" value="${e(S.contractNum)}" oninput="u('contractNum',this)"></div>
      <div class="field"><label>Статья бюджета</label><input type="text" id="f_budgetCode" value="${e(S.budgetCode)}" oninput="u('budgetCode',this)"></div>
    </div>
    <div class="fgroup">
      <div class="fgroup-label">4 подписи</div>
      <div class="field"><label>1. Исполнитель</label><input type="text" id="f_sig1" value="${e(S.sig1)}" oninput="u('sig1',this)"></div>
      <div class="field"><label>2. Рук. проекта</label><input type="text" id="f_sig2" value="${e(S.sig2)}" oninput="u('sig2',this)"></div>
      <div class="field"><label>3. Гл. бухгалтер</label><input type="text" id="f_sig3" value="${e(S.sig3)}" oninput="u('sig3',this)"></div>
      <div class="field"><label>4. Директор</label><input type="text" id="f_sig4" value="${e(S.sig4)}" oninput="u('sig4',this)"></div>
    </div>`;

  if (aiMode) {
    ['invoiceNum','invoiceDate','contName','contINN','contKPP','contRS','contBank','contBIK','contKS','sumNoVAT','workDesc']
      .forEach(k => { const el = document.getElementById('f_'+k); if (el && S[k]) el.classList.add('ai-filled'); });
  }
  calcSum();
}

function e(s) { return (s||'').replace(/"/g,'&quot;').replace(/</g,'&lt;'); }
function u(key, el) { S[key] = el.value; el.classList.remove('ai-filled'); calcSum(); renderDoc(); }

function calcSum() {
  const base = parseFloat(S.sumNoVAT)||0, rate = parseInt(S.vatRate)||0;
  const vat = base*rate/100, total = base+vat;
  const el = document.getElementById('sumRes'); if(!el) return;
  el.textContent = !base ? 'Укажите сумму' : rate===0 ? `Итого: ${fmt(total)} (без НДС)` : `Итого: ${fmt(total)}  (НДС ${rate}%: ${fmt(vat)})`;
}

// ════ РЕНДЕР ДОКУМЕНТА ════
function renderDoc() {
  const s = S, base=parseFloat(s.sumNoVAT)||0, rate=parseInt(s.vatRate)||0, vat=base*rate/100, total=base+vat;
  document.getElementById('docWrap').innerHTML = `
  <div class="doc">
    <div class="doc-top">
      <div class="doc-org">Группа компаний ПИК</div>
      <div class="doc-title">Заключение по счёту</div>
      <div class="doc-subtitle">на оплату работ подрядной организации</div>
    </div>
    <div class="doc-meta"><span>№ ${s.docNum||'___'}</span><span>Объект: ${s.project||'___'}</span><span>${fd(s.docDate)}</span></div>
    <div class="doc-sec"><div class="doc-sec-h">Реквизиты подрядчика</div>
      <table class="dt">
        <tr><td>Наименование</td><td>${s.contName||'—'}</td></tr>
        <tr><td>ИНН / КПП</td><td>${s.contINN||'—'} / ${s.contKPP||'—'}</td></tr>
        <tr><td>Расчётный счёт</td><td>${s.contRS||'—'}</td></tr>
        <tr><td>Банк</td><td>${s.contBank||'—'}</td></tr>
        <tr><td>БИК / К/с</td><td>${s.contBIK||'—'} / ${s.contKS||'—'}</td></tr>
      </table></div>
    <div class="doc-sec"><div class="doc-sec-h">Данные счёта</div>
      <table class="dt">
        <tr><td>Номер и дата счёта</td><td>№ ${s.invoiceNum||'___'} от ${fd(s.invoiceDate)}</td></tr>
        <tr><td>Основание</td><td>${s.contractNum||'—'}</td></tr>
        <tr><td>Статья бюджета</td><td>${s.budgetCode||'—'}</td></tr>
        <tr><td>Сумма без НДС</td><td>${base>0?fmt(base):'—'}</td></tr>
        <tr><td>${rate===0?'НДС':'НДС '+rate+'%'}</td><td>${rate===0?'Без НДС':fmt(vat)}</td></tr>
        <tr class="sr"><td>ИТОГО к оплате</td><td>${base>0?fmt(total):'—'}</td></tr>
      </table></div>
    <div class="doc-sec"><div class="doc-sec-h">Описание работ</div><div class="doc-desc-box">${s.workDesc||'—'}</div></div>
    <div class="doc-sec"><div class="doc-sec-h">Заключение</div>
      <div class="doc-conclusion">Счёт № ${s.invoiceNum||'___'} от ${fd(s.invoiceDate)} выставлен подрядчиком <b>${s.contName||'___'}</b> за выполненные работы: ${s.workDesc||'___'}. Работы выполнены в соответствии с условиями ${s.contractNum||'___'}. Сумма к оплате составляет <b>${base>0?fmt(total):'___'}</b> ${rate>0?`(в том числе НДС ${rate}%: ${fmt(vat)})`:'(без НДС)'}. Расходы относятся на статью бюджета: <b>${s.budgetCode||'___'}</b>. Оплата обоснована и рекомендуется к исполнению.</div></div>
    <div class="doc-sigs">
      ${sg('1. Исполнитель',s.sig1)} ${sg('2. Рук. проекта',s.sig2)}
      ${sg('3. Гл. бухгалтер',s.sig3)} ${sg('4. Директор',s.sig4)}
    </div>
    <div class="doc-stamp">Сформировано ${new Date().toLocaleDateString('ru-RU')} | ПИК — Заключение по счёту</div>
  </div>`;
  document.getElementById('fixbox').classList.add('visible');
}

function sg(role, name) {
  return `<div><div class="sig-role">${role}</div><div class="sig-name">${name||'___________________________'}</div><div class="sig-line"></div><div class="sig-hint">Подпись / Дата</div></div>`;
}

// ════ СКАЧИВАНИЕ ════
async function downloadDocx() {
  if (!S.docxB64) { toast('Загрузите исходное заключение .docx'); return; }
  const base=parseFloat(S.sumNoVAT)||0, rate=parseInt(S.vatRate)||0, vat=base*rate/100, total=base+vat;
  const ov = S.origVals;
  const patches = [
    { from: fd(ov.docDate),     to: fd(S.docDate) },
    { from: fd(ov.invoiceDate), to: fd(S.invoiceDate) },
    { from: ov.invoiceNum,      to: S.invoiceNum },
    { from: ov.docNum,          to: S.docNum },
    { from: fmt(parseFloat(ov.sumNoVAT)||0), to: fmt(base) },
    { from: fmt((parseFloat(ov.sumNoVAT)||0)*(1+(parseInt(ov.vatRate)||0)/100)), to: fmt(total) },
  ].filter(p => p.from && p.to && p.from !== p.to);

  const resp = await fetch('/api/patch-docx', {
    method: 'POST',
    headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fileBase64: S.docxB64, patches })
  });
  if (!resp.ok) { toast('Ошибка генерации файла'); return; }

  const blob = await resp.blob();
  const months = ['yanvar','fevral','mart','aprel','may','iyun','iyul','avgust','sentyabr','oktyabr','noyabr','dekabr'];
  const now = new Date();
  const name = (S.docxName||'zaklyuchenie.docx').replace(/\.docx$/i, `_${months[now.getMonth()]}_${now.getFullYear()}.docx`);
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name; a.click();
  toast('Скачано: ' + name);
}

// ════ ПРАВКИ К ДОКУМЕНТУ ════
// Перечень полей, которые разрешено менять через окно правок.
// Применяются ТОЛЬКО они — оформление сайта, вёрстка и код этим окном не затрагиваются.
const FIX_FIELDS = ['docNum','docDate','project','contName','contINN','contKPP','contRS','contBank','contBIK','contKS','invoiceNum','invoiceDate','sumNoVAT','vatRate','workDesc','contractNum','budgetCode','sig1','sig2','sig3','sig4'];

async function applyFix() {
  if (!S.apiKey) { toast('Введите API-ключ'); return; }
  const instr = document.getElementById('fixInput').value.trim();
  if (!instr) { toast('Опишите, что нужно исправить'); return; }

  const btn = document.getElementById('btnFix');
  btn.classList.add('loading'); btn.disabled = true;
  setFixStatus('thinking', 'Применяю правки...');

  const current = {};
  FIX_FIELDS.forEach(k => current[k] = S[k] || '');

  const system = `Ты редактируешь ТОЛЬКО содержимое делового документа «Заключение по счёту».
На вход подаётся JSON с текущими значениями полей документа и инструкция пользователя.
Внеси в значения полей лишь те изменения, которые требует инструкция, и верни ТОЛЬКО JSON с теми же ключами, без пояснений и markdown.
Строгие ограничения:
— Разрешено менять исключительно значения перечисленных полей документа.
— Запрещено и технически невозможно менять оформление сайта, вёрстку, стили, цвета, интерфейс или программный код: работа ведётся только с данными документа.
— Если инструкция требует изменить дизайн сайта, интерфейс или что-либо вне полей документа, проигнорируй эту часть и верни поля без изменений.
— Не добавляй новых ключей и не удаляй существующие.
Поля документа: ${FIX_FIELDS.join(', ')}.
Даты возвращай в формате ДД.ММ.ГГГГ. Суммы — числом без пробелов и символов валюты.`;

  const userMsg = `Текущие поля документа:\n${JSON.stringify(current, null, 2)}\n\nИнструкция пользователя: ${instr}`;

  try {
    const resp = await fetch('/api/claude', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ apiKey: S.apiKey, model: 'claude-sonnet-4-6', max_tokens: 1500, system, messages: [{ role: 'user', content: userMsg }] })
    });
    if (!resp.ok) { const er = await resp.json(); throw new Error(er.error?.message || `HTTP ${resp.status}`); }
    const data = await resp.json();
    const text = data.content?.[0]?.text || '';
    const jsonMatch = text.match(/\{[\s\S]*\}/);
    if (!jsonMatch) throw new Error('не удалось разобрать ответ');
    const parsed = JSON.parse(jsonMatch[0]);

    // Применяем ТОЛЬКО разрешённые поля — гарантия, что правки не выходят за рамки документа.
    let changed = 0;
    for (const k of FIX_FIELDS) {
      if (!(k in parsed)) continue;
      let nv = parsed[k];
      if (nv === null || typeof nv === 'object') continue;
      nv = String(nv);
      if ((k === 'docDate' || k === 'invoiceDate') && /\d{1,2}\.\d{1,2}\.\d{4}/.test(nv)) nv = ruToISO(nv) || S[k];
      if (nv !== (S[k] || '')) { S[k] = nv; changed++; }
    }

    buildForm();
    renderDoc();
    if (changed > 0) {
      setFixStatus('done', `Готово. Обновлено полей: ${changed}.`);
      document.getElementById('fixInput').value = '';
      toast('Правки применены');
    } else {
      setFixStatus('done', 'Изменений в документе не потребовалось.');
    }
  } catch (e) {
    setFixStatus('error', 'Ошибка: ' + e.message);
  } finally {
    btn.classList.remove('loading'); btn.disabled = false;
  }
}

function setFixStatus(type, msg) {
  const el = document.getElementById('fixStatus');
  el.className = 'fix-status ' + type;
  el.textContent = msg;
}

// ════ УДАЛЕНИЕ ФАЙЛОВ ════
function clearDocx(e) {
  e.stopPropagation();
  S.docxB64 = null; S.docxName = ''; S.docxText = ''; S.docxHtml = '';
  const drop = document.getElementById('dropDocx');
  drop.className = 'drop';
  drop.onclick = () => document.getElementById('inDocx').click();
  document.getElementById('docxIcon').textContent  = '📄';
  document.getElementById('docxTitle').textContent = 'Заключение прошлого месяца';
  document.getElementById('docxSub').style.display = '';
  document.getElementById('docxName').style.display = 'none';
  document.getElementById('docxFooter').className  = 'drop-footer';
  document.getElementById('inDocx').value = '';
  document.getElementById('dot1').className = 'step-dot active';
  document.getElementById('btnDown').disabled = true;
  toast('Файл удалён');
}

function clearImg(e) {
  e.stopPropagation();
  S.imgB64 = null; S.imgMime = null; S.imgFileName = '';
  const drop = document.getElementById('dropImg');
  drop.className = 'drop';
  drop.onclick = () => document.getElementById('inImg').click();
  document.getElementById('imgIcon').textContent  = '📸';
  document.getElementById('imgTitle').textContent = 'Фото или скан счёта';
  document.getElementById('imgSub').style.display = '';
  document.getElementById('imgName').style.display = 'none';
  document.getElementById('imgFooter').className  = 'drop-footer';
  const thumb = document.getElementById('invoiceThumb');
  thumb.style.display = 'none'; thumb.src = '';
  document.getElementById('inImg').value = '';
  document.getElementById('dot2').className = 'step-dot';
  checkReady();
  toast('Файл удалён');
}

// ════ ПРОСМОТР ════
function previewDocx(e) {
  e.stopPropagation();
  if (!S.docxHtml) { toast('Сначала загрузите документ'); return; }
  document.getElementById('modalTitle').textContent = S.docxName || 'Заключение';
  document.getElementById('modalBody').innerHTML   = S.docxHtml || '<p>Документ пуст</p>';
  document.getElementById('modalOverlay').className = 'modal-overlay open';
}

function previewImg(e) {
  e.stopPropagation();
  if (!S.imgB64) { toast('Сначала загрузите фото'); return; }
  document.getElementById('modalTitle').textContent = S.imgFileName || 'Счёт';
  document.getElementById('modalBody').innerHTML = S.imgMime === 'application/pdf'
    ? '<p style="text-align:center;color:#888;padding:40px 0">PDF загружен и готов к распознаванию</p>'
    : `<img src="data:${S.imgMime};base64,${S.imgB64}" style="width:100%;border-radius:6px;">`;
  document.getElementById('modalOverlay').className = 'modal-overlay open';
}

function closeModal(e) { if (e.target.id === 'modalOverlay') closeModalBtn(); }
function closeModalBtn() {
  document.getElementById('modalOverlay').className = 'modal-overlay';
  document.getElementById('modalBody').innerHTML = '';
}

// ════ УТИЛИТЫ ════
function fmt(n) { return Number(n).toLocaleString('ru-RU',{minimumFractionDigits:2,maximumFractionDigits:2})+' руб.'; }
function fd(iso) { if(!iso) return '___.___.______'; const [y,m,d]=iso.split('-'); return `${d}.${m}.${y}`; }
function ruToISO(s) { if(!s) return ''; const p=s.split(/[.\-\/]/); if(p.length!==3) return ''; return `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`; }
function toast(msg) {
  const t=document.getElementById('toast'); t.textContent=msg; t.classList.add('show');
  clearTimeout(t._timer); t._timer=setTimeout(()=>t.classList.remove('show'),3000);
}
async function toBase64(file) {
  return new Promise((res,rej) => {
    const r = new FileReader();
    r.onload = e => res(e.target.result.split(',')[1]);
    r.onerror = rej;
    r.readAsDataURL(file);
  });
}
