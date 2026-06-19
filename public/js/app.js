// ════ СОСТОЯНИЕ ════
const MONTHS_GEN = ['января','февраля','марта','апреля','мая','июня','июля','августа','сентября','октября','ноября','декабря'];
const SUBJECT_DEFAULT = 'Оказание услуг по ремонту и техническому обслуживанию специального электронного оборудования в 2026 году';

// Поля = токены шаблона заключения
function emptyFields() {
  return {
    ZAK_NUM:'', CONTRACT_NUM:'', CONTRACT_DATE:'', SUBJECT:'', EXECUTOR:'', ADDRESS:'',
    S_NAME:'', S_VOL_CONTRACT:'', S_VOL_PERIOD:'', S_VOL_FACT_PERIOD:'', S_VOL_FACT:'', S_COST_PERIOD:'', S_COST_FACT:'',
    INV_NUM:'', INV_DATE:'', INV_DATE_PLAN:'', INV_DATE_FACT:'',
    UPD_NUM:'', UPD_DATE:'', UPD_DATE_PLAN:'', UPD_DATE_FACT:''
  };
}

const S = {
  apiKey:'',
  mode:'existing',                                  // existing | new
  srcB64:null, srcName:'', srcMime:null, srcText:'', srcHtml:'',
  invFiles: [],                                     // листы по счёту: [{b64, mime, name}]
  zakDate: new Date().toISOString().split('T')[0],  // дата заключения (ISO)
  f: emptyFields()
};

// Поля, доступные окну правок (только содержимое документа)
const FIX_FIELDS = Object.keys(emptyFields());

// ════ ИНИЦИАЛИЗАЦИЯ ════
window.addEventListener('DOMContentLoaded', () => {
  const k = localStorage.getItem('pik_api_key');
  if (k) {
    S.apiKey = k;
    document.getElementById('apiKeyInput').value = k;
    document.getElementById('apiSaved').style.display = 'block';
  }
  S.f.SUBJECT = SUBJECT_DEFAULT;
  setMode('existing');
  buildForm();
  renderDoc();
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

// ════ РЕЖИМ РАБОТЫ ════
function setMode(mode) {
  S.mode = mode;
  document.getElementById('tabExisting').classList.toggle('active', mode === 'existing');
  document.getElementById('tabNew').classList.toggle('active', mode === 'new');

  if (mode === 'existing') {
    document.getElementById('modeHint').textContent =
      'Контракт уже вёлся ранее. Загрузите заключение прошлого месяца — реквизиты контракта возьмутся из него, читать сам контракт не нужно. Добавьте листы по счёту за новый период.';
    document.getElementById('srcTitle').textContent = 'Заключение прошлого месяца';
    document.getElementById('srcSub').textContent   = '.docx — нажмите или перетащите';
  } else {
    document.getElementById('modeHint').textContent =
      'Новый контракт, первый месяц. Загрузите контракт для ознакомления (реквизиты извлекутся автоматически) и листы по счёту.';
    document.getElementById('srcTitle').textContent = 'Контракт (для ознакомления)';
    document.getElementById('srcSub').textContent   = 'DOCX, PDF или фото';
  }
  checkReady();
}

// ════ DRAG & DROP ════
function dz(e, id, on) {
  e.preventDefault();
  document.getElementById(id).classList.toggle('dragging', on);
}
function dropFile(e, dropId, fn) {
  e.preventDefault();
  document.getElementById(dropId).classList.remove('dragging');
  const f = e.dataTransfer.files[0];
  if (f) fn(f);
}

// ════ ЗАГРУЗКА ИСТОЧНИКА (контракт / прошлое заключение) ════
async function loadSrc(file) {
  if (!file) return;
  toast('Читаю документ...');
  try {
    S.srcB64  = await toBase64(file);
    S.srcName = file.name;
    S.srcMime = file.type || guessMime(file.name);
    S.srcText = ''; S.srcHtml = '';

    // DOCX — извлекаем текст (нужен для разбора прошлого заключения и контракта)
    if (/\.docx$/i.test(file.name) || (S.srcMime || '').includes('wordprocessingml')) {
      const resp = await fetch('/api/parse-docx', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ fileBase64: S.srcB64 })
      });
      if (resp.ok) { const d = await resp.json(); S.srcText = d.text || ''; S.srcHtml = d.html || ''; }
    }

    const drop = document.getElementById('dropSrc');
    drop.className = 'drop loaded-green';
    drop.onclick = null;
    document.getElementById('srcIcon').textContent = '📋';
    document.getElementById('srcSub').style.display = 'none';
    document.getElementById('srcName').style.display = 'block';
    document.getElementById('srcName').textContent = '✓ ' + file.name;
    document.getElementById('srcFooter').className = 'drop-footer visible';
    checkReady();
    toast('Документ загружен!');
  } catch (e) { toast('Ошибка: ' + e.message); }
}

// ════ ЗАГРУЗКА ЛИСТОВ ПО СЧЁТУ (несколько файлов) ════
function onInvInput(input) {
  const files = Array.from(input.files || []); // копируем ДО очистки (input.files — живой список)
  input.value = '';
  addInvFiles(files);
}
function dropInv(e) {
  e.preventDefault();
  document.getElementById('dropInv').classList.remove('dragging');
  addInvFiles(Array.from(e.dataTransfer.files || []));
}

async function addInvFiles(fileList) {
  const files = Array.from(fileList || []);
  if (!files.length) return;
  for (const file of files) {
    try {
      const b64 = await toBase64(file);
      S.invFiles.push({ b64, mime: file.type || guessMime(file.name), name: file.name });
    } catch (_) { toast('Не удалось прочитать файл: ' + file.name); }
  }
  renderInvZone();
  checkReady();
  toast('Листов по счёту: ' + S.invFiles.length);
}

function removeInvFile(i, ev) {
  if (ev) ev.stopPropagation();
  S.invFiles.splice(i, 1);
  renderInvZone();
  checkReady();
}

function renderInvZone() {
  const drop = document.getElementById('dropInv');
  const n = S.invFiles.length;
  const sub = document.getElementById('invSub');
  const nm  = document.getElementById('invName');
  const tc  = document.getElementById('invThumbs');

  if (n === 0) {
    drop.className = 'drop';
    drop.onclick = () => document.getElementById('inInv').click();
    document.getElementById('invIcon').textContent = '📸';
    sub.style.display = ''; nm.style.display = 'none';
    document.getElementById('invFooter').className = 'drop-footer';
    tc.innerHTML = '';
    return;
  }

  drop.className = 'drop loaded-blue';
  drop.onclick = () => document.getElementById('inInv').click();   // клик добавляет ещё файлы
  document.getElementById('invIcon').textContent = '🧾';
  sub.style.display = 'none';
  nm.style.display = 'block';
  nm.textContent = '✓ ' + n + ' ' + plural(n, 'файл', 'файла', 'файлов') + ' · нажмите, чтобы добавить';
  document.getElementById('invFooter').className = 'drop-footer visible';

  tc.innerHTML = '';
  S.invFiles.forEach((f, i) => {
    const cell = document.createElement('div'); cell.className = 'inv-thumb-cell';
    if ((f.mime || '').startsWith('image/')) {
      const img = document.createElement('img'); img.src = 'data:' + f.mime + ';base64,' + f.b64; img.className = 'inv-thumb-sm';
      cell.appendChild(img);
    } else {
      const d = document.createElement('div'); d.className = 'inv-thumb-doc'; d.textContent = (f.mime === 'application/pdf' ? 'PDF' : 'ФАЙЛ');
      cell.appendChild(d);
    }
    const x = document.createElement('button'); x.className = 'inv-thumb-x'; x.textContent = '✕'; x.title = 'Удалить';
    x.onclick = (ev) => removeInvFile(i, ev);
    cell.appendChild(x);
    tc.appendChild(cell);
  });
}

function checkReady() {
  // Формировать можно, когда есть источник реквизитов
  document.getElementById('btnRun').disabled = !S.srcB64;
}

// ════ ФОРМИРОВАНИЕ ЗАКЛЮЧЕНИЯ ════
async function runExtract() {
  if (!S.apiKey) { toast('Введите API-ключ'); return; }
  if (!S.srcB64) { toast(S.mode === 'existing' ? 'Загрузите заключение прошлого месяца' : 'Загрузите контракт'); return; }

  const btn = document.getElementById('btnRun');
  btn.classList.add('loading'); btn.disabled = true;
  setRunStatus('thinking', '✦ Обрабатываю документы... (~10–20 сек)');

  try {
    if (S.mode === 'existing') {
      parsePrevZak(S.srcText);
    } else {
      await extractContract();
    }
    if (S.invFiles.length) await extractInvoice();

    buildForm(true);
    renderDoc();
    document.getElementById('btnDown').disabled = false;
    setRunStatus('done', '✓ Готово. Проверьте поля и при необходимости поправьте, затем скачайте .docx.');
    toast('Заключение сформировано!');
  } catch (e) {
    setRunStatus('error', '✗ Ошибка: ' + e.message);
  } finally {
    btn.classList.remove('loading'); btn.disabled = false;
  }
}

// Разбор прошлого заключения (документ уже в нашем формате)
function parsePrevZak(text) {
  if (!text) { toast('Не удалось прочитать текст заключения'); return; }
  const clean = v => (v && !/искать/i.test(v)) ? v.trim() : '';
  let m;
  if ((m = text.match(/ЗАКЛЮЧЕНИЕ\s*№\s*([^\n]+)/i)))            S.f.ZAK_NUM      = clean(m[1]) || S.f.ZAK_NUM;
  if ((m = text.match(/Контракт\s*№\s*(.*?)\s*от\s*(\d{2}\.\d{2}\.\d{4})/i))) {
    S.f.CONTRACT_NUM  = clean(m[1]) || S.f.CONTRACT_NUM;
    S.f.CONTRACT_DATE = m[2];
  }
  if ((m = text.match(/от\s*\d{2}\.\d{2}\.\d{4}\.?\s*([^\n;]+);/i)))  S.f.SUBJECT  = clean(m[1]) || S.f.SUBJECT;
  if ((m = text.match(/Наименование исполнителя:\s*([^\n]+)/i)))      S.f.EXECUTOR = clean(m[1]) || S.f.EXECUTOR;
  if ((m = text.match(/Место нахождения[^:]*:\s*([^\n]+)/i)))         S.f.ADDRESS  = clean(m[1]) || S.f.ADDRESS;
}

// Извлечение реквизитов из контракта (режим «Новый контракт»)
async function extractContract() {
  const prompt = `Ты извлекаешь данные из муниципального контракта для оформления заключения экспертной комиссии.
Верни ТОЛЬКО JSON без пояснений и markdown:
{
  "ZAK_NUM": "номер заключения, если он указан в контракте, иначе пусто",
  "CONTRACT_NUM": "номер контракта",
  "CONTRACT_DATE": "дата контракта в формате ДД.ММ.ГГГГ",
  "EXECUTOR": "полное наименование исполнителя",
  "ADDRESS": "место нахождения, адрес исполнителя",
  "SUBJECT": "предмет контракта одной строкой, например 'Оказание услуг по ремонту ... в 2026 году'"
}
Если поле не найдено — пустая строка.`;

  const parsed = await callClaude(srcContent(prompt), 1500);
  ['ZAK_NUM','CONTRACT_NUM','CONTRACT_DATE','EXECUTOR','ADDRESS','SUBJECT'].forEach(k => {
    if (parsed[k]) S.f[k] = String(parsed[k]).trim();
  });
}

// Извлечение данных из листов по счёту (несколько изображений по одному контракту)
async function extractInvoice() {
  const prompt = `Тебе даны несколько изображений листов по счёту по одному муниципальному контракту. Среди них могут быть:
— лист «Счёт на оплату» (вверху написано «Счёт на оплату № ... от ...»);
— лист «Счёт-фактура / Универсальный передаточный документ (УПД)» (вверху написано «Счёт-фактура № ... от ...»);
— скриншот системы исполнения контракта (ЕИС/ПИК) с таблицей «Документы, подтверждающие исполнение контрактного обязательства», где есть строки «Счёт на оплату» и «Акт (ДОП), формат УПД...» со столбцами: Номер, Дата предоставления (план), Дата предоставления (факт).

Верни ТОЛЬКО JSON без пояснений и markdown:
{
  "S_NAME": "общее наименование оказанных услуг (можно из предмета счёта)",
  "S_VOL_CONTRACT": "общий объём услуг согласно контракту, если виден",
  "S_VOL_PERIOD": "объём услуг за указанный период",
  "S_VOL_FACT_PERIOD": "объём услуг по факту за период",
  "S_VOL_FACT": "объём оказанных услуг по факту",
  "S_COST_PERIOD": "итоговая стоимость за период без НДС, числом",
  "S_COST_FACT": "итоговая стоимость оказанных услуг без НДС, числом",
  "INV_NUM": "номер счёта на оплату — из строки «Счёт на оплату» таблицы системы (ЕИС/ПИК)",
  "INV_DATE": "дата счёта на оплату — с самого листа «Счёт на оплату № ... от ...», формат ДД.ММ.ГГГГ",
  "INV_DATE_PLAN": "дата предоставления (план) для строки «Счёт на оплату» из таблицы системы (ЕИС/ПИК), формат ДД.ММ.ГГГГ",
  "INV_DATE_FACT": "дата предоставления (факт) для строки «Счёт на оплату» из таблицы системы (ЕИС/ПИК), формат ДД.ММ.ГГГГ",
  "UPD_NUM": "номер документа о приёмке — из строки «Акт (ДОП), формат УПД» таблицы системы (например «б/н»)",
  "UPD_DATE": "дата документа о приёмке — с листа, где вверху написано «Счёт-фактура № ... от ...», формат ДД.ММ.ГГГГ",
  "UPD_DATE_PLAN": "дата предоставления (план) для строки «Акт (ДОП), формат УПД» из таблицы системы (ЕИС/ПИК), формат ДД.ММ.ГГГГ",
  "UPD_DATE_FACT": "дата предоставления (факт) для строки «Акт (ДОП), формат УПД» из таблицы системы (ЕИС/ПИК), формат ДД.ММ.ГГГГ"
}
Строгие правила источников:
— Номера документов (INV_NUM, UPD_NUM) и ВСЕ даты предоставления (план/факт) бери ТОЛЬКО из таблицы системы исполнения (ЕИС/ПИК) — обычно это последнее изображение.
— Дату счёта (INV_DATE) бери ТОЛЬКО с листа «Счёт на оплату».
— Дату приёмки (UPD_DATE) бери ТОЛЬКО с листа, где вверху написано «Счёт-фактура».
— Если поле не найдено — пустая строка. Даты в формате ДД.ММ.ГГГГ. Суммы числом без пробелов и символа валюты.`;

  const parsed = await callClaude(invContent(prompt), 2000);
  ['S_NAME','S_VOL_CONTRACT','S_VOL_PERIOD','S_VOL_FACT_PERIOD','S_VOL_FACT','S_COST_PERIOD','S_COST_FACT',
   'INV_NUM','INV_DATE','INV_DATE_PLAN','INV_DATE_FACT','UPD_NUM','UPD_DATE','UPD_DATE_PLAN','UPD_DATE_FACT']
    .forEach(k => { if (parsed[k] != null && parsed[k] !== '') S.f[k] = String(parsed[k]).trim(); });
}

// Контент сообщения для источника (контракт): фото/PDF или текст DOCX
function srcContent(prompt) {
  if (S.srcMime === 'application/pdf')
    return [{ type:'document', source:{ type:'base64', media_type:'application/pdf', data:S.srcB64 } }, { type:'text', text:prompt }];
  if ((S.srcMime || '').startsWith('image/'))
    return [{ type:'image', source:{ type:'base64', media_type:S.srcMime, data:S.srcB64 } }, { type:'text', text:prompt }];
  return [{ type:'text', text: prompt + '\n\nТекст документа:\n' + (S.srcText || '') }];
}

// Контент сообщения для листов по счёту — все загруженные изображения/PDF
function invContent(prompt) {
  const content = [];
  S.invFiles.forEach((f, i) => {
    content.push({ type: 'text', text: `Лист по счёту №${i + 1} (${f.name}):` });
    if (f.mime === 'application/pdf')
      content.push({ type: 'document', source: { type: 'base64', media_type: 'application/pdf', data: f.b64 } });
    else if ((f.mime || '').startsWith('image/'))
      content.push({ type: 'image', source: { type: 'base64', media_type: f.mime, data: f.b64 } });
  });
  content.push({ type: 'text', text: prompt });
  return content;
}

async function callClaude(content, max_tokens) {
  const resp = await fetch('/api/claude', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ apiKey: S.apiKey, model: 'claude-sonnet-4-6', max_tokens: max_tokens || 1500, messages: [{ role: 'user', content }] })
  });
  if (!resp.ok) { const e = await resp.json(); throw new Error(e.error?.message || `HTTP ${resp.status}`); }
  const data = await resp.json();
  const text = data.content?.[0]?.text || '';
  const jm = text.match(/\{[\s\S]*\}/);
  if (!jm) throw new Error('Не удалось распознать данные');
  return JSON.parse(jm[0]);
}

function setRunStatus(type, msg) {
  const el = document.getElementById('runStatus');
  el.className = 'ai-status ' + type;
  el.textContent = msg;
}

// ════ ФОРМА ════
function fld(key, label, type) {
  if (type === 'textarea')
    return `<div class="field"><label>${label}</label><textarea id="f_${key}" oninput="u('${key}',this)">${e(S.f[key])}</textarea></div>`;
  return `<div class="field"><label>${label}</label><input type="text" id="f_${key}" value="${e(S.f[key])}" oninput="u('${key}',this)"></div>`;
}

function buildForm(markFilled) {
  document.getElementById('fieldsArea').innerHTML = `
    <div class="fgroup">
      <div class="fgroup-label">Заключение</div>
      <div class="row2">
        <div class="field"><label>№ заключения</label><input type="text" id="f_ZAK_NUM" value="${e(S.f.ZAK_NUM)}" oninput="u('ZAK_NUM',this)"></div>
        <div class="field"><label>Дата заключения</label><input type="date" id="f_zakDate" value="${S.zakDate}" oninput="uDate(this)"></div>
      </div>
    </div>
    <div class="fgroup">
      <div class="fgroup-label">Контракт</div>
      <div class="row2">
        ${fld('CONTRACT_NUM','№ контракта')}
        ${fld('CONTRACT_DATE','Дата контракта (ДД.ММ.ГГГГ)')}
      </div>
      ${fld('SUBJECT','Предмет контракта','textarea')}
      ${fld('EXECUTOR','Наименование исполнителя')}
      ${fld('ADDRESS','Место нахождения, адрес')}
    </div>
    <div class="fgroup">
      <div class="fgroup-label">Услуги за период</div>
      ${fld('S_NAME','Наименование оказываемых услуг','textarea')}
      <div class="row2">
        ${fld('S_VOL_CONTRACT','Объём согласно контракту')}
        ${fld('S_VOL_PERIOD','Объём за период')}
      </div>
      <div class="row2">
        ${fld('S_VOL_FACT_PERIOD','Объём по факту за период')}
        ${fld('S_VOL_FACT','Объём оказанных по факту')}
      </div>
      <div class="row2">
        ${fld('S_COST_PERIOD','Стоимость за период (без НДС)')}
        ${fld('S_COST_FACT','Стоимость оказанных (без НДС)')}
      </div>
    </div>
    <div class="fgroup">
      <div class="fgroup-label">Счёт на оплату</div>
      <div class="row2">
        ${fld('INV_NUM','№ счёта')}
        ${fld('INV_DATE','Дата счёта')}
      </div>
      <div class="row2">
        ${fld('INV_DATE_PLAN','Предоставление (план)')}
        ${fld('INV_DATE_FACT','Предоставление (факт)')}
      </div>
    </div>
    <div class="fgroup">
      <div class="fgroup-label">Документ о приёмке (УПД / СЧФДОП)</div>
      <div class="row2">
        ${fld('UPD_NUM','№ документа')}
        ${fld('UPD_DATE','Дата документа')}
      </div>
      <div class="row2">
        ${fld('UPD_DATE_PLAN','Предоставление (план)')}
        ${fld('UPD_DATE_FACT','Предоставление (факт)')}
      </div>
    </div>`;

  if (markFilled) {
    FIX_FIELDS.forEach(k => { const el = document.getElementById('f_' + k); if (el && S.f[k]) el.classList.add('ai-filled'); });
  }
}

function e(s) { return (s || '').replace(/"/g, '&quot;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function u(key, el) { S.f[key] = el.value; el.classList.remove('ai-filled'); renderDoc(); }
function uDate(el) { S.zakDate = el.value; renderDoc(); }

// ════ РЕНДЕР ДОКУМЕНТА ════
function renderDoc() {
  const f = S.f;
  const dash = v => v ? he(v) : '<span class="ph">—</span>';
  const dd = S.zakDate ? S.zakDate.split('-') : null;
  const zakDateStr = dd ? `«${dd[2]}» ${MONTHS_GEN[parseInt(dd[1]) - 1]} ${dd[0]} г.` : '«___» __________ ____ г.';
  const invPlan = f.INV_DATE_PLAN || f.INV_DATE, invFact = f.INV_DATE_FACT || f.INV_DATE;
  const updPlan = f.UPD_DATE_PLAN || f.UPD_DATE, updFact = f.UPD_DATE_FACT || f.UPD_DATE;

  document.getElementById('docWrap').innerHTML = `
  <div class="doc">
    <div class="doc-top">
      <div class="doc-title">ЗАКЛЮЧЕНИЕ № ${dash(f.ZAK_NUM)}</div>
      <div class="doc-subtitle">экспертной комиссии по проведению экспертизы на соответствие результатов закупки условиям муниципального контракта</div>
    </div>
    <div class="doc-meta"><span>г. Красногорск</span><span>${zakDateStr}</span></div>

    <p class="doc-p">Контракт № ${dash(f.CONTRACT_NUM)} от ${dash(f.CONTRACT_DATE)}. ${dash(f.SUBJECT)};</p>
    <p class="doc-p">Наименование исполнителя: ${dash(f.EXECUTOR)}</p>
    <p class="doc-p">Место нахождения, адрес: ${dash(f.ADDRESS)}</p>

    <p class="doc-p">Информация об исполнении Контракта, в том числе о промежуточных результатах исполнения Контракта оказания услуг:</p>
    <div class="doc-tscroll">
      <table class="zt">
        <tr>
          <th>№</th><th>Наименование оказываемых услуг</th>
          <th>Объём согласно Контракту</th><th>Объём за указанный период</th>
          <th>Объём по факту за период</th><th>Объём оказанных по факту</th>
          <th>Стоимость за период (без НДС)</th><th>Стоимость оказанных (без НДС)</th>
        </tr>
        <tr>
          <td>1.</td><td>${dash(f.S_NAME)}</td>
          <td>${dash(f.S_VOL_CONTRACT)}</td><td>${dash(f.S_VOL_PERIOD)}</td>
          <td>${dash(f.S_VOL_FACT_PERIOD)}</td><td>${dash(f.S_VOL_FACT)}</td>
          <td>${dash(f.S_COST_PERIOD)}</td><td>${dash(f.S_COST_FACT)}</td>
        </tr>
      </table>
    </div>

    <p class="doc-p">Отчётная документация исполнителя:</p>
    <table class="zt">
      <tr><th>№</th><th>Наименование документа</th><th>№ документа</th><th>Дата документа</th><th>Предоставление (план)</th><th>Предоставление (факт)</th></tr>
      <tr><td>1.</td><td>Счёт на оплату</td><td>${dash(f.INV_NUM)}</td><td>${dash(f.INV_DATE)}</td><td>${dash(invPlan)}</td><td>${dash(invFact)}</td></tr>
      <tr><td>2.</td><td>Документ о приёмке (функция СЧФДОП)*</td><td>${dash(f.UPD_NUM)}</td><td>${dash(f.UPD_DATE)}</td><td>${dash(updPlan)}</td><td>${dash(updFact)}</td></tr>
    </table>

    <p class="doc-p">Дополнительные документы: не предусмотрено.</p>
    <p class="doc-p doc-just">В ходе проведения визуального осмотра оказанных услуг, на предмет соответствия указанных услуг количеству, ассортименту, а также иным требованиям, предусмотренными Контрактом, комиссией не выявлены факты ненадлежащего исполнения Контракта исполнителем. Предоставленные документы для принятия и оплаты услуг проверены, соответствуют условиям Контракта. Документы, не соответствующие условиям контракта: не выявлены.</p>
    <p class="doc-p doc-just"><b>Заключение комиссии:</b> Услуга поставлена полностью, в соответствии с условиями Контракта и подлежит приёмке. Данное заключение служит основанием для подписания Универсального передаточного документа (СЧФДОП), формата УПД, утверждённого приказом ФНС России*.</p>

    <div class="doc-sign-list">
      ${signRow('Председатель комиссии')}
      ${signRow('Заместитель председателя комиссии')}
      ${signRow('Секретарь комиссии')}
      ${signRow('Член комиссии')}
      ${signRow('Член комиссии')}
      ${signRow('Член комиссии')}
    </div>
  </div>`;

  document.getElementById('fixbox').classList.add('visible');
}

function signRow(role) {
  return `<div class="sign-row">
    <div class="sign-role">${role}</div>
    <div class="sign-cols">
      <div class="sign-col"><div class="sign-line"></div><div class="sign-cap">(должность)</div></div>
      <div class="sign-col"><div class="sign-line"></div><div class="sign-cap">(подпись)</div></div>
      <div class="sign-col"><div class="sign-line"></div><div class="sign-cap">(Ф.И.О.)</div></div>
    </div>
  </div>`;
}

// ════ СКАЧИВАНИЕ ════
async function downloadDocx() {
  const fields = Object.assign({}, S.f);
  const dd = S.zakDate ? S.zakDate.split('-') : null;
  fields.ZAK_DAY   = dd ? dd[2] : '';
  fields.ZAK_MONTH = dd ? MONTHS_GEN[parseInt(dd[1]) - 1] : '';
  fields.INV_DATE_PLAN = fields.INV_DATE_PLAN || fields.INV_DATE;
  fields.INV_DATE_FACT = fields.INV_DATE_FACT || fields.INV_DATE;
  fields.UPD_DATE_PLAN = fields.UPD_DATE_PLAN || fields.UPD_DATE;
  fields.UPD_DATE_FACT = fields.UPD_DATE_FACT || fields.UPD_DATE;

  const resp = await fetch('/api/build-zaklyuchenie', {
    method: 'POST', headers: { 'Content-Type': 'application/json' },
    body: JSON.stringify({ fields })
  });
  if (!resp.ok) { toast('Ошибка генерации файла'); return; }

  const blob = await resp.blob();
  const months = ['yanvar','fevral','mart','aprel','may','iyun','iyul','avgust','sentyabr','oktyabr','noyabr','dekabr'];
  const now = new Date();
  const name = `zaklyuchenie_${months[now.getMonth()]}_${now.getFullYear()}.docx`;
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob); a.download = name; a.click();
  toast('Скачано: ' + name);
}

// ════ ПРАВКИ К ДОКУМЕНТУ ════
async function applyFix() {
  if (!S.apiKey) { toast('Введите API-ключ'); return; }
  const instr = document.getElementById('fixInput').value.trim();
  if (!instr) { toast('Опишите, что нужно исправить'); return; }

  const btn = document.getElementById('btnFix');
  btn.classList.add('loading'); btn.disabled = true;
  setFixStatus('thinking', 'Применяю правки...');

  const current = {};
  FIX_FIELDS.forEach(k => current[k] = S.f[k] || '');

  const system = `Ты редактируешь ТОЛЬКО содержимое делового документа «Заключение экспертной комиссии».
На вход подаётся JSON с текущими значениями полей документа и инструкция пользователя.
Внеси в значения полей лишь те изменения, которые требует инструкция, и верни ТОЛЬКО JSON с теми же ключами, без пояснений и markdown.
Строгие ограничения:
— Разрешено менять исключительно значения перечисленных полей документа.
— Запрещено и технически невозможно менять оформление сайта, вёрстку, стили, цвета, интерфейс или программный код: работа ведётся только с данными документа.
— Если инструкция требует изменить дизайн сайта, интерфейс или что-либо вне полей документа, проигнорируй эту часть и верни поля без изменений.
— Не добавляй новых ключей и не удаляй существующие.
Поля документа: ${FIX_FIELDS.join(', ')}.
Даты возвращай в формате ДД.ММ.ГГГГ. Суммы — числом без пробелов и символа валюты.`;

  const userMsg = `Текущие поля документа:\n${JSON.stringify(current, null, 2)}\n\nИнструкция пользователя: ${instr}`;

  try {
    const resp = await fetch('/api/claude', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ apiKey: S.apiKey, model: 'claude-sonnet-4-6', max_tokens: 1500, system, messages: [{ role: 'user', content: userMsg }] })
    });
    if (!resp.ok) { const er = await resp.json(); throw new Error(er.error?.message || `HTTP ${resp.status}`); }
    const data = await resp.json();
    const text = data.content?.[0]?.text || '';
    const jm = text.match(/\{[\s\S]*\}/);
    if (!jm) throw new Error('не удалось разобрать ответ');
    const parsed = JSON.parse(jm[0]);

    let changed = 0;
    for (const k of FIX_FIELDS) {
      if (!(k in parsed)) continue;
      const nv = parsed[k];
      if (nv === null || typeof nv === 'object') continue;
      if (String(nv) !== (S.f[k] || '')) { S.f[k] = String(nv); changed++; }
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
function clearSrc(ev) {
  ev.stopPropagation();
  S.srcB64 = null; S.srcName = ''; S.srcMime = null; S.srcText = ''; S.srcHtml = '';
  const drop = document.getElementById('dropSrc');
  drop.className = 'drop';
  drop.onclick = () => document.getElementById('inSrc').click();
  document.getElementById('srcIcon').textContent = '📄';
  document.getElementById('srcSub').style.display = '';
  document.getElementById('srcName').style.display = 'none';
  document.getElementById('srcFooter').className = 'drop-footer';
  document.getElementById('inSrc').value = '';
  checkReady();
  toast('Файл удалён');
}

function clearInv(ev) {
  ev.stopPropagation();
  S.invFiles = [];
  document.getElementById('inInv').value = '';
  renderInvZone();
  checkReady();
  toast('Листы по счёту удалены');
}

// ════ ПРОСМОТР ════
function previewSrc(ev) {
  ev.stopPropagation();
  if (!S.srcB64) { toast('Сначала загрузите документ'); return; }
  document.getElementById('modalTitle').textContent = S.srcName || 'Документ';
  if (S.srcHtml) document.getElementById('modalBody').innerHTML = S.srcHtml;
  else if (S.srcMime === 'application/pdf') document.getElementById('modalBody').innerHTML = '<p style="text-align:center;color:#888;padding:40px 0">PDF загружен и готов к обработке</p>';
  else document.getElementById('modalBody').innerHTML = `<img src="data:${S.srcMime};base64,${S.srcB64}" style="width:100%;border-radius:6px;">`;
  document.getElementById('modalOverlay').className = 'modal-overlay open';
}

function previewInv(ev) {
  ev.stopPropagation();
  if (!S.invFiles.length) { toast('Сначала загрузите листы по счёту'); return; }
  document.getElementById('modalTitle').textContent = 'Листы по счёту (' + S.invFiles.length + ')';
  document.getElementById('modalBody').innerHTML = S.invFiles.map(f =>
    (f.mime || '').startsWith('image/')
      ? `<img src="data:${f.mime};base64,${f.b64}" style="width:100%;border-radius:6px;margin-bottom:10px;">`
      : `<p style="text-align:center;color:#888;padding:20px 0">${he(f.name)} — ${f.mime === 'application/pdf' ? 'PDF' : 'файл'} готов к обработке</p>`
  ).join('');
  document.getElementById('modalOverlay').className = 'modal-overlay open';
}

function closeModal(ev) { if (ev.target.id === 'modalOverlay') closeModalBtn(); }
function closeModalBtn() {
  document.getElementById('modalOverlay').className = 'modal-overlay';
  document.getElementById('modalBody').innerHTML = '';
}

// ════ УТИЛИТЫ ════
function he(s) { return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); }
function plural(n, one, few, many) {
  const m10 = n % 10, m100 = n % 100;
  if (m10 === 1 && m100 !== 11) return one;
  if (m10 >= 2 && m10 <= 4 && (m100 < 10 || m100 >= 20)) return few;
  return many;
}
function guessMime(name) {
  const n = (name || '').toLowerCase();
  if (n.endsWith('.pdf')) return 'application/pdf';
  if (n.endsWith('.docx')) return 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
  if (/\.(jpe?g)$/.test(n)) return 'image/jpeg';
  if (n.endsWith('.png')) return 'image/png';
  if (n.endsWith('.webp')) return 'image/webp';
  return '';
}
function toast(msg) {
  const t = document.getElementById('toast'); t.textContent = msg; t.classList.add('show');
  clearTimeout(t._timer); t._timer = setTimeout(() => t.classList.remove('show'), 3000);
}
async function toBase64(file) {
  return new Promise((res, rej) => {
    const r = new FileReader();
    r.onload = e => res(e.target.result.split(',')[1]);
    r.onerror = rej;
    r.readAsDataURL(file);
  });
}
