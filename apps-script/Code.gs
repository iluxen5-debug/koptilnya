// База заказов коптильни в этой Google Таблице + приём заявок с формы клиента.
// 1) Впишите логин и пароль ниже (правьте только эти строки, не через «Заменить всё»).
// 2) Развернуть → Управление развертываниями → ✏️ → Версия: Новая версия → Развернуть
//    (первый раз: Новое развертывание → Веб-приложение, запуск от меня, доступ: Все).
const LOGIN = 'admin';
const PASSWORD = 'ЗАМЕНИТЕ_НА_СВОЙ_ПАРОЛЬ';
// Уведомления о заявках с формы (оставьте пустым, что не нужно):
const NOTIFY_EMAIL = '';   // ваша почта, бесплатно
const SMSRU_API_ID = '';   // ключ API с sms.ru (платно, ~3–5 ₽ за SMS)
const SMS_TO = '';         // номер для SMS, например 79171234567
// Until the password is changed, the base stays closed
const READY = PASSWORD.length >= 8 && !/^ЗАМЕНИТЕ/.test(PASSWORD);

const UNITS = ['кг', 'шт', 'уп'];
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const str = (v, n) => String(v == null ? '' : v).trim().slice(0, n);
// A cell starting with = + - @ would run as a formula
const cell = v => /^[=+\-@]/.test(v) ? "'" + v : v;
const props = () => PropertiesService.getScriptProperties();

function sheet(name, head) {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(head); }
  return s;
}
// One order per row: id | изменён | клиент | json
const orders = () => sheet('Заказы', ['id', 'Изменён', 'Клиент', 'json']);
const requests = () => sheet('Заявки', ['Время', 'Имя', 'Телефон', 'Адрес', 'Дата доставки', 'Товары', 'Комментарий']);

function rowOf(s, id) {
  const f = s.getRange('A:A').createTextFinder(String(id)).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}
function put(o) {
  if (!o || !o.id) return;
  const s = orders(), id = str(o.id, 40), r = rowOf(s, id);
  const row = [id, new Date(), cell(str(o.name, 80)), JSON.stringify(o)];
  if (r > 1) s.getRange(r, 1, 1, 4).setValues([row]); else s.appendRow(row);
}

function doGet(e) {
  // Public: the order form reads product names and prices
  if (e.parameter.action === 'catalog') return out(JSON.parse(props().getProperty('catalog') || '[]'));
  return out({ error: 'bad request' });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out({ error: 'bad request' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (d.action === 'order') return out(formOrder(d));
    if (!READY || d.token !== LOGIN + ':' + PASSWORD) return out({ error: 'forbidden' });
    if (d.action === 'state') {
      const rows = orders().getDataRange().getValues().slice(1).filter(r => r[0] && r[3]);
      const p = props();
      return out({ ok: true, orders: rows.map(r => JSON.parse(r[3])), catalog: p.getProperty('catalogText') || null,
        settings: JSON.parse(p.getProperty('settings') || '{}') });
    }
    if (d.action === 'put') { (Array.isArray(d.orders) ? d.orders : []).forEach(put); return out({ ok: true }); }
    if (d.action === 'del') {
      const s = orders();
      (Array.isArray(d.ids) ? d.ids : []).forEach(id => { const r = rowOf(s, id); if (r > 1) s.deleteRow(r); });
      return out({ ok: true });
    }
    if (d.action === 'catalog') {
      const cat = (Array.isArray(d.catalog) ? d.catalog : []).slice(0, 300)
        .map(c => ({ name: str(c.name, 80), price: +c.price || 0, unit: UNITS.includes(c.unit) ? c.unit : 'кг' })).filter(c => c.name);
      props().setProperties({ catalog: JSON.stringify(cat), catalogText: str(d.text, 20000) });
      return out({ ok: true });
    }
    if (d.action === 'settings') {
      const s = d.settings || {};
      props().setProperty('settings', JSON.stringify({ key: str(s.key, 300), url: str(s.url, 300), model: str(s.model, 100) }));
      return out({ ok: true });
    }
    return out({ error: 'bad request' });
  } finally {
    lock.releaseLock();
  }
}

// A request from the customer form goes straight into the base as a new order
function formOrder(d) {
  const items = (Array.isArray(d.items) ? d.items : []).slice(0, 30)
    .map(i => ({ name: str(i.name, 80), qty: Math.min(+i.qty || 0, 1000), unit: UNITS.includes(i.unit) ? i.unit : 'шт', sum: 0 }))
    .filter(i => i.name && i.qty > 0);
  const o = {
    id: 'f' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6),
    name: str(d.name, 80), phone: str(d.phone, 30), address: str(d.address, 200),
    date: /^\d{4}-\d\d-\d\d$/.test(d.date) ? d.date : Utilities.formatDate(new Date(), 'Asia/Yekaterinburg', 'yyyy-MM-dd'),
    note: str(d.note, 500), raw: '', paid: false, status: 'new', doneAt: null, createdAt: Date.now(), updatedAt: Date.now(),
    from: 'form', items: items
  };
  if (!o.name || !o.phone || !items.length) return { error: 'Заполните имя, телефон и хотя бы один товар' };
  put(o);
  requests().appendRow([new Date(), cell(o.name), cell(o.phone), cell(o.address), o.date,
    cell(items.map(i => i.name + ' ' + i.qty + ' ' + i.unit).join(', ')), cell(o.note)]);
  notify(o);
  return { ok: true };
}

// A failed notification must not lose the order, so errors are only logged
function notify(o) {
  const list = o.items.map(i => i.name + ' ' + i.qty + ' ' + i.unit).join(', ');
  try {
    if (NOTIFY_EMAIL) MailApp.sendEmail(NOTIFY_EMAIL, 'Новая заявка: ' + o.name,
      [o.name, o.phone, o.address, 'Дата: ' + o.date, list, o.note].filter(Boolean).join('\n'));
  } catch (err) { console.error('email', err); }
  try {
    // One Cyrillic SMS fits 70 characters
    if (SMSRU_API_ID && SMS_TO) UrlFetchApp.fetch('https://sms.ru/sms/send?json=1&api_id=' + encodeURIComponent(SMSRU_API_ID) +
      '&to=' + encodeURIComponent(SMS_TO) + '&msg=' + encodeURIComponent(('Заявка: ' + o.name + ', ' + list).slice(0, 70)));
  } catch (err) { console.error('sms', err); }
}
