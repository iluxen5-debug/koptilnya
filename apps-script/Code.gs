// Приём заявок с формы клиента в Google Таблицу.
// 1) Замените SECRET на свою длинную фразу — ту же впишите в ⚙ на странице заказов.
// 2) Развернуть → Новое развертывание → Веб-приложение, «У кого есть доступ»: Все.
const SECRET = 'ЗАМЕНИТЕ_НА_СВОЙ_СЕКРЕТ';  // правьте только эту строку, не через «Заменить всё»
// Until SECRET is changed, reading orders and changing the catalog stay closed
const SECRET_SET = SECRET.length >= 8 && !/^ЗАМЕНИТЕ/.test(SECRET);

const HEAD = ['id', 'Время', 'Имя', 'Телефон', 'Адрес', 'Дата доставки', 'Товары', 'Комментарий', 'json', 'Забрано'];
const UNITS = ['кг', 'шт', 'уп'];

function sheet() {
  const ss = SpreadsheetApp.getActive();
  let s = ss.getSheetByName('Заявки');
  if (!s) { s = ss.insertSheet('Заявки'); s.appendRow(HEAD); }
  return s;
}
const out = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
const str = (v, n) => String(v == null ? '' : v).trim().slice(0, n);
// A cell starting with = + - @ would run as a formula
const cell = v => /^[=+\-@]/.test(v) ? "'" + v : v;

function doGet(e) {
  const p = e.parameter;
  if (p.action === 'catalog') return out(JSON.parse(PropertiesService.getScriptProperties().getProperty('catalog') || '[]'));
  if (p.action === 'orders' && SECRET_SET && p.token === SECRET) {
    const rows = sheet().getDataRange().getValues().slice(1).filter(r => r[0] && !r[9]);
    return out(rows.map(r => Object.assign({ id: String(r[0]) }, JSON.parse(r[8]))));
  }
  return out({ error: 'bad request' });
}

function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return out({ error: 'bad request' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    if (d.action === 'order') {
      const items = (Array.isArray(d.items) ? d.items : []).slice(0, 30)
        .map(i => ({ name: str(i.name, 80), qty: Math.min(+i.qty || 0, 1000), unit: UNITS.includes(i.unit) ? i.unit : 'шт' }))
        .filter(i => i.name && i.qty > 0);
      const o = {
        name: str(d.name, 80), phone: str(d.phone, 30), address: str(d.address, 200),
        date: /^\d{4}-\d\d-\d\d$/.test(d.date) ? d.date : '', note: str(d.note, 500), items: items
      };
      if (!o.name || !o.phone || !items.length) return out({ error: 'Заполните имя, телефон и хотя бы один товар' });
      const id = 'f' + Date.now().toString(36) + Math.random().toString(36).slice(2, 6);
      sheet().appendRow([id, new Date(), cell(o.name), cell(o.phone), cell(o.address), o.date,
        cell(items.map(i => i.name + ' ' + i.qty + ' ' + i.unit).join(', ')), cell(o.note), JSON.stringify(o), '']);
      return out({ ok: true });
    }
    if (!SECRET_SET || d.token !== SECRET) return out({ error: 'forbidden' });
    if (d.action === 'catalog') {
      const cat = (Array.isArray(d.catalog) ? d.catalog : []).slice(0, 200)
        .map(c => ({ name: str(c.name, 80), price: +c.price || 0, unit: UNITS.includes(c.unit) ? c.unit : 'кг' })).filter(c => c.name);
      PropertiesService.getScriptProperties().setProperty('catalog', JSON.stringify(cat));
      return out({ ok: true });
    }
    if (d.action === 'taken') {
      const s = sheet(), v = s.getDataRange().getValues(), ids = new Set((d.ids || []).map(String));
      v.forEach((r, k) => { if (k && ids.has(String(r[0])) && !r[9]) s.getRange(k + 1, 10).setValue(new Date()); });
      return out({ ok: true });
    }
    return out({ error: 'bad request' });
  } finally {
    lock.releaseLock();
  }
}
