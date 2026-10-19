const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const html = fs.readFileSync(path.join(__dirname, 'index.html'), 'utf8');
assert.equal(html.replace(/\r\n/g, '\n'), fs.readFileSync(path.join(__dirname, 'public/index.html'), 'utf8').replace(/\r\n/g, '\n'));
for (const match of html.matchAll(/<script([^>]*)>([\s\S]*?)<\/script>/g)) {
  if (!match[1].includes('src=') && match[2].trim()) new vm.Script(match[2]);
}
const read = name => JSON.parse(html.match(new RegExp('^const ' + name + ' = (.+);$', 'm'))[1]);
const code = html.slice(html.indexOf('function buildDailyItemRows('), html.indexOf('function renderDailyItems('));
const dailyCode = html.slice(html.indexOf('function buildDailyMap('), html.indexOf('function buildHistoricalDailyMap('));
const context = vm.createContext({rawRows: read('rawRows'), MANUAL_SALES: read('MANUAL_SALES')});
vm.runInContext(code + dailyCode, context);
const sum = (rows, field) => rows.reduce((s, r) => s + r[field], 0);
const check = (retailer, start, end) => {
  const items = context.buildDailyItemRows(retailer, start, end, '2026');
  const summary = context.buildDailyMap(retailer).filter(d => (!start || d.date >= start) && (!end || d.date <= end));
  for (const key of ['qty', 'payment']) assert.equal(sum(items, key), sum(summary, key), [retailer, start, end, key].join(' '));
  return items;
};
for (const retailer of ['', ...new Set(context.rawRows.map(r => r.retailer))]) {
  check(retailer, '', '');
  check(retailer, '2026-10-16', '2026-10-18');
  check(retailer, '2026-10-16', '2026-10-16');
  check(retailer, '2026-10-17', '2026-10-17');
  check(retailer, '2026-10-18', '2026-10-18');
}
const weekend = check('', '2026-10-16', '2026-10-18');
assert.equal(sum(weekend, 'qty'), 36);
assert.equal(sum(weekend, 'payment'), 1023180);
assert.ok(weekend.some(r => r.size === 'XXL' && r.qty < 0));
assert.ok(weekend.some(r => r.retailer === '무신사' && !r.size));
assert.ok(weekend.some(r => r.kind === 'adjustment' && r.qty === 0));
assert.equal(context.buildDailyItemRows('', '', '', '2025').length, 0);
assert.equal(context.buildDailyItemRows('', '2026-10-18', '2026-10-16', '2026').length, 0);
assert.equal(context.escapeDailyItemText('<img onerror="bad">&'), '&lt;img onerror=&quot;bad&quot;&gt;&amp;');

const daily = (date, qty, payment) => ({date, qty, payment, gross: payment, orders: 1});
const date = '2026-10-01';
context.rawRows = [
  {retailer: 'A', name: 'Pack', color: 'Black', size: 'XXL', daily: [daily(date, 2, 100), daily('2026-10-02', -1, -50)]},
  {retailer: 'B', name: 'Pack', color: 'Black', size: 'XXL', daily: [daily(date, 1, 70)]},
  {retailer: 'A', name: 'Shipping', source_type: '정산조정', daily: [daily(date, 0, 10)]},
];
context.MANUAL_SALES = [{date, retailers: [
  {retailer: 'A', qty: 2, payment: 110, items: [{name: 'Do not duplicate', qty: 2, payment: 110}]},
  {retailer: 'C', qty: 3, payment: 90, items: [{name: 'Unknown options', qty: 3, payment: 90}]},
  {retailer: 'D', qty: 1, payment: 40},
]}];
const synthetic = context.buildDailyItemRows('', '', '', '2026');
assert.equal(sum(synthetic, 'qty'), 6);
assert.equal(sum(synthetic, 'payment'), 260);
assert.equal(synthetic.filter(r => r.name === 'Pack').length, 2);
assert.ok(!synthetic.some(r => r.name === 'Do not duplicate'));
assert.ok(synthetic.some(r => r.retailer === 'C' && r.kind === 'unverified' && !r.size));
assert.ok(synthetic.some(r => r.retailer === 'D' && r.name === '상품 상세 미확인'));
context.MANUAL_SALES[0].retailers[0].qty = 3;
context.MANUAL_SALES[0].retailers[0].payment = 130;
const partial = context.buildDailyItemRows('A', date, date, '2026');
assert.equal(sum(partial, 'qty'), 3);
assert.equal(sum(partial, 'payment'), 130);
assert.ok(partial.some(r => r.kind === 'unverified' && r.qty === 1 && r.payment === 20));
console.log('PASS: generated HTML, daily/channel totals, refunds, adjustments, unknown options, manual deduplication, date/year filters, escaping.');
