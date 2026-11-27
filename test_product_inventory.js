const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const snapshot = {
  source:'stock.xlsx',captured_at:'2026-11-27T13:26:56+09:00',
  products:[
    {key:'code:A',code:'A',barcode:'SKU1',name:'Shirt <Blue>',stock:10,available:8,instructed:2,processing:0},
    {key:'code:B',code:'B',barcode:'DUP',name:'Duplicate',stock:0,available:0,instructed:0,processing:0},
    {key:'code:C',code:'C',barcode:'DUP',name:'Duplicate',stock:9,available:9,instructed:0,processing:0}
  ]
};
const ledger={coverage:{from:'2026-11-01',to:'2026-11-27'},rows:[
  {sku:'code:A',date:'2026-11-26',type:'in',qty:4},
  {sku:'code:A',date:'2026-11-26',type:'out',qty:2},
  {sku:'code:A',date:'2026-11-26',type:'marketing',qty:1,confirmed:true,approved_by:'owner'},
  {sku:'code:A',date:'2026-11-26',type:'marketing',qty:100,confirmed:false}
]};
const nodes = {};
function el(id){return nodes[id] ||= {value:'',textContent:'',innerHTML:'',events:{},addEventListener(e,f){this.events[e]=f}}}
const rawRows=[{stock_barcode:'SKU1',source_type:'상품',daily:[
  {date:'2026-11-25',qty:3},{date:'2026-11-26',qty:-1}
]}];
async function main(){
  const context={document:{getElementById:el},rawRows,fetch:async url=>({ok:true,json:async()=>url.includes('snapshot')?snapshot:ledger})};
  vm.runInNewContext(fs.readFileSync('product_inventory.js','utf8'),context);
  await new Promise(r=>setImmediate(r));
  assert.match(el('flowRows').innerHTML,/상품 미선택/);
  el('flowProduct').value='Shirt <Blue> · SKU1 [A]';
  el('flowStart').value='2026-11-26';el('flowEnd').value='2026-11-26';
  const render=el('flowProduct').events.change;
  render();
  assert.match(el('flowRows').innerHTML,/Shirt &lt;Blue&gt;/);
  assert.match(el('flowRows').innerHTML,/<td>4<\/td><td>2<\/td><td>1<\/td><td>-1<\/td>/);
  assert.match(el('flowRows').innerHTML,/flow-stock">10/);
  assert.equal(snapshot.products[0].stock,10,'Sales must not decrement physical stock');
  ledger.coverage=null;ledger.rows=[];render();
  assert.match(el('flowRows').innerHTML,/미수집/);
  assert.match(el('flowRows').innerHTML,/확인 필요/);
  el('flowProduct').value='Duplicate · DUP [B]';render();
  assert.match(el('flowRows').innerHTML,/바코드 중복 확인 필요/);
  assert.match(el('flowRows').innerHTML,/flow-stock">0/);
  el('flowStart').value='2026-12-01';render();
  assert.match(el('flowRows').innerHTML,/조회 기간/);
  const h=fs.readFileSync('index.html','utf8');
  assert.equal(h,fs.readFileSync('public/index.html','utf8'));
  assert.ok(h.indexOf('inventoryFlow')<h.indexOf('</main>'));
  for(const m of h.matchAll(/<script\b([^>]*)>([\s\S]*?)<\/script>/gi)){
    if(!/src=|application\/json|module/.test(m[1]))new vm.Script(m[2]);
  }
  const actual=JSON.parse(fs.readFileSync('inventory_snapshot.json','utf8'));
  assert.equal(actual.products.length,1297);
  assert.equal(new Set(actual.products.map(p=>p.key)).size,1297);
  assert.equal(actual.products.reduce((s,p)=>s+p.stock,0),179875);
  assert.equal(actual.basis,'총재고');
  console.log('PASS: inventory selection, dates, refunds, confirmation, duplicate SKUs, XSS, zero stock, snapshot totals, HTML and scripts');
}
main().catch(e=>{console.error(e);process.exitCode=1});
