/* WMS snapshots are physical stock; sales never decrement them a second time. */
(() => {
  const root = document.getElementById('inventoryFlow');
  if (!root) return;
  const $ = id => document.getElementById(id);
  const clean = v => String(v ?? '').trim().toUpperCase();
  const escape = v => String(v ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
  const number = v => Number(v).toLocaleString('ko-KR');
  let snapshot, ledger;
  const label = p => p.name + ' · ' + (p.barcode || '') + ' [' + p.code + ']';
  const within = (d, from, to) => (!from || d >= from) && (!to || d <= to);
  function render() {
    if (!snapshot) return;
    const selected = snapshot.products.find(p => label(p) === $('flowProduct').value);
    const from = $('flowStart').value, to = $('flowEnd').value;
    if (from && to && from > to) {
      $('flowRows').innerHTML = '<tr><td colspan="9" class="flow-state">조회 기간을 확인해 주세요.</td></tr>';
      return;
    }
    if (!selected) {
      $('flowRows').innerHTML = '<tr><td colspan="9" class="flow-state">상품 미선택</td></tr>';
      return;
    }
    const key = selected.key;
    const moves = ledger.rows.filter(r => r.sku === key && within(r.date, from, to));
    const coverage = ledger.coverage;
    const covered = Boolean(coverage && from && to && from >= coverage.from && to <= coverage.to);
    const units = kind => {
      const list = moves.filter(r => r.type === kind);
      const sum = list.reduce((a, r) => a + Number(r.qty), 0);
      return list.length ? number(sum) + (covered ? '' : ' (일부)') : covered ? '0' : '미수집';
    };
    const marketing = moves.filter(r => r.type === 'marketing' && r.confirmed === true && r.approved_by);
    const marketingQty = marketing.reduce((a, r) => a + Number(r.qty), 0);
    const duplicateBarcode = selected.barcode && snapshot.products.filter(p => clean(p.barcode) === clean(selected.barcode)).length > 1;
    const candidates = rawRows.filter(r => !duplicateBarcode && selected.barcode && clean(r.stock_barcode || r.match_sku) === clean(selected.barcode) && r.source_type !== '정산조정');
    const ambiguous = r => (r.stock_barcode && r.match_sku && clean(r.stock_barcode) !== clean(r.match_sku)) ||
      (r.option_quantities && Object.keys(r.option_quantities).length > 1);
    const sales = candidates.filter(r => !ambiguous(r));
    const pending = candidates.some(r => ambiguous(r) || r.qty_verified === false || r.options_verified === false || (r.daily || []).some(d => within(d.date,from,to) && d.qty_verified === false));
    const saleQty = sales.reduce((sum, r) => sum + (r.daily || []).filter(d => within(d.date,from,to)).reduce((a,d) => a + Number(d.qty || 0),0),0);
    const dates = rawRows.flatMap(r => (r.daily || []).map(d => d.date)).filter(Boolean).sort();
    $('flowPeriod').textContent = (from || dates[0] || '全期間') + ' ~ ' + (to || dates[dates.length-1] || '全期間');
    const saleText = duplicateBarcode ? '바코드 중복 확인 필요' : !sales.length ? '미연결' : pending ? number(saleQty) + ' (일부 대사 중)' : number(saleQty);
    const cols = [
      escape(selected.name), escape(units('in')), escape(units('out')),
      marketing.length ? number(marketingQty) : '<span class="flow-state">확인 필요</span>',
      escape(saleText), '<span class="flow-stock">' + number(selected.stock) + '</span>',
      number(selected.available), number(selected.instructed), number(selected.processing)
    ];
    $('flowRows').innerHTML = '<tr>' + cols.map(c => '<td>' + c + '</td>').join('') + '</tr>';
    $('flowSku').textContent = 'SKU ' + selected.code + (selected.barcode ? ' · ' + selected.barcode : '');
  }
  async function init() {
    try {
      const responses = await Promise.all([fetch('/inventory_snapshot.json', {cache:'no-store'}), fetch('/inventory_movements.json', {cache:'no-store'})]);
      if (responses.some(r => !r.ok)) throw new Error('inventory source unavailable');
      [snapshot, ledger] = await Promise.all(responses.map(r => r.json()));
      if (!Array.isArray(snapshot.products) || !Array.isArray(ledger.rows)) throw new Error('invalid inventory source');
      $('flowOptions').innerHTML = snapshot.products.map(p => '<option value="' + escape(label(p)) + '"></option>').join('');
      $('flowAsOf').textContent = 'WMS 총재고 ' + number(snapshot.products.reduce((n,p)=>n+p.stock,0)) + ' · ' + number(snapshot.products.length) + '개 상품 · ' + snapshot.captured_at.replace('T',' ').replace('+09:00','') + ' 기준';
      $('flowSource').textContent = snapshot.source;
      const dates = rawRows.flatMap(r => (r.daily || []).map(d => d.date)).filter(Boolean).sort();
      $('flowStart').value = dates[0] || '';
      $('flowEnd').value = dates[dates.length-1] || '';
      for (const id of ['flowProduct','flowStart','flowEnd']) {
        $(id).addEventListener('input',render);
        $(id).addEventListener('change',render);
      }
      render();
    } catch (e) {
      $('flowRows').innerHTML = '<tr><td colspan="9" class="flow-state">재고 원본 불러오기 실패</td></tr>';
    }
  }
  init();
})();
