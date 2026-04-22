"""
build_product_dashboard.py
data.json 을 읽어 index.html 을 생성합니다.
"""
import json, datetime, pathlib, html as _html

DATA_FILE  = pathlib.Path("data.json")
OUT_FILE   = pathlib.Path("index.html")

# ── data.json 로드 ─────────────────────────────────────────────
with DATA_FILE.open(encoding="utf-8") as f:
    rows = json.load(f)          # list of row dicts

now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")

# ── KPI 계산 ───────────────────────────────────────────────────
total_qty      = sum(r.get("qty", 0)     for r in rows)
total_payment  = sum(r.get("payment", 0) for r in rows)
total_gross    = sum(r.get("gross", 0)   for r in rows)
total_orders   = sum(r.get("orders", 0)  for r in rows)
avg_unit       = int(total_payment / total_qty) if total_qty else 0
avg_discount   = round((1 - total_payment / total_gross) * 100, 1) if total_gross else 0
matched        = sum(1 for r in rows if r.get("match_status") == "매칭됨")
unmatched      = len(rows) - matched
match_rate     = round(matched / len(rows) * 100, 1) if rows else 0

rows_json = json.dumps(rows, ensure_ascii=False)

# ── HTML 템플릿 ────────────────────────────────────────────────
html = f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>상품DATA 운영 대시보드</title>
<link href="https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>
<script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
<style>
:root{{
  --navy:#1e2535;--navy2:#2b3548;--bg:#f0f2f6;--panel:#fff;
  --border:#e2e6ed;--border2:#d0d5df;--ink:#1e2535;--ink2:#4a5568;--ink3:#8a94a6;
  --blue:#3b82f6;--blue2:#2563eb;--blue-soft:#eff6ff;
  --teal:#10b981;--teal-soft:#ecfdf5;--amber:#f59e0b;--amber-soft:#fffbeb;
  --red:#ef4444;--red-soft:#fef2f2;
  --shadow:0 1px 3px rgba(0,0,0,0.08),0 1px 2px rgba(0,0,0,0.05);
  --shadow-md:0 4px 12px rgba(0,0,0,0.08),0 2px 4px rgba(0,0,0,0.04);
}}
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'Pretendard',-apple-system,sans-serif;background:var(--bg);color:var(--ink);min-height:100vh;}}
header{{background:var(--navy);position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,0.2);}}
.header-top{{display:flex;align-items:center;height:52px;padding:0 20px;gap:0;}}
.logo-area{{display:flex;align-items:center;gap:10px;min-width:156px;border-right:1px solid rgba(255,255,255,0.1);padding-right:20px;margin-right:4px;height:100%;}}
.logo-icon{{width:32px;height:32px;border-radius:8px;background:linear-gradient(135deg,#3b82f6,#6366f1);display:grid;place-items:center;font-size:13px;font-weight:800;color:#fff;flex-shrink:0;}}
.logo-text{{font-size:13px;font-weight:700;color:#fff;letter-spacing:-0.2px;line-height:1.3;}}
.logo-sub{{font-size:10px;color:rgba(255,255,255,0.4);font-weight:400;letter-spacing:0.5px;}}
.header-nav{{display:flex;align-items:center;height:100%;flex:1;overflow-x:auto;}}
.nav-item{{display:flex;flex-direction:column;align-items:flex-start;justify-content:center;height:100%;padding:0 18px;cursor:pointer;border-bottom:3px solid transparent;transition:all 0.15s;white-space:nowrap;}}
.nav-item:hover{{background:rgba(255,255,255,0.05);}}
.nav-item.active{{border-bottom-color:var(--blue);}}
.nav-label{{font-size:13px;font-weight:600;color:rgba(255,255,255,0.85);display:flex;align-items:center;gap:6px;}}
.nav-dot{{width:10px;height:10px;border-radius:2px;flex-shrink:0;}}
.nav-sub{{font-size:10px;color:rgba(255,255,255,0.38);margin-top:1px;}}
.header-right{{margin-left:auto;flex-shrink:0;}}
.status-chip{{display:flex;align-items:center;gap:5px;background:rgba(16,185,129,0.15);border:1px solid rgba(16,185,129,0.3);border-radius:20px;padding:4px 10px;font-size:11px;color:#6ee7b7;font-family:'DM Mono',monospace;}}
.live-dot{{width:5px;height:5px;border-radius:50%;background:#10b981;box-shadow:0 0 6px #10b981;animation:blink 2s infinite;}}
@keyframes blink{{0%,100%{{opacity:1}}50%{{opacity:.3}}}}
main{{padding:20px 20px 48px;max-width:1600px;margin:0 auto;}}
.page-hd{{display:flex;align-items:flex-end;justify-content:space-between;margin-bottom:16px;}}
.page-title{{font-size:20px;font-weight:800;color:var(--ink);letter-spacing:-0.4px;}}
.page-date{{font-size:12px;color:var(--ink3);margin-top:3px;}}
.kpis{{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:16px;}}
.kpi{{background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:16px;box-shadow:var(--shadow);transition:box-shadow 0.15s,transform 0.15s;}}
.kpi:hover{{box-shadow:var(--shadow-md);transform:translateY(-1px);}}
.kpi-label{{font-size:11px;font-weight:600;color:var(--ink3);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;}}
.kpi-value{{font-size:22px;font-weight:800;color:var(--ink);letter-spacing:-0.5px;font-family:'DM Mono',monospace;}}
.kpi-value.blue{{color:var(--blue2);}} .kpi-value.green{{color:var(--teal);}} .kpi-value.amber{{color:var(--amber);}}
.kpi-note{{font-size:11px;color:var(--ink3);margin-top:5px;}}
.kpi-bar{{height:3px;border-radius:2px;margin-top:10px;background:var(--border);overflow:hidden;}}
.kpi-bar div{{height:100%;border-radius:2px;}}
.panel{{background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:18px;box-shadow:var(--shadow);}}
.panel-hd{{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;}}
.panel-title{{font-size:14px;font-weight:700;color:var(--ink);letter-spacing:-0.2px;}}
.panel-meta{{font-size:11px;color:var(--ink3);font-family:'DM Mono',monospace;}}
.grid2{{display:grid;grid-template-columns:1.3fr 0.7fr;gap:12px;margin-bottom:12px;}}
.grid2-eq{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px;}}
.charts-row{{display:grid;grid-template-columns:1fr 1fr;gap:12px;}}
.chart-box{{position:relative;height:200px;}} .chart-box canvas{{width:100%!important;height:100%!important;}}
.rank-list{{display:flex;flex-direction:column;}}
.rank-row{{display:grid;grid-template-columns:22px minmax(0,1fr) 90px 80px;align-items:center;gap:8px;padding:5px 6px;border-radius:6px;transition:background 0.12s;}}
.rank-row:hover{{background:var(--bg);}}
.rank-n{{font-size:11px;color:var(--ink3);font-family:'DM Mono',monospace;}}
.rank-name{{font-size:12px;color:var(--ink2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}}
.rank-bar{{height:5px;background:var(--border);border-radius:3px;overflow:hidden;}}
.rank-bar div{{height:100%;border-radius:3px;background:linear-gradient(90deg,var(--blue),#818cf8);}}
.rank-val{{font-size:12px;color:var(--ink);text-align:right;font-family:'DM Mono',monospace;font-weight:600;}}
.mini-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;}}
.mini{{background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:11px 12px;}}
.mini-label{{font-size:10px;color:var(--ink3);text-transform:uppercase;letter-spacing:0.5px;font-weight:600;}}
.mini-value{{font-size:15px;font-weight:700;color:var(--ink);margin-top:4px;font-family:'DM Mono',monospace;}}
.sum-strip{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px;}}
.sum-card{{background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:12px 14px;}}
.sum-label{{font-size:11px;color:var(--ink3);font-weight:500;}}
.sum-value{{font-size:19px;font-weight:800;color:var(--ink);margin-top:3px;font-family:'DM Mono',monospace;}}
.toolbar{{display:flex;gap:8px;align-items:center;margin-bottom:12px;flex-wrap:wrap;}}
.toolbar input,.toolbar select{{height:34px;border:1px solid var(--border2);border-radius:7px;background:var(--panel);color:var(--ink);font-size:13px;padding:0 10px;font-family:'Pretendard',sans-serif;outline:none;min-width:140px;transition:border-color 0.15s,box-shadow 0.15s;}}
.toolbar input:focus,.toolbar select:focus{{border-color:var(--blue);box-shadow:0 0 0 3px rgba(59,130,246,0.1);}}
.toolbar input::placeholder{{color:var(--ink3);}}
.btn-sm{{height:34px;padding:0 14px;background:var(--panel);border:1px solid var(--border2);border-radius:7px;font-size:12px;font-weight:600;color:var(--ink2);cursor:pointer;font-family:'Pretendard',sans-serif;transition:all 0.15s;white-space:nowrap;}}
.btn-sm:hover{{background:var(--bg);color:var(--ink);}}
.table-wrap{{border:1px solid var(--border);border-radius:8px;overflow:auto;max-height:500px;}}
table{{border-collapse:collapse;width:100%;min-width:900px;font-size:12.5px;}}
thead{{position:sticky;top:0;z-index:2;}}
th{{background:#f8fafc;color:var(--ink3);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;padding:9px 10px;border-bottom:1px solid var(--border);text-align:left;white-space:nowrap;}}
td{{padding:9px 10px;border-bottom:1px solid var(--border);color:var(--ink2);vertical-align:middle;}}
tr:last-child td{{border-bottom:none;}} tr:hover td{{background:#fafbfd;}}
.num{{text-align:right;font-family:'DM Mono',monospace;font-size:12px;}}
.td-main{{color:var(--ink);font-weight:600;}} .td-mono{{font-family:'DM Mono',monospace;font-size:11.5px;color:var(--ink3);}}
.badge{{display:inline-flex;align-items:center;height:22px;padding:0 8px;border-radius:5px;font-size:11px;font-weight:600;white-space:nowrap;}}
.badge-blue{{background:var(--blue-soft);color:var(--blue2);}} .badge-green{{background:var(--teal-soft);color:#059669;}}
.badge-amber{{background:var(--amber-soft);color:#b45309;}} .badge-red{{background:var(--red-soft);color:#dc2626;}}
.badge-gray{{background:#f1f5f9;color:var(--ink3);}}
.tab-panel{{display:none;}} .tab-panel.active{{display:block;}}
.foot{{font-size:11px;color:var(--ink3);margin-top:10px;line-height:1.7;}}
::-webkit-scrollbar{{width:5px;height:5px;}} ::-webkit-scrollbar-track{{background:transparent;}}
::-webkit-scrollbar-thumb{{background:var(--border2);border-radius:3px;}}
@media(max-width:1200px){{.kpis{{grid-template-columns:repeat(3,1fr);}}.grid2,.grid2-eq,.charts-row{{grid-template-columns:1fr;}}}}
@media(max-width:680px){{main,.header-top{{padding-left:12px;padding-right:12px;}}.kpis{{grid-template-columns:1fr 1fr;}}}}
</style>
</head>
<body>
<header>
  <div class="header-top">
    <div class="logo-area">
      <div class="logo-icon">PD</div>
      <div><div class="logo-text">상품DATA</div><div class="logo-sub">운영 대시보드</div></div>
    </div>
    <nav class="header-nav">
      <div class="nav-item active" data-tab="overview">
        <div class="nav-label"><span class="nav-dot" style="background:#3b82f6"></span>요약</div>
        <div class="nav-sub">KPI · 상위 상품</div>
      </div>
      <div class="nav-item" data-tab="calendar">
        <div class="nav-label"><span class="nav-dot" style="background:#6366f1"></span>일자별 매출</div>
        <div class="nav-sub">날짜별 트래킹</div>
      </div>
      <div class="nav-item" data-tab="retailer">
        <div class="nav-label"><span class="nav-dot" style="background:#10b981"></span>유통사별 매출</div>
        <div class="nav-sub">채널별 현황</div>
      </div>
      <div class="nav-item" data-tab="detail">
        <div class="nav-label"><span class="nav-dot" style="background:#f59e0b"></span>상품별 매출</div>
        <div class="nav-sub">SKU 상세 분석</div>
      </div>
    </nav>
    <div class="header-right">
      <div class="status-chip"><div class="live-dot"></div>{now}</div>
    </div>
  </div>
</header>

<main>
  <div class="page-hd">
    <div>
      <div class="page-title">상품DATA 운영 대시보드</div>
      <div class="page-date">자사몰 2026년 판매 데이터 기준 · 상품DATA 표준명으로 통합</div>
    </div>
  </div>

  <section class="kpis">
    <div class="kpi"><div class="kpi-label">자사몰 판매수량</div><div class="kpi-value blue">{total_qty:,}</div><div class="kpi-note">세트 분해 반영</div><div class="kpi-bar"><div style="width:75%;background:var(--blue)"></div></div></div>
    <div class="kpi"><div class="kpi-label">자사몰 실판매금액</div><div class="kpi-value">{total_payment:,}</div><div class="kpi-note">결제금액 배분액</div><div class="kpi-bar"><div style="width:58%;background:#6366f1"></div></div></div>
    <div class="kpi"><div class="kpi-label">평균 판매단가</div><div class="kpi-value">{avg_unit:,}</div><div class="kpi-note">실판매금액 / 수량</div><div class="kpi-bar"><div style="width:48%;background:#8b5cf6"></div></div></div>
    <div class="kpi"><div class="kpi-label">평균 할인율</div><div class="kpi-value amber">{avg_discount}%</div><div class="kpi-note">정상가 기준</div><div class="kpi-bar"><div style="width:{avg_discount}%;background:var(--amber)"></div></div></div>
    <div class="kpi"><div class="kpi-label">매칭률</div><div class="kpi-value green">{match_rate}%</div><div class="kpi-note">{matched}/{len(rows)} 상품옵션</div><div class="kpi-bar"><div style="width:{match_rate}%;background:var(--teal)"></div></div></div>
    <div class="kpi"><div class="kpi-label">미매칭</div><div class="kpi-value">{unmatched}</div><div class="kpi-note">검토 필요 항목</div><div class="kpi-bar"><div style="width:0%;background:var(--red)"></div></div></div>
  </section>

  <div id="overview" class="tab-panel active">
    <div class="grid2">
      <div class="panel">
        <div class="panel-hd"><span class="panel-title">판매 추이</span><span class="panel-meta">월별 집계</span></div>
        <div class="charts-row">
          <div class="chart-box"><canvas id="paymentChart"></canvas></div>
          <div class="chart-box"><canvas id="qtyChart"></canvas></div>
        </div>
      </div>
      <div class="panel">
        <div class="panel-hd"><span class="panel-title">결제금액 상위 상품</span><span class="panel-meta">TOP 12</span></div>
        <div class="rank-list" id="rankList"></div>
      </div>
    </div>
    <div class="grid2-eq">
      <div class="panel">
        <div class="panel-hd"><span class="panel-title">업데이트 상태</span></div>
        <div class="mini-grid">
          <div class="mini"><div class="mini-label">집계 상품수</div><div class="mini-value">{len(rows):,}</div></div>
          <div class="mini"><div class="mini-label">총 주문수</div><div class="mini-value">{total_orders:,}</div></div>
          <div class="mini"><div class="mini-label">총 판매수량</div><div class="mini-value">{total_qty:,}</div></div>
          <div class="mini"><div class="mini-label">총 정상가 금액</div><div class="mini-value">{total_gross:,}</div></div>
          <div class="mini"><div class="mini-label">실판매금액</div><div class="mini-value">{total_payment:,}</div></div>
          <div class="mini"><div class="mini-label">마지막 생성</div><div class="mini-value" style="font-size:11px">{now}</div></div>
        </div>
        <div class="foot">data.json 파일을 업데이트하면 GitHub Actions가 자동으로 대시보드를 재생성합니다.</div>
      </div>
      <div class="panel">
        <div class="panel-hd"><span class="panel-title">미매칭 검토</span></div>
        <div class="table-wrap" style="max-height:240px">
          <table>
            <thead><tr><th>상품번호</th><th>표준상품명</th><th class="num">수량</th><th class="num">금액</th></tr></thead>
            <tbody id="unmatchedRows"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <div id="calendar" class="tab-panel">
    <div class="panel">
      <div class="panel-hd"><span class="panel-title">일자별 매출</span></div>
      <div class="toolbar">
        <input id="dateStart" type="date"/>
        <input id="dateEnd" type="date"/>
        <button class="btn-sm" id="clearDate">전체 초기화</button>
      </div>
      <div class="sum-strip">
        <div class="sum-card"><div class="sum-label">선택 기간 매출</div><div class="sum-value" id="dayPayment">0</div></div>
        <div class="sum-card"><div class="sum-label">선택 기간 주문수</div><div class="sum-value" id="dayOrders">0</div></div>
        <div class="sum-card"><div class="sum-label">선택 기간 판매수량</div><div class="sum-value" id="dayQty">0</div></div>
      </div>
      <div class="chart-box" style="height:180px;margin-bottom:14px"><canvas id="dailyChart"></canvas></div>
      <div class="table-wrap" style="max-height:340px">
        <table>
          <thead><tr><th>일자</th><th class="num">매출</th><th class="num">정상가</th><th class="num">주문수</th><th class="num">판매수량</th><th class="num">객단가</th></tr></thead>
          <tbody id="dailyRows"></tbody>
        </table>
      </div>
    </div>
  </div>

  <div id="retailer" class="tab-panel">
    <div class="panel">
      <div class="panel-hd"><span class="panel-title">유통사별 매출</span></div>
      <div class="toolbar">
        <select id="retailerFilter"><option value="">전체 유통사</option></select>
        <button class="btn-sm" id="clearRetailer">초기화</button>
      </div>
      <div class="chart-box" style="height:180px;margin-bottom:14px"><canvas id="retailerChart"></canvas></div>
      <div class="table-wrap" style="max-height:360px">
        <table>
          <thead><tr><th>유통사</th><th class="num">매출</th><th class="num">정상가</th><th class="num">주문수</th><th class="num">판매수량</th><th class="num">객단가</th></tr></thead>
          <tbody id="retailerRows"></tbody>
        </table>
      </div>
      <div class="foot">새 유통사 데이터가 data.json에 추가되면 자동으로 반영됩니다.</div>
    </div>
  </div>

  <div id="detail" class="tab-panel">
    <div class="panel">
      <div class="panel-hd"><span class="panel-title">상품별 매출</span></div>
      <div class="toolbar">
        <input id="q" placeholder="상품명 · 상품번호 검색"/>
        <select id="productRetailer"><option value="">전체 유통사</option></select>
        <select id="status"><option value="">전체 매칭</option><option>매칭됨</option><option>미매칭</option></select>
        <select id="sourceType"><option value="">전체 구분</option><option>단품</option><option>세트분해</option><option>팩분해</option></select>
        <select id="sortBy"><option value="payment">금액순</option><option value="qty">수량순</option><option value="name">상품명순</option></select>
      </div>
      <div class="table-wrap">
        <table>
          <thead><tr><th>유통사</th><th>상품번호</th><th>표준상품명</th><th>자사몰상품명</th><th>구분</th><th>매칭</th><th class="num">판매수량</th><th class="num">정상가금액</th><th class="num">실판매금액</th><th class="num">평균단가</th><th class="num">할인율</th><th class="num">판매율</th><th class="num">재고소진율</th></tr></thead>
          <tbody id="detailRows"></tbody>
        </table>
      </div>
    </div>
  </div>
</main>

<script>
const rawRows = {rows_json};

const fmt = n => Math.round(n).toLocaleString('ko-KR');
const pct = n => n.toFixed(1) + '%';

// Daily agg
const dMap={{}};
rawRows.forEach(r=>(r.daily||[]).forEach(d=>{{
  if(!dMap[d.date])dMap[d.date]={{date:d.date,payment:0,gross:0,orders:0,qty:0}};
  dMap[d.date].payment+=d.payment;dMap[d.date].gross+=d.gross;
  dMap[d.date].orders+=d.orders;dMap[d.date].qty+=d.qty;
}}));
const dailyAll=Object.values(dMap).sort((a,b)=>a.date.localeCompare(b.date));

// Retailer agg
const rMap={{}};
rawRows.forEach(r=>{{
  if(!rMap[r.retailer])rMap[r.retailer]={{retailer:r.retailer,payment:0,gross:0,orders:0,qty:0}};
  rMap[r.retailer].payment+=r.payment;rMap[r.retailer].gross+=r.gross;
  rMap[r.retailer].orders+=r.orders;rMap[r.retailer].qty+=r.qty;
}});
const retailerAll=Object.values(rMap);

// Populate selects dynamically
const retailers=[...new Set(rawRows.map(r=>r.retailer))];
['retailerFilter','productRetailer'].forEach(id=>{{
  const sel=document.getElementById(id);
  retailers.forEach(r=>{{const o=document.createElement('option');o.value=r;o.textContent=r;sel.appendChild(o);}});
}});

// Unmatched
(function(){{
  const um=rawRows.filter(r=>r.match_status!=='매칭됨');
  document.getElementById('unmatchedRows').innerHTML=um.length
    ? um.map(r=>`<tr><td class="td-mono">${{r.mall_no}}</td><td class="td-main">${{r.standard_name||r.name}}</td><td class="num">${{fmt(r.qty)}}</td><td class="num">${{fmt(r.payment)}}</td></tr>`).join('')
    : '<tr><td colspan="4" style="text-align:center;color:var(--ink3);padding:28px">✓ 미매칭 항목 없음</td></tr>';
}})();

// Chart defaults
Chart.defaults.color='#8a94a6';
Chart.defaults.font.family="'Pretendard',sans-serif";
Chart.defaults.font.size=11;
const ttip={{backgroundColor:'#1e2535',borderColor:'#2b3548',borderWidth:1,titleColor:'#fff',bodyColor:'#c8d0e0',padding:10,callbacks:{{label:ctx=>' '+fmt(ctx.raw)}}}};
const scl={{x:{{grid:{{color:'#e2e6ed'}},ticks:{{maxRotation:45}}}},y:{{grid:{{color:'#e2e6ed'}},ticks:{{callback:v=>fmt(v)}}}}}};

// Overview charts
(function(){{
  const m={{}};
  rawRows.forEach(r=>(r.daily||[]).forEach(d=>{{const k=d.date.slice(0,7);if(!m[k])m[k]={{p:0,q:0}};m[k].p+=d.payment;m[k].q+=d.qty;}}));
  const keys=Object.keys(m).sort(),lbl=keys.map(k=>k.slice(5)+'월');
  new Chart(document.getElementById('paymentChart'),{{type:'bar',data:{{labels:lbl,datasets:[{{data:keys.map(k=>m[k].p),backgroundColor:'rgba(59,130,246,0.18)',borderColor:'#3b82f6',borderWidth:1.5,borderRadius:5}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:ttip,title:{{display:true,text:'월별 매출',color:'#4a5568',padding:{{bottom:8}}}}}},scales:scl}}}});
  new Chart(document.getElementById('qtyChart'),{{type:'bar',data:{{labels:lbl,datasets:[{{data:keys.map(k=>m[k].q),backgroundColor:'rgba(16,185,129,0.15)',borderColor:'#10b981',borderWidth:1.5,borderRadius:5}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:ttip,title:{{display:true,text:'월별 판매수량',color:'#4a5568',padding:{{bottom:8}}}}}},scales:scl}}}});
}})();

// Rank
(function(){{
  const top=[...rawRows].sort((a,b)=>b.payment-a.payment).slice(0,12);
  const mx=top[0]?.payment||1;
  document.getElementById('rankList').innerHTML=top.map((r,i)=>`<div class="rank-row"><div class="rank-n">${{String(i+1).padStart(2,'0')}}</div><div class="rank-name" title="${{r.standard_name}}">${{r.standard_name}}</div><div class="rank-bar"><div style="width:${{(r.payment/mx*100).toFixed(1)}}%"></div></div><div class="rank-val">${{fmt(r.payment)}}</div></div>`).join('');
}})();

// Daily
let dcInst=null;
function renderDaily(){{
  const s=document.getElementById('dateStart').value,e=document.getElementById('dateEnd').value;
  const f=dailyAll.filter(d=>(!s||d.date>=s)&&(!e||d.date<=e));
  document.getElementById('dayPayment').textContent=fmt(f.reduce((a,d)=>a+d.payment,0));
  document.getElementById('dayOrders').textContent=fmt(f.reduce((a,d)=>a+d.orders,0));
  document.getElementById('dayQty').textContent=fmt(f.reduce((a,d)=>a+d.qty,0));
  if(dcInst)dcInst.destroy();
  dcInst=new Chart(document.getElementById('dailyChart'),{{type:'line',data:{{labels:f.map(d=>d.date.slice(5)),datasets:[{{data:f.map(d=>d.payment),borderColor:'#3b82f6',backgroundColor:'rgba(59,130,246,0.07)',borderWidth:1.5,fill:true,tension:0.4,pointRadius:2,pointHoverRadius:5}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:ttip}},scales:scl}}}});
  document.getElementById('dailyRows').innerHTML=!f.length?'<tr><td colspan="6" style="text-align:center;color:var(--ink3);padding:24px">데이터 없음</td></tr>':f.map(d=>`<tr><td class="td-main td-mono">${{d.date}}</td><td class="num">${{fmt(d.payment)}}</td><td class="num">${{fmt(d.gross)}}</td><td class="num">${{fmt(d.orders)}}</td><td class="num">${{fmt(d.qty)}}</td><td class="num">${{d.orders?fmt(d.payment/d.orders):'-'}}</td></tr>`).join('');
}}
document.getElementById('dateStart').addEventListener('change',renderDaily);
document.getElementById('dateEnd').addEventListener('change',renderDaily);
document.getElementById('clearDate').addEventListener('click',()=>{{document.getElementById('dateStart').value='';document.getElementById('dateEnd').value='';renderDaily();}});

// Retailer
let rcInst=null;
function renderRetailer(){{
  const colors=['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4','#ec4899'];
  if(rcInst)rcInst.destroy();
  rcInst=new Chart(document.getElementById('retailerChart'),{{type:'bar',data:{{labels:retailerAll.map(r=>r.retailer),datasets:[{{data:retailerAll.map(r=>r.payment),backgroundColor:retailerAll.map((_,i)=>colors[i%colors.length]+'33'),borderColor:retailerAll.map((_,i)=>colors[i%colors.length]),borderWidth:1.5,borderRadius:5}}]}},options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:ttip}},scales:scl}}}});
  document.getElementById('retailerRows').innerHTML=retailerAll.map(r=>`<tr><td class="td-main"><span class="badge badge-blue">${{r.retailer}}</span></td><td class="num">${{fmt(r.payment)}}</td><td class="num">${{fmt(r.gross)}}</td><td class="num">${{fmt(r.orders)}}</td><td class="num">${{fmt(r.qty)}}</td><td class="num">${{r.orders?fmt(r.payment/r.orders):'-'}}</td></tr>`).join('');
}}
document.getElementById('clearRetailer').addEventListener('click',renderRetailer);

// Detail
function renderDetail(){{
  const q=(document.getElementById('q').value||'').toLowerCase();
  const pR=document.getElementById('productRetailer').value;
  const st=document.getElementById('status').value;
  const sc=document.getElementById('sourceType').value;
  const sb=document.getElementById('sortBy').value;
  let rows=[...rawRows];
  if(q)rows=rows.filter(r=>(r.standard_name||'').toLowerCase().includes(q)||(r.mall_no||'').includes(q));
  if(pR)rows=rows.filter(r=>r.retailer===pR);
  if(st)rows=rows.filter(r=>r.match_status===st);
  if(sc)rows=rows.filter(r=>r.source_type===sc);
  if(sb==='payment')rows.sort((a,b)=>b.payment-a.payment);
  else if(sb==='qty')rows.sort((a,b)=>b.qty-a.qty);
  else rows.sort((a,b)=>(a.standard_name||'').localeCompare(b.standard_name||''));
  const tbody=document.getElementById('detailRows');
  if(!rows.length){{tbody.innerHTML='<tr><td colspan="13" style="text-align:center;color:var(--ink3);padding:28px">검색 결과 없음</td></tr>';return;}}
  tbody.innerHTML=rows.map(r=>{{
    const disc=r.gross>0?(1-r.payment/r.gross)*100:0;
    const sale=r.received_qty>0?r.qty/r.received_qty*100:0;
    const stkOut=r.received_qty>0?(r.received_qty-r.stock_qty)/r.received_qty*100:0;
    const srcB=r.source_type==='단품'?'badge-blue':r.source_type==='세트분해'?'badge-amber':'badge-gray';
    const mB=r.match_status==='매칭됨'?'badge-green':'badge-red';
    return `<tr><td><span class="badge badge-blue">${{r.retailer}}</span></td><td class="td-mono">${{r.mall_no}}</td><td class="td-main" style="max-width:180px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r.standard_name}}">${{r.standard_name||'-'}}</td><td style="font-size:12px;color:var(--ink3);max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${{r.name}} ${{r.size}}</td><td><span class="badge ${{srcB}}">${{r.source_type}}</span></td><td><span class="badge ${{mB}}">${{r.match_status}}</span></td><td class="num">${{fmt(r.qty)}}</td><td class="num">${{fmt(r.gross)}}</td><td class="num" style="color:var(--blue2);font-weight:600">${{fmt(r.payment)}}</td><td class="num">${{fmt(r.avg_unit)}}</td><td class="num" style="color:${{disc>50?'var(--red)':disc>30?'var(--amber)':'var(--ink2)'}}">${{pct(disc)}}</td><td class="num">${{sale>0?pct(sale):'-'}}</td><td class="num">${{stkOut>0?pct(stkOut):'-'}}</td></tr>`;
  }}).join('');
}}
['q','productRetailer','status','sourceType','sortBy'].forEach(id=>{{
  document.getElementById(id).addEventListener('input',renderDetail);
  document.getElementById(id).addEventListener('change',renderDetail);
}});

// Nav
document.querySelectorAll('.nav-item[data-tab]').forEach(btn=>{{
  btn.addEventListener('click',()=>{{
    document.querySelectorAll('.nav-item').forEach(b=>b.classList.remove('active'));
    document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
    btn.classList.add('active');
    const tab=btn.dataset.tab;
    document.getElementById(tab).classList.add('active');
    if(tab==='calendar')renderDaily();
    if(tab==='retailer')renderRetailer();
    if(tab==='detail')renderDetail();
  }});
}});
renderDaily();
</script>
</body>
</html>"""

OUT_FILE.write_text(html, encoding="utf-8")
print(f"✅ {OUT_FILE} 생성 완료 ({OUT_FILE.stat().st_size:,} bytes)")
