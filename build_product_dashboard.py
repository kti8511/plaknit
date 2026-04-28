"""
build_product_dashboard.py
data.json → index.html 생성

data.json 구조:
[
  {
    "retailer": "자사몰",          # 유통사명
    "mall_no": "411",
    "name": "상품명",
    "color": "색상",
    "size": "사이즈",
    "qty": 26,
    "gross": 2054000,
    "payment": 1300113,
    "avg_unit": 50004,
    "orders": 25,
    "source_type": "단품",         # 단품 / 세트분해 / 팩분해
    "daily": [
      {"date": "2026-01-03", "qty": 1, "gross": 79000, "payment": 0, "orders": 1}
    ],
    "match_status": "매칭됨",
    "match_sku": "PD4PLM01BLK000L",
    "standard_name": "ESSENTIAL HEAT 기모 러닝팬츠_L",
    "received_qty": 100,           # 누계입고수량
    "stock_qty": 27                # 현재고 (총재고수량)
  }
]
"""
import json, datetime, pathlib

DATA_FILE = pathlib.Path("data.json")
HISTORICAL_DAILY_FILE = pathlib.Path("historical_daily.json")
MANUAL_SALES_FILE = pathlib.Path("manual_sales_updates.json")
OUT_FILE  = pathlib.Path("index.html")

with DATA_FILE.open(encoding="utf-8") as f:
    rows = json.load(f)

historical_daily = {}
if HISTORICAL_DAILY_FILE.exists():
    with HISTORICAL_DAILY_FILE.open(encoding="utf-8") as f:
        historical_daily = json.load(f)

manual_sales_updates = []
if MANUAL_SALES_FILE.exists():
    with MANUAL_SALES_FILE.open(encoding="utf-8") as f:
        manual_sales_updates = json.load(f)

now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
rows_json = json.dumps(rows, ensure_ascii=False)
historical_daily_json = json.dumps(historical_daily, ensure_ascii=False)
manual_sales_json = json.dumps(manual_sales_updates, ensure_ascii=False)

# 2026 플래니트 예산.xlsx > 2025 실매출 > 월별 합계(row 32) 기준.
# GitHub Actions 환경에서도 재생성 가능하도록 기준값은 스크립트에 고정한다.
PREV_YEAR_MONTHLY = [
    {"month": "01", "sales2025": 16573021},
    {"month": "02", "sales2025": 15125056},
    {"month": "03", "sales2025": 16677513},
    {"month": "04", "sales2025": 36350375},
    {"month": "05", "sales2025": 43723980},
    {"month": "06", "sales2025": 48364959},
    {"month": "07", "sales2025": 45331462},
    {"month": "08", "sales2025": 53199752},
    {"month": "09", "sales2025": 45022043},
    {"month": "10", "sales2025": 19962162},
    {"month": "11", "sales2025": 23431013},
    {"month": "12", "sales2025": 20800000},
]
prev_year_monthly_json = json.dumps(PREV_YEAR_MONTHLY, ensure_ascii=False)

HISTORICAL_MONTHLY = {
    "2024": [
        {"month": "01", "sales": 28809285},
        {"month": "02", "sales": 36138277},
        {"month": "03", "sales": 40961929},
        {"month": "04", "sales": 58001498},
        {"month": "05", "sales": 73659807},
        {"month": "06", "sales": 92094266},
        {"month": "07", "sales": 91388596},
        {"month": "08", "sales": 57406405},
        {"month": "09", "sales": 53437759},
        {"month": "10", "sales": 29647346},
        {"month": "11", "sales": 24042161},
        {"month": "12", "sales": 30710000},
    ],
    "2025": [{"month": r["month"], "sales": r["sales2025"]} for r in PREV_YEAR_MONTHLY],
}
historical_monthly_json = json.dumps(HISTORICAL_MONTHLY, ensure_ascii=False)

historical_daily_monthly = {}
for year, items in historical_daily.items():
    month_map = {}
    for item in items:
        month = item["date"][5:7]
        month_map[month] = month_map.get(month, 0) + int(item.get("payment", 0))
    historical_daily_monthly[year] = [
        {"month": month, "sales": month_map[month]}
        for month in sorted(month_map)
    ]
historical_daily_monthly_json = json.dumps(historical_daily_monthly, ensure_ascii=False)

# 유통사 목록 (고정)
RETAILERS = ["자사몰", "무신사", "29cm", "글로리어스워커", "4XR", "애슬러", "롯데온"]

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
  --navy:#1e2535;--bg:#f0f2f6;--panel:#fff;
  --border:#e2e6ed;--border2:#d0d5df;--ink:#1e2535;--ink2:#4a5568;--ink3:#8a94a6;
  --blue:#3b82f6;--blue2:#2563eb;--blue-soft:#eff6ff;
  --teal:#10b981;--teal-soft:#ecfdf5;--amber:#f59e0b;--amber-soft:#fffbeb;
  --red:#ef4444;--red-soft:#fef2f2;--indigo:#6366f1;
  --shadow:0 1px 3px rgba(0,0,0,0.08),0 1px 2px rgba(0,0,0,0.05);
  --shadow-md:0 4px 12px rgba(0,0,0,0.08),0 2px 4px rgba(0,0,0,0.04);
}}
*{{box-sizing:border-box;margin:0;padding:0;}}
body{{font-family:'Pretendard',-apple-system,sans-serif;background:var(--bg);color:var(--ink);min-height:100vh;}}

/* HEADER */
header{{background:var(--navy);position:sticky;top:0;z-index:100;box-shadow:0 2px 8px rgba(0,0,0,0.2);}}
.header-top{{display:flex;align-items:center;height:52px;padding:0 20px;}}
.logo-area{{display:flex;align-items:center;gap:10px;min-width:156px;border-right:1px solid rgba(255,255,255,0.1);padding-right:20px;margin-right:4px;height:100%;}}
.logo-icon{{width:32px;height:32px;border-radius:8px;background:linear-gradient(135deg,#3b82f6,#6366f1);display:grid;place-items:center;font-size:13px;font-weight:800;color:#fff;flex-shrink:0;}}
.logo-text{{font-size:13px;font-weight:700;color:#fff;letter-spacing:-0.2px;line-height:1.3;}}
.logo-sub{{font-size:10px;color:rgba(255,255,255,0.4);font-weight:400;}}
.header-nav{{display:flex;align-items:center;height:100%;flex:1;overflow-x:auto;}}
.nav-item{{display:flex;flex-direction:column;justify-content:center;height:100%;padding:0 18px;cursor:pointer;border-bottom:3px solid transparent;transition:all 0.15s;white-space:nowrap;}}
.nav-item:hover{{background:rgba(255,255,255,0.05);}}
.nav-item.active{{border-bottom-color:var(--blue);}}
.nav-label{{font-size:13px;font-weight:600;color:rgba(255,255,255,0.85);display:flex;align-items:center;gap:6px;}}
.nav-dot{{width:10px;height:10px;border-radius:2px;flex-shrink:0;}}
.nav-sub{{font-size:10px;color:rgba(255,255,255,0.38);margin-top:1px;}}
.header-right{{margin-left:auto;flex-shrink:0;}}
.status-chip{{display:flex;align-items:center;gap:5px;background:rgba(16,185,129,0.15);border:1px solid rgba(16,185,129,0.3);border-radius:20px;padding:4px 10px;font-size:11px;color:#6ee7b7;font-family:'DM Mono',monospace;}}
.live-dot{{width:5px;height:5px;border-radius:50%;background:#10b981;box-shadow:0 0 6px #10b981;animation:blink 2s infinite;}}
@keyframes blink{{0%,100%{{opacity:1}}50%{{opacity:.3}}}}

/* MAIN */
main{{padding:20px 20px 48px;max-width:1600px;margin:0 auto;}}
.page-hd{{margin-bottom:16px;}}
.page-title{{font-size:20px;font-weight:800;color:var(--ink);letter-spacing:-0.4px;}}
.page-date{{font-size:12px;color:var(--ink3);margin-top:3px;}}

/* KPI */
.kpis{{display:grid;grid-template-columns:repeat(6,1fr);gap:10px;margin-bottom:16px;}}
.kpi{{background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:16px;box-shadow:var(--shadow);transition:box-shadow 0.15s,transform 0.15s;}}
.kpi:hover{{box-shadow:var(--shadow-md);transform:translateY(-1px);}}
.kpi-label{{font-size:11px;font-weight:600;color:var(--ink3);text-transform:uppercase;letter-spacing:0.5px;margin-bottom:8px;}}
.kpi-value{{font-size:22px;font-weight:800;color:var(--ink);letter-spacing:-0.5px;font-family:'DM Mono',monospace;}}
.kpi-value.blue{{color:var(--blue2);}} .kpi-value.green{{color:var(--teal);}} .kpi-value.amber{{color:var(--amber);}}
.kpi-note{{font-size:11px;color:var(--ink3);margin-top:5px;}}
.kpi-bar{{height:3px;border-radius:2px;margin-top:10px;background:var(--border);overflow:hidden;}}
.kpi-bar div{{height:100%;border-radius:2px;}}
.kpi-link{{cursor:pointer;}} .kpi-link:hover{{box-shadow:var(--shadow-md);transform:translateY(-1px);}}

/* PANEL */
.panel{{background:var(--panel);border:1px solid var(--border);border-radius:10px;padding:18px;box-shadow:var(--shadow);}}
.panel-hd{{display:flex;align-items:center;justify-content:space-between;margin-bottom:14px;}}
.panel-title{{font-size:14px;font-weight:700;color:var(--ink);letter-spacing:-0.2px;}}
.panel-link{{cursor:pointer;}} .panel-link:hover{{box-shadow:var(--shadow-md);border-color:var(--border2);}}
.panel-meta{{font-size:11px;color:var(--ink3);font-family:'DM Mono',monospace;}}

/* LAYOUT */
.grid2{{display:grid;grid-template-columns:1.4fr 0.6fr;gap:12px;margin-bottom:12px;}}
.grid2-eq{{display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:12px;}}
.charts-row{{display:grid;grid-template-columns:1fr 1fr;gap:12px;}}
.chart-box{{position:relative;height:200px;}}
.chart-box canvas{{width:100%!important;height:100%!important;}}

/* RANK */
.rank-list{{display:flex;flex-direction:column;gap:1px;}}
.rank-row{{display:grid;grid-template-columns:22px minmax(0,1fr) 80px 80px;align-items:center;gap:8px;padding:5px 6px;border-radius:6px;transition:background 0.12s;}}
.rank-row:hover{{background:var(--bg);}}
.rank-n{{font-size:11px;color:var(--ink3);font-family:'DM Mono',monospace;}}
.rank-name{{font-size:12px;color:var(--ink2);overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}}
.rank-bar{{height:5px;background:var(--border);border-radius:3px;overflow:hidden;}}
.rank-bar div{{height:100%;border-radius:3px;background:linear-gradient(90deg,var(--blue),#818cf8);}}
.rank-val{{font-size:12px;color:var(--ink);text-align:right;font-family:'DM Mono',monospace;font-weight:600;}}

/* MINI CARDS */
.mini-grid{{display:grid;grid-template-columns:repeat(3,1fr);gap:8px;}}
.mini{{background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:11px 12px;}}
.mini-label{{font-size:10px;color:var(--ink3);text-transform:uppercase;letter-spacing:0.5px;font-weight:600;}}
.mini-value{{font-size:15px;font-weight:700;color:var(--ink);margin-top:4px;font-family:'DM Mono',monospace;}}

/* SUM STRIP */
.sum-strip{{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin-bottom:14px;}}
.sum-card{{background:var(--bg);border:1px solid var(--border);border-radius:8px;padding:12px 14px;}}
.sum-label{{font-size:11px;color:var(--ink3);font-weight:500;}}
.sum-value{{font-size:19px;font-weight:800;color:var(--ink);margin-top:3px;font-family:'DM Mono',monospace;}}

/* SUB TABS (일자별 내 유통사 탭) */
.sub-tabs{{display:flex;gap:4px;margin-bottom:14px;flex-wrap:wrap;}}
.sub-tab{{height:30px;padding:0 14px;border:1px solid var(--border2);border-radius:20px;font-size:12px;font-weight:600;color:var(--ink3);background:var(--panel);cursor:pointer;font-family:'Pretendard',sans-serif;transition:all 0.15s;white-space:nowrap;}}
.sub-tab:hover{{color:var(--ink2);border-color:var(--blue);}}
.sub-tab.active{{background:var(--blue);color:#fff;border-color:var(--blue);}}

/* TOOLBAR */
.toolbar{{display:flex;gap:8px;align-items:center;margin-bottom:12px;flex-wrap:wrap;}}
.toolbar input,.toolbar select{{height:34px;border:1px solid var(--border2);border-radius:7px;background:var(--panel);color:var(--ink);font-size:13px;padding:0 10px;font-family:'Pretendard',sans-serif;outline:none;min-width:140px;transition:border-color 0.15s,box-shadow 0.15s;}}
.toolbar input:focus,.toolbar select:focus{{border-color:var(--blue);box-shadow:0 0 0 3px rgba(59,130,246,0.1);}}
.toolbar input::placeholder{{color:var(--ink3);}}
.btn-sm{{height:34px;padding:0 14px;background:var(--panel);border:1px solid var(--border2);border-radius:7px;font-size:12px;font-weight:600;color:var(--ink2);cursor:pointer;font-family:'Pretendard',sans-serif;transition:all 0.15s;white-space:nowrap;}}
.btn-sm:hover{{background:var(--bg);color:var(--ink);}}

/* TABLE */
.table-wrap{{border:1px solid var(--border);border-radius:8px;overflow:auto;max-height:500px;}}
table{{border-collapse:collapse;width:100%;min-width:900px;font-size:12.5px;}}
thead{{position:sticky;top:0;z-index:2;}}
th{{background:#f8fafc;color:var(--ink3);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;padding:9px 10px;border-bottom:1px solid var(--border);text-align:left;white-space:nowrap;}}
td{{padding:9px 10px;border-bottom:1px solid var(--border);color:var(--ink2);vertical-align:middle;}}
.table-filter-row td{{position:sticky;top:37px;background:#fff;z-index:1;padding:6px 8px;}}
.table-filter-row input,.table-filter-row select{{width:100%;height:28px;border:1px solid var(--border);border-radius:6px;padding:0 7px;font-size:11px;color:var(--ink2);background:#fff;}}
tr:last-child td{{border-bottom:none;}} tr:hover td{{background:#fafbfd;}}
.num{{text-align:right;font-family:'DM Mono',monospace;font-size:12px;}}
.td-main{{color:var(--ink);font-weight:600;}} .td-mono{{font-family:'DM Mono',monospace;font-size:11.5px;color:var(--ink3);}}

/* BADGE */
.badge{{display:inline-flex;align-items:center;height:22px;padding:0 8px;border-radius:5px;font-size:11px;font-weight:600;white-space:nowrap;}}
.badge-blue{{background:var(--blue-soft);color:var(--blue2);}} .badge-green{{background:var(--teal-soft);color:#059669;}}
.badge-amber{{background:var(--amber-soft);color:#b45309;}} .badge-red{{background:var(--red-soft);color:#dc2626;}}
.badge-gray{{background:#f1f5f9;color:var(--ink3);}} .badge-indigo{{background:#eef2ff;color:#4338ca;}}

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
        <div class="nav-sub">KPI · 최근 트렌드</div>
      </div>
      <div class="nav-item" data-tab="compare">
        <div class="nav-label"><span class="nav-dot" style="background:#06b6d4"></span>매출 비교</div>
        <div class="nav-sub">연도 · 월별 대비</div>
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
    <div class="page-title">상품DATA 운영 대시보드</div>
    <div class="page-date">전체 유통사 통합 판매 데이터 기준 · 상품DATA 표준명으로 통합</div>
  </div>

  <!-- KPI -->
  <section class="kpis" id="kpiSection">
    <div class="kpi kpi-link" data-tab-jump="detail"><div class="kpi-label">주간 판매수량</div><div class="kpi-value blue" id="kpiQty">-</div><div class="kpi-note">최근 7일 총 판매 수량</div><div class="kpi-bar"><div id="kpiQtyBar" style="width:75%;background:var(--blue)"></div></div></div>
    <div class="kpi kpi-link" data-tab-jump="calendar"><div class="kpi-label">주간 실판매 금액</div><div class="kpi-value" id="kpiPayment">-</div><div class="kpi-note">최근 7일 실판매 합계</div><div class="kpi-bar"><div id="kpiPaymentBar" style="width:58%;background:#6366f1"></div></div></div>
    <div class="kpi"><div class="kpi-label">평균 판매단가</div><div class="kpi-value" id="kpiAvgUnit">-</div><div class="kpi-note">실판매금액 / 수량</div><div class="kpi-bar"><div id="kpiAvgBar" style="width:48%;background:#8b5cf6"></div></div></div>
    <div class="kpi"><div class="kpi-label">평균 할인율</div><div class="kpi-value amber" id="kpiDisc">-</div><div class="kpi-note">정상가 기준</div><div class="kpi-bar"><div id="kpiDiscBar" style="width:0%;background:var(--amber)"></div></div></div>
    <div class="kpi"><div class="kpi-label">매칭률</div><div class="kpi-value green" id="kpiMatch">-</div><div class="kpi-note" id="kpiMatchNote">-</div><div class="kpi-bar"><div id="kpiMatchBar" style="width:0%;background:var(--teal)"></div></div></div>
    <div class="kpi kpi-link" data-tab-jump="overview" data-focus="unmatched"><div class="kpi-label">미매칭</div><div class="kpi-value" id="kpiUnmatch">-</div><div class="kpi-note">검토 필요 항목</div><div class="kpi-bar"><div style="width:0%;background:var(--red)"></div></div></div>
  </section>

  <!-- OVERVIEW TAB -->
  <div id="overview" class="tab-panel active">
    <div class="grid2">
      <div class="panel panel-link" data-tab-jump="calendar">
        <div class="panel-hd"><span class="panel-title">최근 3일 일별 매출</span><span class="panel-meta" id="recentDateRange">-</span></div>
        <div class="charts-row">
          <div class="chart-box"><canvas id="paymentChart"></canvas></div>
          <div class="chart-box"><canvas id="qtyChart"></canvas></div>
        </div>
      </div>
      <div class="panel panel-link" data-tab-jump="detail">
        <div class="panel-hd"><span class="panel-title">최근 7일 결제금액 상위</span><span class="panel-meta">TOP 10</span></div>
        <div class="rank-list" id="rankList"></div>
      </div>
    </div>
    <div class="panel panel-link" data-tab-jump="retailer" style="margin-bottom:12px">
      <div class="panel-hd"><span class="panel-title">최근 3일 유통사별 매출</span><span class="panel-meta">최근 3일</span></div>
      <div class="chart-box" style="height:220px"><canvas id="recentRetailerChart"></canvas></div>
    </div>
    <div class="panel panel-link" data-tab-jump="compare" style="margin-bottom:12px">
      <div class="panel-hd"><span class="panel-title">월별 전년 대비 매출 신장률</span><span class="panel-meta">2026 vs 2025 실매출</span></div>
      <div class="chart-box" style="height:190px;margin-bottom:12px"><canvas id="monthlyGrowthChart"></canvas></div>
      <div class="table-wrap" style="max-height:240px">
        <table>
          <thead><tr><th>월</th><th class="num">2026 매출</th><th class="num">2025 매출</th><th class="num">신장률</th></tr></thead>
          <tbody id="monthlyGrowthRows"></tbody>
        </table>
      </div>
    </div>
    <div class="grid2-eq">
      <div class="panel">
        <div class="panel-hd"><span class="panel-title">업데이트 상태</span></div>
        <div class="mini-grid" id="statusMini"></div>
        <div class="foot">data.json 업데이트 시 GitHub Actions가 자동으로 대시보드를 재생성합니다.</div>
      </div>
      <div class="panel">
        <div class="panel-hd"><span class="panel-title" id="unmatchedSection">미매칭 검토</span></div>
        <div class="table-wrap" style="max-height:240px">
          <table>
            <thead><tr><th>유통사</th><th>상품번호</th><th>상품명</th><th class="num">수량</th><th class="num">금액</th></tr></thead>
            <tbody id="unmatchedRows"></tbody>
          </table>
        </div>
      </div>
    </div>
  </div>

  <!-- COMPARE TAB -->
  <div id="compare" class="tab-panel">
    <div class="panel">
      <div class="panel-hd"><span class="panel-title">매출 비교</span><span class="panel-meta">2024 · 2025 · 2026 월별 대비</span></div>
      <div class="toolbar">
        <select id="compareBaseYear"></select>
        <select id="compareTargetYear"></select>
        <select id="compareMonth">
          <option value="">전체 월</option>
          <option value="01">1월</option><option value="02">2월</option><option value="03">3월</option>
          <option value="04">4월</option><option value="05">5월</option><option value="06">6월</option>
          <option value="07">7월</option><option value="08">8월</option><option value="09">9월</option>
          <option value="10">10월</option><option value="11">11월</option><option value="12">12월</option>
        </select>
      </div>
      <div class="sum-strip">
        <div class="sum-card"><div class="sum-label" id="compareBaseLabel">기준연도 매출</div><div class="sum-value" id="compareBaseSales">0</div></div>
        <div class="sum-card"><div class="sum-label" id="compareTargetLabel">비교연도 매출</div><div class="sum-value" id="compareTargetSales">0</div></div>
        <div class="sum-card"><div class="sum-label">대비 신장률</div><div class="sum-value" id="compareGrowth">-</div></div>
      </div>
      <div class="chart-box" style="height:260px;margin-bottom:14px"><canvas id="compareChart"></canvas></div>
      <div class="table-wrap" style="max-height:360px">
        <table>
          <thead><tr><th>월</th><th class="num" id="compareBaseHead">기준연도</th><th class="num" id="compareTargetHead">비교연도</th><th class="num">차이</th><th class="num">신장률</th></tr></thead>
          <tbody id="compareRows"></tbody>
        </table>
      </div>
      <div class="foot">2024/2025는 예산 파일의 실매출 월합계 기준이며, 2026은 현재 업로드된 일별 판매 데이터 기준입니다.</div>
    </div>
  </div>

  <!-- CALENDAR TAB -->
  <div id="calendar" class="tab-panel">
    <div class="panel">
      <div class="panel-hd"><span class="panel-title">일자별 매출</span></div>

      <!-- 유통사 서브 탭 -->
      <div class="sub-tabs" id="calRetailerTabs">
        <button class="sub-tab active" data-retailer="">전체</button>
        <button class="sub-tab" data-retailer="자사몰">자사몰</button>
        <button class="sub-tab" data-retailer="무신사">무신사</button>
        <button class="sub-tab" data-retailer="29cm">29cm</button>
        <button class="sub-tab" data-retailer="글로리어스워커">글로리어스워커</button>
        <button class="sub-tab" data-retailer="4XR">4XR</button>
        <button class="sub-tab" data-retailer="애슬러">애슬러</button>
        <button class="sub-tab" data-retailer="롯데온">롯데온</button>
      </div>

      <div class="toolbar">
        <select id="calYear">
          <option value="2026">2026년</option>
          <option value="2025">2025년</option>
          <option value="2024">2024년</option>
          <option value="2023">2023년</option>
        </select>
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

  <!-- RETAILER TAB -->
  <div id="retailer" class="tab-panel">
    <div class="panel">
      <div class="panel-hd"><span class="panel-title">유통사별 매출</span></div>
      <div class="toolbar">
        <select id="retailerYear">
          <option value="2026">2026년</option>
          <option value="2025">2025년</option>
          <option value="2024">2024년</option>
          <option value="2023">2023년</option>
        </select>
        <select id="retailerFilter">
          <option value="">전체 유통사</option>
          <option>자사몰</option><option>무신사</option><option>29cm</option>
          <option>글로리어스워커</option><option>4XR</option><option>애슬러</option><option>롯데온</option>
        </select>
        <button class="btn-sm" id="clearRetailer">초기화</button>
      </div>
      <div class="chart-box" style="height:200px;margin-bottom:14px"><canvas id="retailerChart"></canvas></div>
      <div class="table-wrap" style="max-height:360px">
        <table>
          <thead><tr><th>유통사</th><th class="num">매출</th><th class="num">정상가</th><th class="num">주문수</th><th class="num">판매수량</th><th class="num">객단가</th><th class="num">할인율</th></tr></thead>
          <tbody id="retailerRows"></tbody>
        </table>
      </div>
      <div class="foot">새 유통사 데이터가 data.json에 추가되면 자동으로 반영됩니다.</div>
    </div>
  </div>

  <!-- DETAIL TAB -->
  <div id="detail" class="tab-panel">
    <div class="panel">
      <div class="panel-hd"><span class="panel-title">상품별 매출</span></div>
      <div class="toolbar">
        <input id="q" placeholder="상품명 · 상품번호 검색"/>
        <select id="seasonFilter"><option value="">전체 시즌</option></select>
        <select id="categoryLargeFilter"><option value="">전체 복종(대)</option></select>
        <select id="categorySmallFilter"><option value="">전체 복종(소)</option></select>
        <select id="sortBy"><option value="payment">금액순</option><option value="qty">수량순</option><option value="weeklyRate">판매율순</option><option value="name">상품명순</option></select>
      </div>
      <div class="sum-strip">
        <div class="sum-card"><div class="sum-label">전체 출고가능 재고</div><div class="sum-value" id="detailStockQty">0</div></div>
        <div class="sum-card"><div class="sum-label">전체 재고 소진율</div><div class="sum-value" id="detailStockRate">-</div></div>
        <div class="sum-card"><div class="sum-label">주간 상품 판매율</div><div class="sum-value" id="detailWeeklyRate">-</div></div>
      </div>
      <div class="panel" style="margin-bottom:12px;background:#fff7ed;border-color:#fed7aa">
        <div class="panel-hd"><span class="panel-title">리오더 알림</span><span class="panel-meta" id="reorderSummary">-</span></div>
        <div class="rank-list" id="reorderAlerts"></div>
      </div>
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th>알림</th><th>바코드</th><th>표준상품명</th><th>유통사</th>
              <th>시즌</th><th>복종</th>
              <th class="num">판매수량</th>
              <th class="num">주간 판매율</th>
              <th class="num">현재고</th>
              <th class="num">전체 재고 소진율</th>
              <th class="num">정상가금액</th>
              <th class="num">실판매금액</th>
              <th class="num">평균단가</th>
              <th class="num">할인율</th>
            </tr>
          </thead>
          <tbody id="detailFilterRows"></tbody>
          <tbody id="detailRows"></tbody>
        </table>
      </div>
    </div>
  </div>
</main>

<script>
const rawRows = {rows_json};
const RETAILERS = {json.dumps(RETAILERS, ensure_ascii=False)};
const PREV_YEAR_MONTHLY = {prev_year_monthly_json};
const HISTORICAL_MONTHLY = {historical_monthly_json};
const HISTORICAL_DAILY = {historical_daily_json};
const HISTORICAL_DAILY_MONTHLY = {historical_daily_monthly_json};
const MANUAL_SALES = {manual_sales_json};

const fmt  = n => Math.round(n).toLocaleString('ko-KR');
const pct  = n => n.toFixed(1) + '%';
const fmtD = s => s.slice(5); // "2026-01-03" → "01-03"
function uniqSorted(values) {{
  return Array.from(new Set(values.filter(v => v !== undefined && v !== null && String(v).trim() !== '').map(v => String(v).trim()))).sort((a,b)=>a.localeCompare(b));
}}
function validDiscount(gross, payment) {{
  gross = Number(gross || 0);
  payment = Number(payment || 0);
  if (gross <= 0 || payment < 0 || payment > gross) return null;
  return (1 - payment / gross) * 100;
}}
function discountText(gross, payment) {{
  const d = validDiscount(gross, payment);
  return d === null ? '-' : pct(d);
}}

// ── 전체 집계 ─────────────────────────────────────────────
function shiftDate(dateStr, days) {{
  const dt = new Date(dateStr + 'T00:00:00');
  dt.setDate(dt.getDate() + days);
  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, '0');
  const d = String(dt.getDate()).padStart(2, '0');
  return `${{y}}-${{m}}-${{d}}`;
}}

const totalQty     = rawRows.reduce((a,r)=>a+r.qty, 0);
const totalPayment = rawRows.reduce((a,r)=>a+r.payment, 0);
const totalGross   = rawRows.reduce((a,r)=>a+r.gross, 0);
const validDiscRows= rawRows.filter(r => validDiscount(r.gross, r.payment) !== null);
const validGross   = validDiscRows.reduce((a,r)=>a+r.gross, 0);
const validPayment = validDiscRows.reduce((a,r)=>a+r.payment, 0);
const totalOrders  = rawRows.reduce((a,r)=>a+r.orders, 0);
const avgUnit      = totalQty ? Math.round(totalPayment/totalQty) : 0;
const avgDisc      = validGross ? (1-validPayment/validGross)*100 : null;
const isMatchedRow = r => !String(r.match_status || '').includes('?');
const matched      = rawRows.filter(isMatchedRow).length;
const unmatched    = rawRows.length - matched;
const matchRate    = rawRows.length ? matched/rawRows.length*100 : 0;
const allDailyForKpi = buildDailyMap(null);
const latestKpiDate = allDailyForKpi.length ? allDailyForKpi[allDailyForKpi.length - 1].date : null;
const weeklyCutoff = latestKpiDate ? shiftDate(latestKpiDate, -6) : null;
const weeklyDaily = latestKpiDate ? allDailyForKpi.filter(d => d.date >= weeklyCutoff && d.date <= latestKpiDate) : [];
const weeklyQty = weeklyDaily.reduce((a,d)=>a + Number(d.qty || 0), 0);
const weeklyPayment = weeklyDaily.reduce((a,d)=>a + Number(d.payment || 0), 0);

// KPI ???
document.getElementById('kpiQty').textContent      = fmt(weeklyQty);
document.getElementById('kpiPayment').textContent  = fmt(weeklyPayment);
document.getElementById('kpiAvgUnit').textContent  = fmt(avgUnit);
document.getElementById('kpiDisc').textContent     = avgDisc === null ? '-' : pct(avgDisc);
document.getElementById('kpiMatch').textContent    = pct(matchRate);
document.getElementById('kpiMatchNote').textContent= matched+'/'+rawRows.length+' ???';
document.getElementById('kpiUnmatch').textContent  = unmatched;
document.getElementById('kpiDiscBar').style.width  = avgDisc === null ? '0%' : Math.max(0, Math.min(100, avgDisc)).toFixed(1)+'%';
document.getElementById('kpiMatchBar').style.width = matchRate.toFixed(1)+'%';

function buildDailyMap(retailerFilter) {{
  const m = {{}};
  const existingByDateRetailer = {{}};
  rawRows.forEach(r => {{
    if (retailerFilter && r.retailer !== retailerFilter) return;
    (r.daily||[]).forEach(d => {{
      if (!m[d.date]) m[d.date] = {{date:d.date,payment:0,gross:0,orders:0,qty:0}};
      m[d.date].payment += d.payment;
      m[d.date].gross   += d.gross;
      m[d.date].orders  += d.orders;
      m[d.date].qty     += d.qty;
      const key = d.date + '|' + (r.retailer || '');
      if (!existingByDateRetailer[key]) existingByDateRetailer[key] = {{payment:0,gross:0,orders:0,qty:0}};
      existingByDateRetailer[key].payment += Number(d.payment || 0);
      existingByDateRetailer[key].gross += Number(d.gross || 0);
      existingByDateRetailer[key].orders += Number(d.orders || 0);
      existingByDateRetailer[key].qty += Number(d.qty || 0);
    }});
  }});
  MANUAL_SALES.forEach(day => {{
    (day.retailers || []).forEach(r => {{
      if (retailerFilter && r.retailer !== retailerFilter) return;
      if (!m[day.date]) m[day.date] = {{date:day.date,payment:0,gross:0,orders:0,qty:0}};
      const existing = existingByDateRetailer[day.date + '|' + (r.retailer || '')] || {{payment:0,gross:0,orders:0,qty:0}};
      m[day.date].payment += Math.max(0, Number(r.payment || 0) - existing.payment);
      if (r.gross !== undefined) m[day.date].gross += Math.max(0, Number(r.gross || 0) - existing.gross);
      if (r.orders !== undefined) m[day.date].orders += Number(r.orders || 0) - existing.orders;
      if (r.qty !== undefined) m[day.date].qty += Number(r.qty || 0) - existing.qty;
    }});
  }});
  return Object.values(m).sort((a,b)=>a.date.localeCompare(b.date));
}}

function buildHistoricalDailyMap(yearFilter, retailerFilter) {{
  const rows = HISTORICAL_DAILY[yearFilter] || [];
  return rows.map(r => {{
    const payment = retailerFilter ? Number((r.retailers || {{}})[retailerFilter] || 0) : Number(r.payment || 0);
    return {{date:r.date,payment,gross:0,orders:0,qty:0}};
  }}).filter(r => !retailerFilter || r.payment !== 0);
}}

// ── 유통사 집계 ────────────────────────────────────────────
const retailerMap = {{}};
const retailerExisting = {{}};
rawRows.forEach(r => {{
  if (!retailerMap[r.retailer]) retailerMap[r.retailer] = {{retailer:r.retailer,payment:0,gross:0,orders:0,qty:0,validPayment:0,validGross:0}};
  retailerMap[r.retailer].payment += r.payment;
  retailerMap[r.retailer].gross   += r.gross;
  retailerMap[r.retailer].orders  += r.orders;
  retailerMap[r.retailer].qty     += r.qty;
  (r.daily || []).forEach(d => {{
    const key = d.date + '|' + (r.retailer || '');
    if (!retailerExisting[key]) retailerExisting[key] = {{payment:0,gross:0,orders:0,qty:0}};
    retailerExisting[key].payment += Number(d.payment || 0);
    retailerExisting[key].gross += Number(d.gross || 0);
    retailerExisting[key].orders += Number(d.orders || 0);
    retailerExisting[key].qty += Number(d.qty || 0);
  }});
  if (validDiscount(r.gross, r.payment) !== null) {{
    retailerMap[r.retailer].validPayment += r.payment;
    retailerMap[r.retailer].validGross += r.gross;
  }}
}});
MANUAL_SALES.forEach(day => {{
  (day.retailers || []).forEach(r => {{
    if (!retailerMap[r.retailer]) retailerMap[r.retailer] = {{retailer:r.retailer,payment:0,gross:0,orders:0,qty:0,validPayment:0,validGross:0}};
    const existing = retailerExisting[day.date + '|' + (r.retailer || '')] || {{payment:0,gross:0,orders:0,qty:0}};
    retailerMap[r.retailer].payment += Math.max(0, Number(r.payment || 0) - existing.payment);
    if (r.gross !== undefined) retailerMap[r.retailer].gross += Math.max(0, Number(r.gross || 0) - existing.gross);
    if (r.orders !== undefined) retailerMap[r.retailer].orders += Number(r.orders || 0) - existing.orders;
    if (r.qty !== undefined) retailerMap[r.retailer].qty += Number(r.qty || 0) - existing.qty;
  }});
}});
const retailerAll = Object.values(retailerMap).sort((a,b)=>b.payment-a.payment);

function buildHistoricalRetailerSummary(yearFilter) {{
  const totals = {{}};
  (HISTORICAL_DAILY[yearFilter] || []).forEach(day => {{
    Object.entries(day.retailers || {{}}).forEach(([retailer, payment]) => {{
      if (!totals[retailer]) totals[retailer] = {{retailer, payment:0, gross:0, orders:0, qty:0, validPayment:0, validGross:0}};
      totals[retailer].payment += Number(payment || 0);
    }});
  }});
  return Object.values(totals).sort((a,b)=>b.payment-a.payment);
}}

function buildMonthlySalesByYear() {{
  const byYear = {{}};
  Object.entries(HISTORICAL_MONTHLY).forEach(([year, rows]) => {{
    byYear[year] = {{}};
    rows.forEach(r => byYear[year][r.month] = Number(r.sales || 0));
  }});
  Object.entries(HISTORICAL_DAILY_MONTHLY).forEach(([year, rows]) => {{
    if (!byYear[year]) byYear[year] = {{}};
    rows.forEach(r => {{
      if (!byYear[year][r.month]) byYear[year][r.month] = Number(r.sales || 0);
    }});
  }});
  rawRows.forEach(r => {{
    (r.daily || []).forEach(d => {{
      const year = d.date.slice(0,4);
      const month = d.date.slice(5,7);
      if (!byYear[year]) byYear[year] = {{}};
      byYear[year][month] = (byYear[year][month] || 0) + Number(d.payment || 0);
    }});
  }});
  const monthlyExisting = {{}};
  rawRows.forEach(r => {{
    (r.daily || []).forEach(d => {{
      const key = d.date + '|' + (r.retailer || '');
      monthlyExisting[key] = (monthlyExisting[key] || 0) + Number(d.payment || 0);
    }});
  }});
  MANUAL_SALES.forEach(day => {{
    const year = day.date.slice(0,4);
    const month = day.date.slice(5,7);
    const payment = (day.retailers || []).reduce((sum, r) => {{
      const existing = monthlyExisting[day.date + '|' + (r.retailer || '')] || 0;
      return sum + Math.max(0, Number(r.payment || 0) - existing);
    }}, 0);
    if (!byYear[year]) byYear[year] = {{}};
    byYear[year][month] = (byYear[year][month] || 0) + payment;
  }});
  return byYear;
}}

const monthlyByYear = buildMonthlySalesByYear();
const compareYears = Object.keys(monthlyByYear).sort();
let compareChartInst = null;

// ── Chart 기본값 ───────────────────────────────────────────
Chart.defaults.color = '#8a94a6';
Chart.defaults.font.family = "'Pretendard',sans-serif";
Chart.defaults.font.size = 11;
const ttip = {{backgroundColor:'#1e2535',borderColor:'#2b3548',borderWidth:1,titleColor:'#fff',bodyColor:'#c8d0e0',padding:10,callbacks:{{label:ctx=>' '+fmt(ctx.raw)}}}};
const scl  = {{x:{{grid:{{color:'#e2e6ed'}},ticks:{{maxRotation:0}}}},y:{{grid:{{color:'#e2e6ed'}},ticks:{{callback:v=>fmt(v)}}}}}};

// ── OVERVIEW: 최근 3일 일별 차트 ──────────────────────────
(function(){{
  const allDaily = buildDailyMap('');
  const recent3  = allDaily.slice(-3);
  const labels   = recent3.map(d => fmtD(d.date));

  if (recent3.length) {{
    const first = recent3[0].date, last = recent3[recent3.length-1].date;
    document.getElementById('recentDateRange').textContent = fmtD(first)+' ~ '+fmtD(last);
  }}

  new Chart(document.getElementById('paymentChart'),{{
    type:'bar',
    data:{{labels, datasets:[{{data:recent3.map(d=>d.payment),backgroundColor:'rgba(59,130,246,0.2)',borderColor:'#3b82f6',borderWidth:1.5,borderRadius:5}}]}},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:ttip,title:{{display:true,text:'일별 매출',color:'#4a5568',padding:{{bottom:8}}}}}},scales:scl}}
  }});
  new Chart(document.getElementById('qtyChart'),{{
    type:'bar',
    data:{{labels, datasets:[{{data:recent3.map(d=>d.qty),backgroundColor:'rgba(16,185,129,0.18)',borderColor:'#10b981',borderWidth:1.5,borderRadius:5}}]}},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:ttip,title:{{display:true,text:'일별 판매수량',color:'#4a5568',padding:{{bottom:8}}}}}},scales:scl}}
  }});

  const retailerPalette = ['#3b82f6','#10b981','#f59e0b','#8b5cf6','#ef4444','#06b6d4','#64748b'];
  const recentDates = recent3.map(d => d.date);
  const retailerLineData = RETAILERS.map((retailer, idx) => {{
    const retailerDailyMap = Object.fromEntries(
      buildDailyMap(retailer).map(d => [d.date, Number(d.payment || 0)])
    );
    return {{
      label: retailer,
      data: recentDates.map(date => retailerDailyMap[date] || 0),
      borderColor: retailerPalette[idx % retailerPalette.length],
      backgroundColor: retailerPalette[idx % retailerPalette.length],
      borderWidth: 2,
      pointRadius: 2,
      pointHoverRadius: 5,
      tension: 0.35,
      fill: false
    }};
  }}).filter(ds => ds.data.some(v => v !== 0));

  new Chart(document.getElementById('recentRetailerChart'), {{
    type:'line',
    data:{{labels, datasets:retailerLineData}},
    options:{{
      responsive:true,
      maintainAspectRatio:false,
      interaction:{{mode:'index',intersect:false}},
      plugins:{{
        legend:{{display:true, position:'bottom', labels:{{usePointStyle:true, boxWidth:8, boxHeight:8}}}},
        tooltip:ttip,
        title:{{display:true,text:'유통사별 일별 매출',color:'#4a5568',padding:{{bottom:8}}}}
      }},
      scales:scl
    }}
  }});

  // ── 최근 7일 상위 상품 ─────────────────────────────────
  const monthMap = {{}};
  allDaily.forEach(d => {{
    const m = d.date.slice(5,7);
    if (!monthMap[m]) monthMap[m] = 0;
    monthMap[m] += d.payment;
  }});
  const latestMonth = allDaily.length ? allDaily[allDaily.length - 1].date.slice(5,7) : null;
  const monthlyRows = PREV_YEAR_MONTHLY
  .filter(row => !latestMonth || row.month <= latestMonth)
  .map(row => {{
    const current = monthMap[row.month] || 0;
    const prev = row.sales2025 || 0;
    const growth = prev ? (current - prev) / prev * 100 : null;
    return {{month: row.month, current, prev, growth}};
  }});
  new Chart(document.getElementById('monthlyGrowthChart'), {{
    type:'bar',
    data:{{
      labels:monthlyRows.map(r=>r.month+'월'),
      datasets:[{{data:monthlyRows.map(r=>r.growth ?? 0),backgroundColor:monthlyRows.map(r=>(r.growth ?? 0)>=0?'rgba(16,185,129,.22)':'rgba(239,68,68,.22)'),borderColor:monthlyRows.map(r=>(r.growth ?? 0)>=0?'#10b981':'#ef4444'),borderWidth:1.5,borderRadius:5}}]
    }},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:{{...ttip,callbacks:{{label:ctx=>' '+pct(ctx.raw)}}}}}},scales:{{x:scl.x,y:{{grid:{{color:'#e2e6ed'}},ticks:{{callback:v=>pct(v)}}}}}}}}
  }});
  document.getElementById('monthlyGrowthRows').innerHTML = monthlyRows.map(r=>`
    <tr>
      <td class="td-main">${{r.month}}월</td>
      <td class="num">${{fmt(r.current)}}</td>
      <td class="num">${{fmt(r.prev)}}</td>
      <td class="num" style="color:${{(r.growth ?? 0)>=0?'var(--teal)':'var(--red)'}}">${{r.growth===null?'-':pct(r.growth)}}</td>
    </tr>
  `).join('');

  const allDates = allDaily.map(d=>d.date);
  const cutoff   = allDates.length >= 7 ? allDates[allDates.length-7] : allDates[0];
  const payMap   = {{}};
  rawRows.forEach(r => {{
    (r.daily||[]).forEach(d => {{
      if (d.date >= cutoff) {{
        if (!payMap[r.standard_name]) payMap[r.standard_name] = 0;
        payMap[r.standard_name] += d.payment;
      }}
    }});
  }});
  const top10 = Object.entries(payMap).sort((a,b)=>b[1]-a[1]).slice(0,10);
  const maxPay = top10[0]?.[1] || 1;
  document.getElementById('rankList').innerHTML = top10.map(([name,pay],i)=>
    `<div class="rank-row">
      <div class="rank-n">${{String(i+1).padStart(2,'0')}}</div>
      <div class="rank-name" title="${{name}}">${{name}}</div>
      <div class="rank-bar"><div style="width:${{(pay/maxPay*100).toFixed(1)}}%"></div></div>
      <div class="rank-val">${{fmt(pay)}}</div>
    </div>`
  ).join('');

  // ── 상태 미니카드 ──────────────────────────────────────
  const srcTypes = rawRows.reduce((a,r)=>{{a[r.source_type]=(a[r.source_type]||0)+1;return a;}},{{}});
  const activeRetailers = [...new Set(rawRows.map(r=>r.retailer))];
  document.getElementById('statusMini').innerHTML = `
    <div class="mini"><div class="mini-label">집계 SKU</div><div class="mini-value">${{fmt(rawRows.length)}}</div></div>
    <div class="mini"><div class="mini-label">총 주문수</div><div class="mini-value">${{fmt(totalOrders)}}</div></div>
    <div class="mini"><div class="mini-label">총 판매수량</div><div class="mini-value">${{fmt(totalQty)}}</div></div>
    <div class="mini"><div class="mini-label">유통사 수</div><div class="mini-value">${{activeRetailers.length}}</div></div>
    <div class="mini"><div class="mini-label">단품</div><div class="mini-value">${{srcTypes['단품']||0}}</div></div>
    <div class="mini"><div class="mini-label">세트분해</div><div class="mini-value">${{srcTypes['세트분해']||0}}</div></div>
  `;

  // 미매칭
  const um = rawRows.filter(r=>!isMatchedRow(r));
  document.getElementById('unmatchedRows').innerHTML = um.length
    ? um.map(r=>`<tr><td><span class="badge badge-blue">${{r.retailer}}</span></td><td class="td-mono">${{r.mall_no}}</td><td class="td-main">${{r.standard_name||r.name}}</td><td class="num">${{fmt(r.qty)}}</td><td class="num">${{fmt(r.payment)}}</td></tr>`).join('')
    : '<tr><td colspan="5" style="text-align:center;color:var(--ink3);padding:28px">✓ 미매칭 항목 없음</td></tr>';
}})();

// ── CALENDAR TAB ───────────────────────────────────────────
// 매출 비교 탭
function initCompareControls() {{
  const base = document.getElementById('compareBaseYear');
  const target = document.getElementById('compareTargetYear');
  const options = compareYears.map(y => `<option value="${{y}}">${{y}}년</option>`).join('');
  base.innerHTML = options;
  target.innerHTML = options;
  base.value = compareYears.includes('2026') ? '2026' : compareYears[compareYears.length - 1];
  target.value = compareYears.includes('2025') ? '2025' : compareYears[0];
}}

function compareValue(year, month) {{
  return Number(monthlyByYear[year]?.[month] || 0);
}}

function renderCompare() {{
  const baseYear = document.getElementById('compareBaseYear').value;
  const targetYear = document.getElementById('compareTargetYear').value;
  const selectedMonth = document.getElementById('compareMonth').value;
  const months = selectedMonth
    ? [selectedMonth]
    : Array.from(new Set([
        ...Object.keys(monthlyByYear[baseYear] || {{}}),
        ...Object.keys(monthlyByYear[targetYear] || {{}})
      ])).sort();

  const rows = months.map(month => {{
    const base = compareValue(baseYear, month);
    const target = compareValue(targetYear, month);
    const diff = base - target;
    const growth = target ? diff / target * 100 : null;
    return {{month, base, target, diff, growth}};
  }});

  const baseTotal = rows.reduce((a,r)=>a+r.base,0);
  const targetTotal = rows.reduce((a,r)=>a+r.target,0);
  const totalGrowth = targetTotal ? (baseTotal - targetTotal) / targetTotal * 100 : null;
  const scope = selectedMonth ? `${{Number(selectedMonth)}}월` : '전체 월';

  document.getElementById('compareBaseLabel').textContent = `${{baseYear}}년 ${{scope}} 매출`;
  document.getElementById('compareTargetLabel').textContent = `${{targetYear}}년 ${{scope}} 매출`;
  document.getElementById('compareBaseHead').textContent = `${{baseYear}}년`;
  document.getElementById('compareTargetHead').textContent = `${{targetYear}}년`;
  document.getElementById('compareBaseSales').textContent = fmt(baseTotal);
  document.getElementById('compareTargetSales').textContent = fmt(targetTotal);
  document.getElementById('compareGrowth').textContent = totalGrowth === null ? '-' : pct(totalGrowth);
  document.getElementById('compareGrowth').style.color = (totalGrowth ?? 0) >= 0 ? 'var(--teal)' : 'var(--red)';

  if (compareChartInst) compareChartInst.destroy();
  compareChartInst = new Chart(document.getElementById('compareChart'), {{
    type:'bar',
    data:{{
      labels:rows.map(r=>Number(r.month)+'월'),
      datasets:[
        {{label:`${{baseYear}}년`,data:rows.map(r=>r.base),backgroundColor:'rgba(59,130,246,.22)',borderColor:'#3b82f6',borderWidth:1.5,borderRadius:5}},
        {{label:`${{targetYear}}년`,data:rows.map(r=>r.target),backgroundColor:'rgba(16,185,129,.18)',borderColor:'#10b981',borderWidth:1.5,borderRadius:5}}
      ]
    }},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:true,position:'bottom'}},tooltip:ttip}},scales:scl}}
  }});

  document.getElementById('compareRows').innerHTML = rows.map(r=>`
    <tr>
      <td class="td-main">${{Number(r.month)}}월</td>
      <td class="num">${{fmt(r.base)}}</td>
      <td class="num">${{fmt(r.target)}}</td>
      <td class="num" style="color:${{r.diff>=0?'var(--teal)':'var(--red)'}}">${{fmt(r.diff)}}</td>
      <td class="num" style="color:${{(r.growth ?? 0)>=0?'var(--teal)':'var(--red)'}}">${{r.growth===null?'-':pct(r.growth)}}</td>
    </tr>
  `).join('');
}}

initCompareControls();
['compareBaseYear','compareTargetYear','compareMonth'].forEach(id=>{{
  document.getElementById(id).addEventListener('change',renderCompare);
}});

let dcInst = null;
let calRetailer = '';
let calYear = '2026';

function renderDaily() {{
  const s = document.getElementById('dateStart').value;
  const e = document.getElementById('dateEnd').value;
  const allDaily = calYear === '2026' ? buildDailyMap(calRetailer) : buildHistoricalDailyMap(calYear, calRetailer);
  const f = allDaily.filter(d=>(!s||d.date>=s)&&(!e||d.date<=e));

  document.getElementById('dayPayment').textContent = fmt(f.reduce((a,d)=>a+d.payment,0));
  document.getElementById('dayOrders').textContent  = calYear === '2026' ? fmt(f.reduce((a,d)=>a+d.orders,0)) : '-';
  document.getElementById('dayQty').textContent     = calYear === '2026' ? fmt(f.reduce((a,d)=>a+d.qty,0)) : '-';

  if (dcInst) dcInst.destroy();
  dcInst = new Chart(document.getElementById('dailyChart'),{{
    type:'line',
    data:{{labels:f.map(d=>fmtD(d.date)),datasets:[{{data:f.map(d=>d.payment),borderColor:'#3b82f6',backgroundColor:'rgba(59,130,246,0.07)',borderWidth:1.5,fill:true,tension:0.4,pointRadius:2,pointHoverRadius:5}}]}},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:ttip}},scales:scl}}
  }});

  document.getElementById('dailyRows').innerHTML = !f.length
    ? '<tr><td colspan="6" style="text-align:center;color:var(--ink3);padding:24px">데이터 없음</td></tr>'
    : f.map(d=>`<tr>
        <td class="td-main td-mono">${{d.date}}</td>
        <td class="num">${{fmt(d.payment)}}</td>
        <td class="num">${{calYear === '2026' ? fmt(d.gross) : '-'}}</td>
        <td class="num">${{calYear === '2026' ? fmt(d.orders) : '-'}}</td>
        <td class="num">${{calYear === '2026' ? fmt(d.qty) : '-'}}</td>
        <td class="num">${{calYear === '2026' && d.orders ? fmt(d.payment/d.orders) : '-'}}</td>
      </tr>`).join('');
}}

// 유통사 서브탭
document.querySelectorAll('#calRetailerTabs .sub-tab').forEach(btn => {{
  btn.addEventListener('click', () => {{
    document.querySelectorAll('#calRetailerTabs .sub-tab').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    calRetailer = btn.dataset.retailer;
    renderDaily();
  }});
}});
document.getElementById('calYear').addEventListener('change', e => {{
  calYear = e.target.value;
  renderDaily();
}});
document.getElementById('dateStart').addEventListener('change', renderDaily);
document.getElementById('dateEnd').addEventListener('change', renderDaily);
document.getElementById('clearDate').addEventListener('click', () => {{
  document.getElementById('calYear').value = '2026';
  calYear = '2026';
  document.getElementById('dateStart').value = '';
  document.getElementById('dateEnd').value = '';
  renderDaily();
}});

// ── RETAILER TAB ───────────────────────────────────────────
let rcInst = null;
let retailerYear = '2026';
function renderRetailer() {{
  const colors = ['#3b82f6','#10b981','#f59e0b','#ef4444','#8b5cf6','#06b6d4','#ec4899'];
  const filterVal = document.getElementById('retailerFilter').value;
  const baseData = retailerYear === '2026' ? retailerAll : buildHistoricalRetailerSummary(retailerYear);
  const data = filterVal ? baseData.filter(r=>r.retailer===filterVal) : baseData;

  if (rcInst) rcInst.destroy();
  rcInst = new Chart(document.getElementById('retailerChart'),{{
    type:'bar',
    data:{{
      labels:data.map(r=>r.retailer),
      datasets:[{{data:data.map(r=>r.payment),backgroundColor:data.map((_,i)=>colors[i%colors.length]+'33'),borderColor:data.map((_,i)=>colors[i%colors.length]),borderWidth:1.5,borderRadius:5}}]
    }},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:false}},tooltip:ttip}},scales:scl}}
  }});

  document.getElementById('retailerRows').innerHTML = data.map(r=>{{
    const disc = retailerYear === '2026' && r.validGross>0?(1-r.validPayment/r.validGross)*100:null;
    return `<tr>
      <td class="td-main"><span class="badge badge-blue">${{r.retailer}}</span></td>
      <td class="num">${{fmt(r.payment)}}</td>
      <td class="num">${{retailerYear === '2026' ? fmt(r.gross) : '-'}}</td>
      <td class="num">${{retailerYear === '2026' ? fmt(r.orders) : '-'}}</td>
      <td class="num">${{retailerYear === '2026' ? fmt(r.qty) : '-'}}</td>
      <td class="num">${{retailerYear === '2026' && r.orders?fmt(r.payment/r.orders):'-'}}</td>
      <td class="num">${{disc===null?'-':pct(disc)}}</td>
    </tr>`;
  }}).join('');
}}
document.getElementById('retailerYear').addEventListener('change', e => {{
  retailerYear = e.target.value;
  renderRetailer();
}});
document.getElementById('retailerFilter').addEventListener('change', renderRetailer);
document.getElementById('clearRetailer').addEventListener('click', ()=>{{
  document.getElementById('retailerYear').value='2026';
  retailerYear = '2026';
  document.getElementById('retailerFilter').value='';
  renderRetailer();
}});

// ── DETAIL TAB ─────────────────────────────────────────────
function initDetailFilters() {{
  const fill = (id, label, values) => {{
    const el = document.getElementById(id);
    const current = el.value;
    el.innerHTML = `<option value="">${{label}}</option>` + values.map(v=>`<option value="${{v}}">${{v}}</option>`).join('');
    if (values.includes(current)) el.value = current;
  }};
  fill('seasonFilter', '전체 시즌', uniqSorted(rawRows.map(r=>r.season)));
  fill('categoryLargeFilter', '전체 복종(대)', uniqSorted(rawRows.map(r=>r.category_large)));
  fill('categorySmallFilter', '전체 복종(소)', uniqSorted(rawRows.map(r=>r.category_small)));
}}

function renderDetail() {{
  const q   = (document.getElementById('q').value||'').toLowerCase();
  const seasonFilter = document.getElementById('seasonFilter').value;
  const categoryLargeFilter = document.getElementById('categoryLargeFilter').value;
  const categorySmallFilter = document.getElementById('categorySmallFilter').value;
  const sb  = document.getElementById('sortBy').value;

  let baseRows = [...rawRows];
  if (seasonFilter) baseRows = baseRows.filter(r=>r.season===seasonFilter);
  if (categoryLargeFilter) baseRows = baseRows.filter(r=>r.category_large===categoryLargeFilter);
  if (categorySmallFilter) baseRows = baseRows.filter(r=>r.category_small===categorySmallFilter);

  const grouped = {{}};
  baseRows.forEach(r=>{{
    const key = r.match_sku || r.standard_name || `${{r.name}}_${{r.color}}_${{r.size}}`;
    if (!grouped[key]) grouped[key] = {{
      key,
      match_sku:r.match_sku || '',
      standard_name:r.standard_name || r.name || '-',
      retailers:new Set(),
      seasons:new Set(),
      categoryLarge:new Set(),
      categorySmall:new Set(),
      qty:0,
      gross:0,
      payment:0,
      orders:0,
      validGross:0,
      validPayment:0,
      stock_qty:0,
      received_qty:0,
      stock_known:false,
      dailyMap:{{}}
    }};
    const g = grouped[key];
    g.retailers.add(r.retailer || '-');
    g.seasons.add(r.season || '-');
    g.categoryLarge.add(r.category_large || '-');
    g.categorySmall.add(r.category_small || '-');
    g.qty += Number(r.qty || 0);
    g.gross += Number(r.gross || 0);
    g.payment += Number(r.payment || 0);
    g.orders += Number(r.orders || 0);
    if (validDiscount(r.gross, r.payment) !== null) {{
      g.validGross += Number(r.gross || 0);
      g.validPayment += Number(r.payment || 0);
    }}
    g.stock_qty = Math.max(g.stock_qty, Number(r.stock_qty || 0));
    g.received_qty = Math.max(g.received_qty, Number(r.received_qty || 0));
    if (r.stock_barcode || r.stock_name) g.stock_known = true;
    (r.daily || []).forEach(d=>{{
      if (!g.dailyMap[d.date]) g.dailyMap[d.date] = {{date:d.date,qty:0,gross:0,payment:0,orders:0}};
      g.dailyMap[d.date].qty += Number(d.qty || 0);
      g.dailyMap[d.date].gross += Number(d.gross || 0);
      g.dailyMap[d.date].payment += Number(d.payment || 0);
      g.dailyMap[d.date].orders += Number(d.orders || 0);
    }});
  }});

  let rows = Object.values(grouped).map(g=>{{
    const avg_unit = g.qty ? g.payment / g.qty : 0;
    return {{
      ...g,
      retailers:Array.from(g.retailers),
      seasons:Array.from(g.seasons),
      categoryLarge:Array.from(g.categoryLarge),
      categorySmall:Array.from(g.categorySmall),
      daily:Object.values(g.dailyMap).sort((a,b)=>a.date.localeCompare(b.date)),
      avg_unit
    }};
  }});

  if (q) rows = rows.filter(r=>
    (r.standard_name||'').toLowerCase().includes(q) ||
    (r.match_sku||'').toLowerCase().includes(q)
  );

  const detailDates = buildDailyMap('').map(d=>d.date);
  const detailCutoff = detailDates.length >= 7 ? detailDates[detailDates.length-7] : detailDates[0];
  const filteredQty = rows.reduce((a,r)=>a + Number(r.qty || 0), 0);
  const filteredStock = rows.reduce((a,r)=>a + Math.max(0, Number(r.stock_qty || 0)), 0);
  const weeklyQty = rows.reduce((sum,r)=>sum + (r.daily||[]).filter(d=>!detailCutoff || d.date>=detailCutoff).reduce((a,d)=>a + Number(d.qty || 0), 0), 0);
  const stockRate = filteredQty + filteredStock ? filteredQty / (filteredQty + filteredStock) * 100 : null;
  const weeklyRate = weeklyQty + filteredStock ? weeklyQty / (weeklyQty + filteredStock) * 100 : null;
  document.getElementById('detailStockQty').textContent = fmt(filteredStock);
  document.getElementById('detailStockRate').textContent = stockRate === null ? '-' : pct(stockRate);
  document.getElementById('detailWeeklyRate').textContent = weeklyRate === null ? '-' : pct(weeklyRate);

  rows.forEach(r=>{{
    const stockBase = Number(r.stock_qty || 0);
    const weekly = (r.daily||[]).filter(d=>!detailCutoff || d.date>=detailCutoff).reduce((a,d)=>a + Number(d.qty || 0), 0);
    r.weekly_qty = weekly;
    r.weekly_rate = weekly + stockBase > 0 ? weekly / (weekly + stockBase) * 100 : 0;
    r.stock_rate = r.qty + stockBase > 0 ? r.qty / (r.qty + stockBase) * 100 : 0;
    r.reorder_reasons = [];
    if (r.stock_known && r.weekly_rate >= 7) r.reorder_reasons.push('주간 판매율 7% 이상');
    if (r.stock_known && r.stock_rate >= 20) r.reorder_reasons.push('재고 소진율 20% 이상');
  }});

  const reorderRows = rows
    .filter(r=>r.reorder_reasons.length)
    .sort((a,b)=>b.weekly_rate-a.weekly_rate || b.stock_rate-a.stock_rate)
    .slice(0, 8);
  document.getElementById('reorderSummary').textContent = `${{rows.filter(r=>r.reorder_reasons.length).length}}개 상품`;
  document.getElementById('reorderAlerts').innerHTML = reorderRows.length
    ? reorderRows.map(r=>`<div class="rank-item"><div><div class="rank-name">${{r.standard_name}}</div><div class="rank-meta">${{r.reorder_reasons.join(' · ')}} · 현재고 ${{fmt(r.stock_qty)}} · 주간 ${{fmt(r.weekly_qty)}}개</div></div><div class="rank-value">${{pct(r.weekly_rate)}}</div></div>`).join('')
    : '<div style="padding:14px;color:var(--ink3)">현재 기준 리오더 알림 상품 없음</div>';

  if (sb==='payment') rows.sort((a,b)=>b.payment-a.payment);
  else if (sb==='qty') rows.sort((a,b)=>b.qty-a.qty);
  else if (sb==='weeklyRate') rows.sort((a,b)=>b.weekly_rate-a.weekly_rate || b.stock_rate-a.stock_rate);
  else rows.sort((a,b)=>(a.standard_name||'').localeCompare(b.standard_name||''));

  const tbody = document.getElementById('detailRows');
  const filterBody = document.getElementById('detailFilterRows');
  filterBody.innerHTML = `<tr class="table-filter-row">
    <td><select id="tableAlertFilter"><option value="">전체</option><option value="확인">확인</option><option value="정상">정상</option><option value="재고확인">재고확인</option></select></td>
    <td><input id="tableSkuFilter" placeholder="바코드"></td>
    <td><input id="tableNameFilter" placeholder="상품명"></td>
    <td><input id="tableRetailerFilter" placeholder="유통사"></td>
    <td><select id="tableSeasonFilter"><option value="">전체</option>${{uniqSorted(rows.flatMap(r=>r.seasons)).map(v=>`<option value="${{v}}">${{v}}</option>`).join('')}}</select></td>
    <td><select id="tableCategoryFilter"><option value="">전체</option>${{uniqSorted(rows.flatMap(r=>r.categoryLarge.concat(r.categorySmall))).map(v=>`<option value="${{v}}">${{v}}</option>`).join('')}}</select></td>
    <td colspan="8"></td>
  </tr>`;

  const tableFilters = {{
    alert:'',
    sku:'',
    name:'',
    retailer:'',
    season:'',
    category:''
  }};
  const applyTableFilters = () => {{
    tableFilters.alert = document.getElementById('tableAlertFilter').value;
    tableFilters.sku = (document.getElementById('tableSkuFilter').value||'').toLowerCase();
    tableFilters.name = (document.getElementById('tableNameFilter').value||'').toLowerCase();
    tableFilters.retailer = (document.getElementById('tableRetailerFilter').value||'').toLowerCase();
    tableFilters.season = document.getElementById('tableSeasonFilter').value;
    tableFilters.category = document.getElementById('tableCategoryFilter').value;
    drawDetailRows(rows.filter(r=>{{
      const alertText = !r.stock_known ? '재고확인' : r.reorder_reasons.length ? '확인' : '정상';
      if (tableFilters.alert && alertText !== tableFilters.alert) return false;
      if (tableFilters.sku && !(r.match_sku||'').toLowerCase().includes(tableFilters.sku)) return false;
      if (tableFilters.name && !(r.standard_name||'').toLowerCase().includes(tableFilters.name)) return false;
      if (tableFilters.retailer && !r.retailers.join(',').toLowerCase().includes(tableFilters.retailer)) return false;
      if (tableFilters.season && !r.seasons.includes(tableFilters.season)) return false;
      if (tableFilters.category && !r.categoryLarge.concat(r.categorySmall).includes(tableFilters.category)) return false;
      return true;
    }}));
  }};
  ['tableAlertFilter','tableSkuFilter','tableNameFilter','tableRetailerFilter','tableSeasonFilter','tableCategoryFilter'].forEach(id=>{{
    const el = document.getElementById(id);
    el.addEventListener('input', applyTableFilters);
    el.addEventListener('change', applyTableFilters);
  }});

  if (!rows.length) {{ tbody.innerHTML='<tr><td colspan="14" style="text-align:center;color:var(--ink3);padding:28px">검색 결과 없음</td></tr>'; return; }}

  function drawDetailRows(displayRows) {{
  if (!displayRows.length) {{
    tbody.innerHTML='<tr><td colspan="14" style="text-align:center;color:var(--ink3);padding:28px">검색 결과 없음</td></tr>';
    return;
  }}
  tbody.innerHTML = displayRows.map(r=>{{
    const disc    = r.validGross > 0 && r.validPayment <= r.validGross ? (1 - r.validPayment / r.validGross) * 100 : null;
    const stockQty= r.stock_qty ?? '-';
    const alertB  = !r.stock_known ? 'badge-amber' : r.reorder_reasons.length ? 'badge-red' : 'badge-green';
    const alertT  = !r.stock_known ? '재고확인' : r.reorder_reasons.length ? '확인' : '정상';
    return `<tr>
      <td><span class="badge ${{alertB}}" title="${{r.reorder_reasons.join(' / ')}}">${{alertT}}</span></td>
      <td class="td-mono">${{r.match_sku || '-'}}</td>
      <td class="td-main" style="max-width:170px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${{r.standard_name}}">${{r.standard_name||'-'}}</td>
      <td style="font-size:12px;color:var(--ink3);max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${{r.retailers.join(', ')}}</td>
      <td><span class="badge badge-blue">${{r.seasons.join(', ')}}</span></td>
      <td><span class="badge badge-indigo">${{r.categoryLarge.concat(r.categorySmall).filter(v=>v && v !== '-').join(' / ') || '-'}}</span></td>
      <td class="num">${{fmt(r.qty)}}</td>
      <td class="num" style="color:${{r.weekly_rate>=7?'var(--red)':r.weekly_rate>=4?'var(--amber)':'var(--ink2)'}}">${{r.weekly_rate>0?pct(r.weekly_rate):'-'}}</td>
      <td class="num">${{stockQty!=='-'?fmt(stockQty):'-'}}</td>
      <td class="num" style="color:${{r.stock_rate>=20?'var(--red)':r.stock_rate>=12?'var(--amber)':'var(--ink2)'}}">${{r.stock_rate>0?pct(r.stock_rate):'-'}}</td>
      <td class="num">${{fmt(r.gross)}}</td>
      <td class="num" style="color:var(--blue2);font-weight:600">${{fmt(r.payment)}}</td>
      <td class="num">${{fmt(r.avg_unit)}}</td>
      <td class="num" style="color:${{disc!==null&&disc>50?'var(--red)':disc!==null&&disc>30?'var(--amber)':'var(--ink2)'}}">${{disc===null?'-':pct(disc)}}</td>
    </tr>`;
  }}).join('');
  }}
  drawDetailRows(rows);
}}
initDetailFilters();
['q','seasonFilter','categoryLargeFilter','categorySmallFilter','sortBy'].forEach(id=>{{
  document.getElementById(id).addEventListener('input',renderDetail);
  document.getElementById(id).addEventListener('change',renderDetail);
}});

// ?? NAV TABS ???????????????????????????????????????????????
function activateTab(tab) {{
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  const nav = document.querySelector(`.nav-item[data-tab="${{tab}}"]`);
  if (nav) nav.classList.add('active');
  const panel = document.getElementById(tab);
  if (panel) panel.classList.add('active');
  if(tab==='compare')   renderCompare();
  if(tab==='calendar')  renderDaily();
  if(tab==='retailer')  renderRetailer();
  if(tab==='detail')    renderDetail();
}}

document.querySelectorAll('.nav-item[data-tab]').forEach(btn=>{{
  btn.addEventListener('click',()=>activateTab(btn.dataset.tab));
}});

document.querySelectorAll('[data-tab-jump]').forEach(el=>{{
  el.addEventListener('click',()=>{{
    activateTab(el.dataset.tabJump);
    if (el.dataset.focus === 'unmatched') {{
      setTimeout(()=>document.getElementById('unmatchedSection')?.scrollIntoView({{behavior:'smooth',block:'start'}}), 50);
    }}
  }});
}});

// init
renderDaily();
</script>
</body>
</html>"""

OUT_FILE.write_text(html, encoding="utf-8")
print(f"✅ index.html 생성 완료 ({OUT_FILE.stat().st_size:,} bytes)")
