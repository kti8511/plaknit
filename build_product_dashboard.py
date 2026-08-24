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
import json, datetime, os, pathlib

DATA_FILE = pathlib.Path("data.json")
HISTORICAL_DAILY_FILE = pathlib.Path("historical_daily.json")
MANUAL_SALES_FILE = pathlib.Path("manual_sales_updates.json")
OUT_FILE  = pathlib.Path("index.html")
PUBLIC_OUT_FILE = pathlib.Path("public") / "index.html"

TODO_SUPABASE_PROJECT = (
    os.environ.get("PLAKNIT_SUPABASE_URL")
    or os.environ.get("NEXT_PUBLIC_SUPABASE_URL")
    or os.environ.get("VITE_SUPABASE_URL")
    or os.environ.get("SUPABASE_URL")
    or ""
)
TODO_SUPABASE_PROJECT = TODO_SUPABASE_PROJECT.rstrip("/")
TODO_SUPABASE_REST = f"{TODO_SUPABASE_PROJECT}/rest/v1" if TODO_SUPABASE_PROJECT else ""
TODO_SUPABASE_KEY = (
    os.environ.get("PLAKNIT_SUPABASE_ANON_KEY")
    or os.environ.get("NEXT_PUBLIC_SUPABASE_ANON_KEY")
    or os.environ.get("NEXT_PUBLIC_SUPABASE_PUBLISHABLE_KEY")
    or os.environ.get("VITE_SUPABASE_ANON_KEY")
    or os.environ.get("SUPABASE_ANON_KEY")
    or os.environ.get("SUPABASE_PUBLISHABLE_KEY")
    or ""
)
TODO_SUPABASE_BUCKET = os.environ.get("PLAKNIT_SUPABASE_BUCKET") or os.environ.get("SUPABASE_BUCKET") or "todo-files"

with DATA_FILE.open(encoding="utf-8-sig") as f:
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
JSON_KWARGS = {"ensure_ascii": False, "separators": (",", ":")}

rows_json = json.dumps(rows, **JSON_KWARGS)
historical_daily_json = json.dumps(historical_daily, **JSON_KWARGS)
manual_sales_json = json.dumps(manual_sales_updates, **JSON_KWARGS)

# ★ 2026 일일판매정리(4월).xlsx 첫 시트의 몰별 월 목표 매출.
# 기타몰은 별도 목표가 없는 유통사(애슬러, 롯데온, ABC마트, 러너블, 기타 등)의 합산 목표다.
SALES_TARGET_DEFAULTS = {
    "2026": {
        "01": {"자사몰": 15795630, "무신사": 5748871.5, "글로리어스워커": 7000000, "4XR": 2246880, "29cm": 405144, "기타몰": 500000},
        "02": {"자사몰": 11808300, "무신사": 9469110, "글로리어스워커": 2000000, "4XR": 2849376, "29cm": 287642, "기타몰": 0},
        "03": {"자사몰": 17273160, "무신사": 11052276, "글로리어스워커": 7000000, "4XR": 3664470, "29cm": 143630, "기타몰": 1000000},
        "04": {"자사몰": 58495000, "무신사": 24174517.5, "글로리어스워커": 10000000, "4XR": 3177560, "29cm": 458166, "기타몰": 1000000},
        "05": {"자사몰": 62507425, "무신사": 41056875, "글로리어스워커": 8000000, "4XR": 4450720, "29cm": 145800, "기타몰": 1000000},
        "06": {"자사몰": 65453788.8, "무신사": 65725290, "글로리어스워커": 10000000, "4XR": 4857200, "29cm": 1277366, "기타몰": 1000000},
        "07": {"자사몰": 62175808, "무신사": 56030015, "글로리어스워커": 10000000, "4XR": 3932488, "29cm": 850544, "기타몰": 1000000},
        "08": {"자사몰": 85804550, "무신사": 41865787.5, "글로리어스워커": 0, "4XR": 3654300, "29cm": 188274, "기타몰": 500000},
        "09": {"자사몰": 28801750, "무신사": 19924535, "글로리어스워커": 0, "4XR": 3088000, "29cm": 200648, "기타몰": 0},
        "10": {"자사몰": 20721180, "무신사": 17174784, "글로리어스워커": 0, "4XR": 1803060, "29cm": 118380, "기타몰": 0},
        "11": {"자사몰": 21936980, "무신사": 21668370, "글로리어스워커": 10000000, "4XR": 664800, "29cm": 705814, "기타몰": 1000000},
        "12": {"자사몰": 16000000, "무신사": 16000000, "글로리어스워커": 750000, "4XR": 750000, "29cm": 600000, "기타몰": 0},
    }
}
sales_target_defaults_json = json.dumps(SALES_TARGET_DEFAULTS, **JSON_KWARGS)

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
prev_year_monthly_json = json.dumps(PREV_YEAR_MONTHLY, **JSON_KWARGS)

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
historical_monthly_json = json.dumps(HISTORICAL_MONTHLY, **JSON_KWARGS)

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
historical_daily_monthly_json = json.dumps(historical_daily_monthly, **JSON_KWARGS)

# 유통사 목록 (고정)
RETAILERS = ["자사몰", "무신사", "29cm", "글로리어스워커", "4XR", "애슬러", "롯데온", "ABC마트", "러너블", "기타"]

html = f"""<!doctype html>
<html lang="ko">
<head>
<meta charset="utf-8"/>
<meta name="viewport" content="width=device-width, initial-scale=1"/>
<title>상품DATA 운영 대시보드</title>
<link href="https://fonts.googleapis.com/css2?family=Pretendard:wght@300;400;500;600;700;800&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet"/>
<script src="vendor/chart.umd.min.js"></script>
<style>
:root{{
  --standard-name-col-width:280px;
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
.nav-item{{display:flex;flex-direction:column;justify-content:center;height:100%;padding:0 18px;cursor:pointer;border-bottom:3px solid transparent;transition:all 0.15s;white-space:nowrap;text-decoration:none;}}
.nav-item:hover{{background:rgba(255,255,255,0.05);}}
.nav-item.active{{border-bottom-color:var(--blue);}}
.nav-label{{font-size:13px;font-weight:600;color:rgba(255,255,255,0.85);display:flex;align-items:center;gap:6px;}}
.nav-dot{{width:10px;height:10px;border-radius:2px;flex-shrink:0;}}
.nav-sub{{font-size:10px;color:rgba(255,255,255,0.38);margin-top:1px;}}
.header-right{{margin-left:auto;flex-shrink:0;}}
.mobile-nav-toggle{{display:none;align-items:center;justify-content:space-between;gap:8px;height:36px;border:1px solid rgba(255,255,255,.18);border-radius:8px;background:rgba(255,255,255,.08);color:#fff;padding:0 10px;font-family:'Pretendard',sans-serif;font-size:12px;font-weight:700;cursor:pointer;min-width:132px;}}
.mobile-nav-toggle::after{{content:'▾';font-size:11px;color:rgba(255,255,255,.58);}}
.mobile-nav-toggle.open::after{{content:'▴';}}
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
.reorder-list{{max-height:320px;overflow-y:auto;padding-right:4px;}}
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
.target-panel{{margin-top:14px;}}
.target-panel-actions{{display:flex;align-items:center;gap:8px;margin-left:auto;}}
.target-sync{{font-size:11px;color:var(--ink3);}}
.target-editor{{display:none;background:#f8fafc;border:1px solid var(--border);border-radius:8px;padding:14px;margin:0 0 14px;}}
.target-editor.active{{display:block;}}
.target-editor-head{{display:flex;align-items:center;justify-content:space-between;gap:12px;margin-bottom:12px;}}
.target-editor-title{{font-size:14px;font-weight:800;color:var(--ink);}}
.target-editor-grid{{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:10px;}}
.target-editor-field{{display:flex;flex-direction:column;gap:5px;font-size:11px;font-weight:700;color:var(--ink2);}}
.target-editor-field input{{width:100%;height:36px;border:1px solid var(--border2);border-radius:7px;background:#fff;padding:0 10px;font:12px 'DM Mono',monospace;color:var(--ink);}}
.target-editor-foot{{display:flex;justify-content:flex-end;gap:8px;margin-top:12px;}}
.target-rate{{font-weight:800;}}

/* TABLE */
.table-wrap{{border:1px solid var(--border);border-radius:8px;overflow:auto;max-height:500px;}}
table{{border-collapse:collapse;width:100%;min-width:900px;font-size:12.5px;}}
.detail-table{{table-layout:fixed;min-width:1510px;}}
.detail-table th:nth-child(3),.detail-table td.standard-name-cell{{width:var(--standard-name-col-width);}}
.standard-name-cell{{max-width:none;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}}
.resizable-th{{position:relative;padding-right:18px;}}
.col-resizer{{position:absolute;right:0;top:0;width:8px;height:100%;cursor:col-resize;user-select:none;touch-action:none;}}
.col-resizer:hover{{background:rgba(59,130,246,.18);}}
thead{{position:sticky;top:0;z-index:2;}}
th{{background:#f8fafc;color:var(--ink3);font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:0.5px;padding:9px 10px;border-bottom:1px solid var(--border);text-align:left;white-space:nowrap;}}
td{{padding:9px 10px;border-bottom:1px solid var(--border);color:var(--ink2);vertical-align:middle;}}
.detail-table thead{{position:sticky;top:0;z-index:5;box-shadow:0 1px 0 var(--border);}}
.detail-table thead th{{position:relative;z-index:5;}}
.table-filter-row th{{background:#fff;padding:6px 8px;border-bottom:1px solid var(--border);}}
.table-filter-row input,.table-filter-row select{{width:100%;height:28px;border:1px solid var(--border);border-radius:6px;padding:0 7px;font-size:11px;color:var(--ink2);background:#fff;}}
.sortable-th{{cursor:pointer;user-select:none;}}
.sortable-th::after{{content:attr(data-sort-mark);float:right;color:var(--ink3);font-size:10px;margin-left:6px;}}
.sortable-th:hover{{color:var(--ink2);background:#f1f5f9;}}
tr:last-child td{{border-bottom:none;}} tr:hover td{{background:#fafbfd;}}
.num{{text-align:right;font-family:'DM Mono',monospace;font-size:12px;}}
.td-main{{color:var(--ink);font-weight:600;}} .td-mono{{font-family:'DM Mono',monospace;font-size:11.5px;color:var(--ink3);}}

/* BADGE */
.badge{{display:inline-flex;align-items:center;height:22px;padding:0 8px;border-radius:5px;font-size:11px;font-weight:600;white-space:nowrap;}}
.badge-blue{{background:var(--blue-soft);color:var(--blue2);}} .badge-green{{background:var(--teal-soft);color:#059669;}}
.badge-amber{{background:var(--amber-soft);color:#b45309;}} .badge-red{{background:var(--red-soft);color:#dc2626;}}
.badge-gray{{background:#f1f5f9;color:var(--ink3);}} .badge-indigo{{background:#eef2ff;color:#4338ca;}}

.tab-panel{{display:none;}} .tab-panel.active{{display:block;}}
.todo-shell{{display:grid;grid-template-columns:minmax(0,1fr) 280px;gap:14px;align-items:start;}}
.todo-top{{display:flex;justify-content:space-between;gap:12px;align-items:flex-start;margin-bottom:14px;}}
.todo-title{{font-size:20px;font-weight:800;color:var(--ink);}}
.todo-sub{{font-size:12px;color:var(--ink3);margin-top:4px;}}
.todo-view-tabs{{display:flex;gap:6px;background:#e5ebf4;border-radius:8px;padding:4px;}}
.todo-view-tabs button{{border:0;background:transparent;padding:8px 14px;border-radius:7px;font-weight:700;color:var(--ink2);cursor:pointer;}}
.todo-view-tabs button.active{{background:#fff;color:var(--ink);box-shadow:0 1px 4px rgba(15,23,42,.12);}}
.todo-actions{{display:flex;gap:8px;flex-wrap:wrap;justify-content:flex-end;}}
.todo-btn{{border:0;border-radius:8px;padding:9px 13px;font-weight:800;cursor:pointer;background:#eef2ff;color:#1e40af;}}
.todo-btn.primary{{background:#2563eb;color:#fff;}} .todo-btn.dark{{background:#111827;color:#fff;}} .todo-btn.green{{background:#16a34a;color:#fff;}} .todo-btn.subtle{{background:#f1f5f9;color:#475569;}}
.todo-stats{{display:grid;grid-template-columns:repeat(5,1fr);gap:10px;margin-bottom:14px;}}
.todo-stat{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:14px;text-align:center;box-shadow:var(--shadow);}}
.todo-stat b{{display:block;font-size:24px;color:var(--ink);}} .todo-stat span{{font-size:12px;color:var(--ink2);}}
.todo-stat.blue{{border-color:#3b82f6;background:#eff6ff;}} .todo-stat.green b{{color:#10b981;}} .todo-stat.red b{{color:#ef4444;}} .todo-stat.gray b{{color:#94a3b8;}}
.todo-form{{display:none;background:#f8fafc;border:1px solid var(--line);border-radius:8px;padding:14px;margin-bottom:14px;}}
.todo-form.active{{display:block;}}
.todo-form-head{{display:flex;justify-content:space-between;align-items:flex-start;gap:12px;margin-bottom:12px;}}
.todo-form-title{{font-size:15px;font-weight:800;color:var(--ink);}}
.todo-form-sub{{font-size:12px;color:var(--ink3);margin-top:3px;}}
.todo-form-grid{{display:grid;grid-template-columns:1.2fr 1fr 1fr 1fr;gap:10px;}}
.todo-form textarea{{grid-column:1/-1;min-height:78px;resize:vertical;}}
.todo-form input,.todo-form select,.todo-form textarea{{border:1px solid var(--line);border-radius:7px;padding:9px 10px;font:inherit;background:#fff;}}
.todo-form .file-line{{grid-column:1/-1;display:flex;align-items:center;gap:10px;flex-wrap:wrap;color:var(--ink2);font-size:12px;}}
.todo-board{{display:none;}} .todo-board.active{{display:block;}}
.gantt-project{{background:#fff;border:1px solid var(--line);border-radius:8px;overflow:hidden;margin-bottom:12px;box-shadow:var(--shadow);}}
.gantt-head{{display:flex;justify-content:space-between;align-items:center;background:#fbfdff;padding:12px 14px;border-bottom:1px solid var(--line);font-weight:800;}}
.gantt-head-main{{display:flex;align-items:center;gap:10px;min-width:0;}}
.gantt-head-main>span:first-child{{overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}}
.todo-project-delete{{border:1px solid #fecaca;background:#fff;color:#b91c1c;border-radius:6px;padding:6px 8px;font:inherit;font-size:12px;font-weight:800;cursor:pointer;white-space:nowrap;}}
.todo-project-delete:hover{{background:#fef2f2;}} .todo-project-delete:disabled{{cursor:wait;opacity:.6;}}
.gantt-meta{{font-size:12px;color:var(--ink3);font-weight:600;}}
.gantt-row{{display:grid;grid-template-columns:210px 1fr;min-height:48px;border-bottom:1px solid #edf2f7;}}
.gantt-row:last-child{{border-bottom:0;}}
.gantt-task-name{{padding:12px 14px;font-size:13px;display:flex;gap:8px;align-items:center;min-width:0;}}
.gantt-row{{cursor:pointer;}}
.gantt-row:hover .gantt-task-name,.gantt-row:hover .gantt-timeline{{background:#f8fafc;}}
.gantt-dot{{width:9px;height:9px;border-radius:50%;background:#94a3b8;flex:0 0 auto;}}
.gantt-dot.progress{{background:#10b981;}} .gantt-dot.done{{background:#3b82f6;}} .gantt-dot.delay{{background:#ef4444;}}
.gantt-timeline{{position:relative;background:repeating-linear-gradient(90deg,#fff 0,#fff calc(100% / 30 - 1px),#edf2f7 calc(100% / 30));overflow:hidden;}}
.gantt-bar{{position:absolute;top:12px;height:24px;border-radius:6px;background:#34d399;color:#fff;font-size:11px;font-weight:800;display:flex;align-items:center;justify-content:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;padding:0 8px;min-width:42px;}}
.gantt-bar.done{{background:#60a5fa;}} .gantt-bar.delay{{background:#f97316;}} .gantt-bar.todo{{background:#cbd5e1;color:#334155;}}
.gantt-today{{position:absolute;top:0;bottom:0;width:2px;background:#ef4444;left:50%;}}
.calendar-grid{{display:grid;grid-template-columns:repeat(7,1fr);gap:8px;}}
.calendar-day{{min-height:96px;background:#fff;border:1px solid var(--line);border-radius:8px;padding:8px;}}
.calendar-date{{font-size:12px;font-weight:800;color:var(--ink2);margin-bottom:6px;}}
.calendar-chip{{font-size:11px;border-radius:6px;padding:5px 6px;margin-top:4px;background:#ecfdf5;color:#047857;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;}}
.calendar-chip{{cursor:pointer;}}
.calendar-chip:hover{{filter:brightness(.96);}}
.calendar-chip.delay{{background:#fef2f2;color:#dc2626;}} .calendar-chip.done{{background:#eff6ff;color:#2563eb;}}
.todo-list{{width:100%;border-collapse:collapse;}}
.todo-list tbody tr{{cursor:pointer;}}
.todo-list tbody tr:hover td{{background:#f8fafc;}}
.todo-row-actions{{display:flex;gap:6px;justify-content:flex-end;}}
.todo-row-actions button{{border:0;background:#eef2ff;color:#1e40af;border-radius:6px;padding:6px 8px;cursor:pointer;font-weight:700;}}
.todo-side{{position:sticky;top:86px;}}
.todo-deadline{{background:#fff;border:1px solid var(--line);border-radius:8px;padding:14px;box-shadow:var(--shadow);}}
.todo-deadline h3{{margin:0 0 10px;font-size:14px;color:#ef4444;}}
.todo-deadline-item{{border-bottom:1px solid #edf2f7;padding:9px 0;font-size:12px;}}
.todo-deadline-item:last-child{{border-bottom:0;}} .todo-deadline-item b{{display:block;color:var(--ink);margin-bottom:4px;}} .todo-deadline-item span{{color:#f97316;font-weight:800;}}
@media (max-width:980px){{.todo-shell{{grid-template-columns:1fr;}}.todo-stats{{grid-template-columns:repeat(2,1fr);}}.todo-form-grid{{grid-template-columns:1fr;}}.gantt-row{{grid-template-columns:1fr;}}.gantt-timeline{{height:50px;}}.todo-side{{position:static;}}}}
.foot{{font-size:11px;color:var(--ink3);margin-top:10px;line-height:1.7;}}
::-webkit-scrollbar{{width:5px;height:5px;}} ::-webkit-scrollbar-track{{background:transparent;}}
::-webkit-scrollbar-thumb{{background:var(--border2);border-radius:3px;}}

@media(max-width:1200px){{.kpis{{grid-template-columns:repeat(3,1fr);}}.grid2,.grid2-eq,.charts-row{{grid-template-columns:1fr;}}}}
@media(max-width:680px){{
  header{{overflow:visible;}}
  .header-top{{height:auto;min-height:52px;padding:8px 12px;gap:8px;flex-wrap:wrap;position:relative;}}
  main{{padding-left:12px;padding-right:12px;}}
  .logo-area{{min-width:auto;width:auto;max-width:128px;padding-right:10px;margin-right:0;}}
  .logo-text{{font-size:12px;}}
  .logo-sub{{font-size:9px;}}
  .mobile-nav-toggle{{display:flex;flex:1;min-width:128px;}}
  .header-right{{margin-left:0;}}
  .status-chip{{height:30px;padding:3px 8px;font-size:10px;}}
  .header-nav{{display:none;order:5;width:100%;height:auto;max-height:calc(100vh - 76px);overflow-y:auto;background:#111827;border:1px solid rgba(255,255,255,.1);border-radius:10px;padding:6px;box-shadow:0 10px 24px rgba(0,0,0,.28);}}
  .header-nav.open{{display:grid;grid-template-columns:1fr;gap:4px;}}
  .nav-item{{height:auto;min-height:48px;border-bottom:0;border-left:3px solid transparent;border-radius:8px;padding:9px 10px;text-decoration:none;}}
  .nav-item.active{{border-left-color:var(--blue);border-bottom-color:transparent;background:rgba(59,130,246,.12);}}
  .nav-label{{font-size:13px;}}
  .nav-sub{{font-size:10px;}}
  .kpis{{grid-template-columns:1fr 1fr;}}
  .target-editor-grid{{grid-template-columns:1fr;}}
  .target-panel .panel-hd{{align-items:flex-start;gap:8px;}}
  .target-panel-actions{{width:100%;margin-left:0;justify-content:space-between;}}
}}
</style>
</head>
<body>
<header>
  <div class="header-top">
    <div class="logo-area">
      <div class="logo-icon">PD</div>
      <div><div class="logo-text">상품DATA</div><div class="logo-sub">운영 대시보드</div></div>
    </div>
    <button type="button" class="mobile-nav-toggle" id="mobileNavToggle" aria-expanded="false" aria-controls="mainNav">요약</button>
    <nav class="header-nav" id="mainNav">
      <a class="nav-item active" href="#overview" data-tab="overview">
        <div class="nav-label"><span class="nav-dot" style="background:#3b82f6"></span>요약</div>
        <div class="nav-sub">KPI · 최근 트렌드</div>
      </a>
      <a class="nav-item" href="#todo" data-tab="todo">
        <div class="nav-label"><span class="nav-dot" style="background:#8b5cf6"></span>To do list</div>
        <div class="nav-sub">프로젝트 일정</div>
      </a>
      <a class="nav-item" href="#compare" data-tab="compare">
        <div class="nav-label"><span class="nav-dot" style="background:#06b6d4"></span>매출 비교</div>
        <div class="nav-sub">연도 · 월별 대비</div>
      </a>
      <a class="nav-item" href="#calendar" data-tab="calendar">
        <div class="nav-label"><span class="nav-dot" style="background:#6366f1"></span>일자별 매출</div>
        <div class="nav-sub">날짜별 트래킹</div>
      </a>
      <a class="nav-item" href="#retailer" data-tab="retailer">
        <div class="nav-label"><span class="nav-dot" style="background:#10b981"></span>유통사별 매출</div>
        <div class="nav-sub">채널별 현황</div>
      </a>
      <a class="nav-item" href="#detail" data-tab="detail">
        <div class="nav-label"><span class="nav-dot" style="background:#f59e0b"></span>상품별 매출</div>
        <div class="nav-sub">SKU 상세 분석</div>
      </a>
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

  <!-- TODO TAB -->
  <div id="todo" class="tab-panel">
    <div class="todo-top">
      <div>
        <div class="todo-title">프로젝트 스케줄</div>
        <div class="todo-sub" id="todoToday">오늘 기준 프로젝트 일정</div>
      </div>
      <div class="todo-actions">
        <div class="todo-view-tabs">
          <button class="active" data-todo-view="gantt">간트</button>
          <button data-todo-view="calendar">캘린더</button>
          <button data-todo-view="list">리스트</button>
        </div>
        <button class="todo-btn primary" id="todoAddProject">+ 프로젝트</button>
        <button class="todo-btn primary" id="todoAddTask">+ 태스크</button>
        <button class="todo-btn dark" id="todoSave">저장</button>
        <button class="todo-btn green" id="todoCsv">CSV</button>
        <button class="todo-btn subtle" id="todoImportBtn">가져오기</button>
        <input type="file" id="todoImportFile" accept=".csv,.json" style="display:none">
      </div>
    </div>

    <div class="todo-stats">
      <div class="todo-stat blue"><b id="todoStatAll">0</b><span>전체</span></div>
      <div class="todo-stat green"><b id="todoStatProgress">0</b><span>진행중</span></div>
      <div class="todo-stat"><b id="todoStatDone">0</b><span>완료</span></div>
      <div class="todo-stat red"><b id="todoStatDelay">0</b><span>지연</span></div>
      <div class="todo-stat gray"><b id="todoStatTodo">0</b><span>할 일</span></div>
    </div>

    <div class="todo-form" id="todoForm">
      <div class="todo-form-head">
        <div>
          <div class="todo-form-title" id="todoFormTitle">프로젝트 등록</div>
          <div class="todo-form-sub" id="todoFormSub">프로젝트 또는 태스크를 등록합니다.</div>
        </div>
      </div>
      <div class="todo-form-grid">
        <input id="todoProject" placeholder="프로젝트명">
        <input id="todoTask" placeholder="태스크명">
        <input id="todoOwner" placeholder="담당자">
        <select id="todoStatus">
          <option value="todo">할 일</option>
          <option value="progress">진행중</option>
          <option value="done">완료</option>
          <option value="delay">지연</option>
        </select>
        <input id="todoStart" type="date" title="시작 일자">
        <input id="todoEnd" type="date" title="완료 예정일">
        <input id="todoDoneDate" type="date" title="완료 일자">
        <input id="todoFile" type="file">
        <textarea id="todoMemo" placeholder="내용을 입력하세요"></textarea>
        <div class="file-line">
          <span id="todoFileName">첨부 파일 없음</span>
          <button class="todo-btn primary" id="todoSubmit" type="button">등록</button>
          <button class="todo-btn subtle" id="todoCancel" type="button">취소</button>
        </div>
      </div>
    </div>

    <div class="todo-shell">
      <div>
        <div id="todoGantt" class="todo-board active"></div>
        <div id="todoCalendar" class="todo-board"><div class="calendar-grid" id="todoCalendarGrid"></div></div>
        <div id="todoList" class="todo-board">
          <div class="table-wrap">
            <table class="todo-list">
              <thead><tr><th>프로젝트</th><th>태스크</th><th>담당자</th><th>상태</th><th>시작</th><th>완료 예정</th><th>완료</th><th>첨부</th><th>내용</th><th></th></tr></thead>
              <tbody id="todoRows"></tbody>
            </table>
          </div>
        </div>
      </div>
      <aside class="todo-side">
        <div class="todo-deadline">
          <h3>마감 임박</h3>
          <div id="todoDeadline"></div>
        </div>
      </aside>
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

    <div class="panel target-panel">
      <div class="panel-hd">
        <div><span class="panel-title">월별 목표 매출 달성</span><span class="panel-meta">몰별 목표 · 실적 · 달성률</span></div>
        <div class="target-panel-actions">
          <span class="target-sync" id="targetSyncStatus">엑셀 원본 목표</span>
          <button class="btn-sm" id="targetEditBtn" type="button">목표 수정</button>
        </div>
      </div>
      <div class="toolbar">
        <select id="targetYear"></select>
        <select id="targetMonth">
          <option value="">전체 월</option>
          <option value="01">1월</option><option value="02">2월</option><option value="03">3월</option>
          <option value="04">4월</option><option value="05">5월</option><option value="06">6월</option>
          <option value="07">7월</option><option value="08">8월</option><option value="09">9월</option>
          <option value="10">10월</option><option value="11">11월</option><option value="12">12월</option>
        </select>
      </div>
      <div class="target-editor" id="targetEditor">
        <div class="target-editor-head">
          <div><div class="target-editor-title" id="targetEditorTitle">월 목표 수정</div><div class="panel-meta">수정값은 모든 대시보드 사용자에게 동일하게 적용됩니다.</div></div>
        </div>
        <div class="target-editor-grid" id="targetEditorGrid"></div>
        <div class="target-editor-foot">
          <button class="btn-sm" id="targetEditCancel" type="button">취소</button>
          <button class="todo-btn primary" id="targetEditSave" type="button">저장</button>
        </div>
      </div>
      <div class="sum-strip">
        <div class="sum-card"><div class="sum-label" id="targetGoalLabel">목표 매출</div><div class="sum-value" id="targetGoalSales">0</div></div>
        <div class="sum-card"><div class="sum-label" id="targetActualLabel">실매출</div><div class="sum-value" id="targetActualSales">0</div></div>
        <div class="sum-card"><div class="sum-label">목표 달성률</div><div class="sum-value" id="targetAchievement">-</div></div>
      </div>
      <div class="chart-box" style="height:280px;margin-bottom:14px"><canvas id="targetChart"></canvas></div>
      <div class="table-wrap" style="max-height:390px">
        <table>
          <thead><tr><th id="targetScopeHead">월</th><th class="num">목표 매출</th><th class="num">실매출</th><th class="num">차이</th><th class="num">달성률</th></tr></thead>
          <tbody id="targetRows"></tbody>
        </table>
      </div>
      <div class="foot">목표는 전달받은 2026 일일판매정리 첫 시트 기준입니다. 기타몰 실적은 별도 목표가 없는 유통사의 합산입니다.</div>
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
        <button class="sub-tab" data-retailer="ABC마트">ABC마트</button>
        <button class="sub-tab" data-retailer="러너블">러너블</button>
        <button class="sub-tab" data-retailer="기타">기타</button>
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
          <option>ABC마트</option><option>러너블</option><option>기타</option>
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
        <div class="rank-list reorder-list" id="reorderAlerts"></div>
      </div>
      <div class="table-wrap">
        <table class="detail-table">
          <thead>
            <tr>
              <th class="sortable-th" data-detail-sort="alert" style="width:72px">알림</th><th class="sortable-th" data-detail-sort="sku" style="width:120px">바코드</th><th class="sortable-th resizable-th" data-detail-sort="name" id="standardNameHeader">표준상품명<span class="col-resizer" id="standardNameResizer"></span></th>
              <th class="sortable-th" data-detail-sort="year">연도</th><th class="sortable-th" data-detail-sort="season">시즌</th><th class="sortable-th" data-detail-sort="category">복종</th>
              <th class="sortable-th num" data-detail-sort="qty">판매수량</th>
              <th class="sortable-th num" data-detail-sort="weeklyQty">주간 판매수량</th>
              <th class="sortable-th num" data-detail-sort="weeklyRate">주간 판매율</th>
              <th class="sortable-th num" data-detail-sort="stockQty">현재고</th>
              <th class="sortable-th num" data-detail-sort="stockRate">전체 재고 소진율</th>
              <th class="sortable-th num" data-detail-sort="gross">정상가금액</th>
              <th class="sortable-th num" data-detail-sort="payment">실판매금액</th>
              <th class="sortable-th num" data-detail-sort="avgUnit">평균단가</th>
              <th class="sortable-th num" data-detail-sort="discount">할인율</th>
            </tr>
            <tr class="table-filter-row" id="detailFilterRows"></tr>
          </thead>
          <tbody id="detailRows"></tbody>
        </table>
      </div>
    </div>
  </div>
</main>

<script>
const rawRows = {rows_json};
const RETAILERS = {json.dumps(RETAILERS, **JSON_KWARGS)};
const PREV_YEAR_MONTHLY = {prev_year_monthly_json};
const HISTORICAL_MONTHLY = {historical_monthly_json};
const HISTORICAL_DAILY = {historical_daily_json};
const HISTORICAL_DAILY_MONTHLY = {historical_daily_monthly_json};
const MANUAL_SALES = {manual_sales_json};
const SALES_TARGET_DEFAULTS = {sales_target_defaults_json};
const SALES_TARGET_GROUPS = ['자사몰','무신사','글로리어스워커','4XR','29cm','기타몰'];
const SALES_TARGET_DIRECT_RETAILERS = new Set(SALES_TARGET_GROUPS.filter(name=>name!=='기타몰'));
const SALES_TARGET_LOCAL_KEY = 'plaknitSalesTargets.v1';
const BARCODE_YEAR_LABELS = {{A:'2022',B:'2023',C:'2024',D:'2025',E:'2026',F:'2027',G:'2028'}};
const REORDER_ALLOWED_YEAR_CODES = new Set(['D','E']); // 바코드 연도 D/E(2025/2026) 상품만 리오더 알림 대상

const fmt  = n => Math.round(n).toLocaleString('ko-KR');
const pct  = n => n.toFixed(1) + '%';
const fmtD = s => s.slice(5); // "2026-01-03" → "01-03"
const num = v => Number.isFinite(Number(v)) ? Number(v) : 0;
const rowGross = r => num(r.gross ?? r.gross_sales);
const rowPayment = r => num(r.payment ?? r.payment_sales);
const rowQty = r => num(r.qty);
const rowOrders = r => num(r.orders);
function uniqSorted(values) {{
  return Array.from(new Set(values.filter(v => v !== undefined && v !== null && String(v).trim() !== '').map(v => String(v).trim()))).sort((a,b)=>a.localeCompare(b));
}}
function validDiscount(gross, payment) {{
  gross = Number(gross || 0);
  payment = Number(payment || 0);
  if (gross <= 0 || payment < 0 || payment > gross) return null;
  return (1 - payment / gross) * 100;
}}
function barcodeYearCode(value) {{
  const match = String(value || '').trim().toUpperCase().match(/^P([A-Z])/);
  return match ? match[1] : '';
}}
function barcodeYearLabel(code) {{
  return code ? (BARCODE_YEAR_LABELS[code] || code) : '';
}}
function rowBarcodeYearCode(row) {{
  const sku = String(row.match_sku || row.stock_barcode || '').trim().toUpperCase();
  return barcodeYearCode(sku);
}}
function discountText(gross, payment) {{
  const d = validDiscount(gross, payment);
  return d === null ? '-' : pct(d);
}}
function unitPriceBase(row) {{
  let qty = 0;
  let payment = 0;
  const daily = Array.isArray(row.daily) ? row.daily : [];
  if (daily.length) {{
    daily.forEach(d => {{
      const q = Number(d.qty || 0);
      const p = Number(d.payment || 0);
      if (q > 0 && p > 0) {{
        qty += q;
        payment += p;
      }}
    }});
  }} else {{
    const q = Number(row.qty || 0);
    const p = Number(row.payment || 0);
    if (q > 0 && p > 0) {{
      qty += q;
      payment += p;
    }}
  }}
  return {{qty, payment}};
}}
const DETAIL_SIZE_TOKEN = '(4XL|3XL|2XL|XXXL|XXL|XL|XS|FREE|F|S|M|L)';
function cleanDetailText(value) {{
  return String(value || '').replace(/^\\[[^\\]]+\\]\\s*/, '').replace(/\\s+/g, ' ').trim();
}}
function normalizeDetailSize(size) {{
  size = String(size || '').toUpperCase();
  return size === 'F' ? 'FREE' : size;
}}
function detailSizeFromValue(value) {{
  const text = cleanDetailText(value).replace(/[()]/g, ' ').toUpperCase();
  const re = new RegExp(`(?:^|[^A-Z0-9])${{DETAIL_SIZE_TOKEN}}(?=$|[^A-Z0-9])`, 'g');
  const matches = Array.from(text.matchAll(re));
  if (!matches.length) return '';
  return normalizeDetailSize(matches[matches.length - 1][1]);
}}
function detailSizeFromSku(value) {{
  const text = String(value || '').trim().toUpperCase();
  const match = text.match(new RegExp(`${{DETAIL_SIZE_TOKEN}}$`));
  return match ? normalizeDetailSize(match[1]) : '';
}}
function detailSize(row) {{
  return detailSizeFromValue(row.size) ||
    detailSizeFromValue(row.standard_name) ||
    detailSizeFromValue(row.stock_name) ||
    detailSizeFromSku(row.match_sku);
}}
function detailProductNameFromText(value) {{
  let name = cleanDetailText(value);
  if (!name) return '';
  name = name
    .replace(/[_-](4XL|3XL|2XL|XXXL|XXL|XL|XS|FREE|F|S|M|L)$/i, '')
    .replace(/\\s+(4XL|3XL|2XL|XXXL|XXL|XL|XS|FREE|F|S|M|L)$/i, '')
    .replace(/[_-](BLACK|WHITE|NAVY|CHARCOAL|GREY|GRAY|M\\/GREY|L\\/KHAKI|블랙|화이트|네이비|차콜|그레이|라이트그레이|멜란지그레이|베이지|카키|브라운|레드|블루|민트블루|세이지|올리브|핑크|라임)$/i, '')
    .replace(/\\s+/g, ' ')
    .trim();
  return name || '';
}}
function detailColor(row) {{
  const explicit = cleanDetailText(row.color);
  if (explicit) return explicit;
  const code = String(row.match_sku || row.stock_barcode || '').toUpperCase();
  if (code.includes('WHT')) return '화이트';
  if (code.includes('BLK')) return '블랙';
  if (code.includes('NVY') || code.includes('NAVY')) return '네이비';
  if (code.includes('GRY') || code.includes('GREY') || code.includes('GRAY')) return '그레이';
  return '';
}}
function detailApplyColor(product, color) {{
  if (color && product && product !== '-' && !detailGroupKeyText(product).includes(detailGroupKeyText(color))) {{
    return `${{product}} ${{color}}`;
  }}
  return product;
}}
function detailProductName(row) {{
  let product = detailProductNameFromText(row.name) ||
    detailProductNameFromText(row.standard_name) ||
    detailProductNameFromText(row.stock_name) ||
    '-';
  return detailApplyColor(product, detailColor(row));
}}
function detailGroupKeyText(value) {{
  return cleanDetailText(value).replace(/\\s+/g, '').toUpperCase();
}}
function detailGroupInfo(row) {{
  const product = detailProductName(row);
  const size = detailSize(row);
  return {{
    key: `${{detailGroupKeyText(product)}}|${{size}}`,
    label: size ? `${{product}} ${{size}}` : product,
    product,
    size
  }};
}}
function detailStockKey(row, groupKey) {{
  return cleanDetailText(row.stock_barcode) ||
    cleanDetailText(row.stock_name) ||
    cleanDetailText(row.match_sku) ||
    `${{groupKey}}|${{cleanDetailText(row.standard_name || row.name)}}|${{cleanDetailText(row.color)}}|${{cleanDetailText(row.size)}}|${{Number(row.stock_qty || 0)}}`;
}}
function detailStockMatchesGroup(row, groupInfo) {{
  const groupProductKey = detailGroupKeyText(groupInfo.product);
  const groupSize = groupInfo.size;
  const displayProductKey = detailGroupKeyText(detailProductName(row));
  const displaySize = detailSize(row);
  const stockLabel = row.stock_name || row.standard_name || row.name || '';
  const stockProductKey = detailGroupKeyText(detailApplyColor(detailProductNameFromText(stockLabel), detailColor(row)));
  const stockSize = detailSizeFromValue(stockLabel);
  if (row.match_sku) {{
    return stockProductKey === groupProductKey && (!stockSize || !groupSize || stockSize === groupSize);
  }}
  return displayProductKey === groupProductKey && (!displaySize || !groupSize || displaySize === groupSize);
}}
function addDetailSearchPart(group, value) {{
  const text = cleanDetailText(value);
  if (text) group.searchParts.add(text);
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
const avgUnitBase  = rawRows.reduce((a,r)=>{{
  const b = unitPriceBase(r);
  a.qty += b.qty;
  a.payment += b.payment;
  return a;
}}, {{qty:0,payment:0}});
const avgUnit      = avgUnitBase.qty ? Math.round(avgUnitBase.payment/avgUnitBase.qty) : 0;
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
  const gross = rowGross(r);
  const payment = rowPayment(r);
  retailerMap[r.retailer].payment += payment;
  retailerMap[r.retailer].gross   += gross;
  retailerMap[r.retailer].orders  += rowOrders(r);
  retailerMap[r.retailer].qty     += rowQty(r);
  (r.daily || []).forEach(d => {{
    const key = d.date + '|' + (r.retailer || '');
    if (!retailerExisting[key]) retailerExisting[key] = {{payment:0,gross:0,orders:0,qty:0}};
    retailerExisting[key].payment += Number(d.payment || 0);
    retailerExisting[key].gross += Number(d.gross || 0);
    retailerExisting[key].orders += Number(d.orders || 0);
    retailerExisting[key].qty += Number(d.qty || 0);
  }});
  if (validDiscount(gross, payment) !== null) {{
    retailerMap[r.retailer].validPayment += payment;
    retailerMap[r.retailer].validGross += gross;
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

  const retailerPalette = ['#3b82f6','#10b981','#f59e0b','#8b5cf6','#ef4444','#06b6d4','#64748b','#db2777','#84cc16'];
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

// 월별 목표 매출 --------------------------------------------------------
let salesTargets = loadSalesTargetsLocal();
let targetChartInst = null;

function mergeSalesTargetRows(base, rows) {{
  (rows || []).forEach(row=>{{
    const year = String(row.year || '');
    const month = String(row.month || '').padStart(2,'0');
    const retailer = String(row.retailer || '');
    if (!year || !month || !SALES_TARGET_GROUPS.includes(retailer)) return;
    if (!base[year]) base[year] = {{}};
    if (!base[year][month]) base[year][month] = {{}};
    base[year][month][retailer] = num(row.target_amount);
  }});
  return base;
}}

function loadSalesTargetsLocal() {{
  const base = JSON.parse(JSON.stringify(SALES_TARGET_DEFAULTS));
  try {{
    const saved = JSON.parse(localStorage.getItem(SALES_TARGET_LOCAL_KEY) || '[]');
    if (Array.isArray(saved)) mergeSalesTargetRows(base, saved);
  }} catch(e) {{}}
  return base;
}}

function flattenSalesTargets() {{
  const rows = [];
  Object.entries(salesTargets).forEach(([year, months])=>{{
    Object.entries(months || {{}}).forEach(([month, retailers])=>{{
      SALES_TARGET_GROUPS.forEach(retailer=>rows.push({{
        year:Number(year), month:Number(month), retailer,
        target_amount:num(retailers?.[retailer])
      }}));
    }});
  }});
  return rows;
}}

function saveSalesTargetsLocal() {{
  localStorage.setItem(SALES_TARGET_LOCAL_KEY, JSON.stringify(flattenSalesTargets()));
}}

function targetValue(year, month, retailer) {{
  return num(salesTargets?.[String(year)]?.[String(month).padStart(2,'0')]?.[retailer]);
}}

function targetDailyData(year, retailer) {{
  if (String(year) === '2026') return buildDailyMap(retailer);
  return buildHistoricalDailyMap(String(year), retailer);
}}

function targetActualValue(year, month, retailer) {{
  const monthPrefix = `${{year}}-${{String(month).padStart(2,'0')}}`;
  const sumDaily = rows => rows.filter(row=>row.date.startsWith(monthPrefix)).reduce((sum,row)=>sum+num(row.payment),0);
  if (retailer !== '기타몰') return sumDaily(targetDailyData(year, retailer));
  const all = sumDaily(targetDailyData(year, ''));
  const direct = [...SALES_TARGET_DIRECT_RETAILERS].reduce((sum,name)=>sum+sumDaily(targetDailyData(year, name)),0);
  return all - direct;
}}

function targetRate(actual, goal) {{
  if (goal > 0) return actual / goal * 100;
  return actual === 0 ? 0 : null;
}}

function targetRowsForSelection(year, selectedMonth) {{
  if (selectedMonth) return SALES_TARGET_GROUPS.map(retailer=>{{
    const goal = targetValue(year, selectedMonth, retailer);
    const actual = targetActualValue(year, selectedMonth, retailer);
    return {{label:retailer, goal, actual, diff:actual-goal, rate:targetRate(actual,goal)}};
  }});
  return Array.from({{length:12}},(_,idx)=>String(idx+1).padStart(2,'0')).map(month=>{{
    const goal = SALES_TARGET_GROUPS.reduce((sum,retailer)=>sum+targetValue(year,month,retailer),0);
    const actual = SALES_TARGET_GROUPS.reduce((sum,retailer)=>sum+targetActualValue(year,month,retailer),0);
    return {{label:`${{Number(month)}}월`, goal, actual, diff:actual-goal, rate:targetRate(actual,goal)}};
  }});
}}

function renderSalesTargets() {{
  const year = document.getElementById('targetYear').value || '2026';
  const selectedMonth = document.getElementById('targetMonth').value;
  const rows = targetRowsForSelection(year, selectedMonth);
  const goal = rows.reduce((sum,row)=>sum+row.goal,0);
  const actual = rows.reduce((sum,row)=>sum+row.actual,0);
  const achievement = targetRate(actual,goal);
  const scope = selectedMonth ? `${{Number(selectedMonth)}}월` : '전체 월';
  document.getElementById('targetGoalLabel').textContent = `${{year}}년 ${{scope}} 목표 매출`;
  document.getElementById('targetActualLabel').textContent = `${{year}}년 ${{scope}} 실매출`;
  document.getElementById('targetGoalSales').textContent = fmt(goal);
  document.getElementById('targetActualSales').textContent = fmt(actual);
  document.getElementById('targetAchievement').textContent = achievement === null ? '-' : pct(achievement);
  document.getElementById('targetAchievement').style.color = achievement !== null && achievement >= 100 ? 'var(--teal)' : 'var(--amber)';
  document.getElementById('targetScopeHead').textContent = selectedMonth ? '유통사' : '월';

  if (targetChartInst) targetChartInst.destroy();
  targetChartInst = new Chart(document.getElementById('targetChart'), {{
    type:'bar',
    data:{{labels:rows.map(row=>row.label),datasets:[
      {{label:'목표 매출',data:rows.map(row=>row.goal),backgroundColor:'rgba(100,116,139,.16)',borderColor:'#64748b',borderWidth:1.5,borderRadius:5,yAxisID:'y'}},
      {{label:'실매출',data:rows.map(row=>row.actual),backgroundColor:'rgba(59,130,246,.20)',borderColor:'#3b82f6',borderWidth:1.5,borderRadius:5,yAxisID:'y'}},
      {{type:'line',label:'달성률',data:rows.map(row=>row.rate),borderColor:'#10b981',backgroundColor:'#10b981',borderWidth:2,pointRadius:3,tension:.25,yAxisID:'y1'}}
    ]}},
    options:{{responsive:true,maintainAspectRatio:false,plugins:{{legend:{{display:true,position:'bottom'}},tooltip:{{...ttip,callbacks:{{label:ctx=>ctx.dataset.yAxisID==='y1' ? ` ${{ctx.dataset.label}}: ${{ctx.raw===null?'-':pct(ctx.raw)}}` : ` ${{ctx.dataset.label}}: ${{fmt(ctx.raw)}}`}}}}}},scales:{{x:scl.x,y:scl.y,y1:{{position:'right',grid:{{drawOnChartArea:false}},ticks:{{callback:value=>value+'%'}}}}}}}}
  }});

  document.getElementById('targetRows').innerHTML = rows.map(row=>`<tr>
    <td class="td-main">${{row.label}}</td>
    <td class="num">${{fmt(row.goal)}}</td>
    <td class="num">${{fmt(row.actual)}}</td>
    <td class="num" style="color:${{row.diff>=0?'var(--teal)':'var(--red)'}}">${{fmt(row.diff)}}</td>
    <td class="num target-rate" style="color:${{row.rate!==null&&row.rate>=100?'var(--teal)':'var(--amber)'}}">${{row.rate===null?'-':pct(row.rate)}}</td>
  </tr>`).join('');
}}

function initSalesTargetControls() {{
  const years = Array.from(new Set([...compareYears,...Object.keys(salesTargets)])).sort();
  const yearSelect = document.getElementById('targetYear');
  yearSelect.innerHTML = years.map(year=>`<option value="${{year}}">${{year}}년</option>`).join('');
  yearSelect.value = years.includes('2026') ? '2026' : years[years.length-1];
}}

function openSalesTargetEditor() {{
  const year = document.getElementById('targetYear').value || '2026';
  let month = document.getElementById('targetMonth').value;
  if (!month) {{
    const now = new Date();
    month = String(String(now.getFullYear()) === year ? now.getMonth()+1 : 1).padStart(2,'0');
    document.getElementById('targetMonth').value = month;
    renderSalesTargets();
  }}
  document.getElementById('targetEditorTitle').textContent = `${{year}}년 ${{Number(month)}}월 목표 수정`;
  document.getElementById('targetEditor').dataset.year = year;
  document.getElementById('targetEditor').dataset.month = month;
  document.getElementById('targetEditorGrid').innerHTML = SALES_TARGET_GROUPS.map(retailer=>`
    <label class="target-editor-field">${{retailer}}
      <input type="number" min="0" step="1" data-target-retailer="${{retailer}}" value="${{targetValue(year,month,retailer)}}">
    </label>`).join('');
  document.getElementById('targetEditor').classList.add('active');
}}

function closeSalesTargetEditor() {{
  document.getElementById('targetEditor').classList.remove('active');
}}

async function saveSalesTargetRemote(rows) {{
  if (!(await ensureTodoSupabaseConfig())) throw new Error('Supabase config is missing.');
  const res = await fetch(`${{TODO_SUPABASE_REST}}/sales_targets?on_conflict=year,month,retailer`, {{
    method:'POST',
    headers:{{...TODO_HEADERS,Prefer:'resolution=merge-duplicates,return=representation'}},
    body:JSON.stringify(rows)
  }});
  if (!res.ok) throw new Error(await res.text());
  return res.json();
}}

async function submitSalesTargetEditor() {{
  const editor = document.getElementById('targetEditor');
  const year = editor.dataset.year;
  const month = editor.dataset.month;
  const rows = Array.from(editor.querySelectorAll('[data-target-retailer]')).map(input=>({{
    year:Number(year),month:Number(month),retailer:input.dataset.targetRetailer,target_amount:Math.max(0,num(input.value))
  }}));
  if (!salesTargets[year]) salesTargets[year] = {{}};
  if (!salesTargets[year][month]) salesTargets[year][month] = {{}};
  rows.forEach(row=>salesTargets[year][month][row.retailer]=row.target_amount);
  saveSalesTargetsLocal();
  const button = document.getElementById('targetEditSave');
  button.disabled = true;
  document.getElementById('targetSyncStatus').textContent = '저장 중';
  try {{
    await saveSalesTargetRemote(rows);
    document.getElementById('targetSyncStatus').textContent = 'Supabase 동기화 완료';
  }} catch(e) {{
    document.getElementById('targetSyncStatus').textContent = '이 브라우저에 저장됨';
    console.warn('Supabase sales target save failed. Local fallback is active.', e);
  }} finally {{
    button.disabled = false;
  }}
  closeSalesTargetEditor();
  renderSalesTargets();
}}

async function refreshSalesTargetsRemote() {{
  try {{
    if (!(await ensureTodoSupabaseConfig())) throw new Error('Supabase config is missing.');
    const res = await fetch(`${{TODO_SUPABASE_REST}}/sales_targets?select=year,month,retailer,target_amount&order=year.asc,month.asc`, {{headers:TODO_HEADERS}});
    if (!res.ok) throw new Error(await res.text());
    mergeSalesTargetRows(salesTargets, await res.json());
    saveSalesTargetsLocal();
    document.getElementById('targetSyncStatus').textContent = 'Supabase 동기화';
    renderSalesTargets();
  }} catch(e) {{
    document.getElementById('targetSyncStatus').textContent = '엑셀 원본 목표';
    console.warn('Supabase sales target load failed. Embedded defaults are active.', e);
  }}
}}

initSalesTargetControls();
renderSalesTargets();
document.getElementById('targetYear').addEventListener('change',renderSalesTargets);
document.getElementById('targetMonth').addEventListener('change',()=>{{closeSalesTargetEditor();renderSalesTargets();}});
document.getElementById('targetEditBtn').addEventListener('click',openSalesTargetEditor);
document.getElementById('targetEditCancel').addEventListener('click',closeSalesTargetEditor);
document.getElementById('targetEditSave').addEventListener('click',submitSalesTargetEditor);

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

  const tableRows = [...f].sort((a,b)=>b.date.localeCompare(a.date));
  document.getElementById('dailyRows').innerHTML = !tableRows.length
    ? '<tr><td colspan="6" style="text-align:center;color:var(--ink3);padding:24px">데이터 없음</td></tr>'
    : tableRows.map(d=>`<tr>
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

let detailTableSort = {{key:'payment', dir:'desc'}};
function detailAlertText(row) {{
  return !row.stock_known ? '재고확인' : row.reorder_reasons.length ? '확인' : '정상';
}}
function detailDiscountValue(row) {{
  return row.validGross > 0 && row.validPayment <= row.validGross ? (1 - row.validPayment / row.validGross) * 100 : -1;
}}
function detailSortValue(row, key) {{
  if (key === 'alert') return detailAlertText(row);
  if (key === 'sku') return row.match_sku || '';
  if (key === 'name') return row.standard_name || '';
  if (key === 'year') return (row.years || []).join(', ');
  if (key === 'season') return (row.seasons || []).join(', ');
  if (key === 'category') return row.categoryLarge.concat(row.categorySmall).filter(v=>v && v !== '-').join(' / ');
  if (key === 'qty') return Number(row.qty || 0);
  if (key === 'weeklyQty') return Number(row.weekly_qty || 0);
  if (key === 'weeklyRate') return Number(row.weekly_rate || 0);
  if (key === 'stockQty') return Number(row.stock_qty || 0);
  if (key === 'stockRate') return Number(row.stock_rate || 0);
  if (key === 'gross') return Number(row.gross || 0);
  if (key === 'payment') return Number(row.payment || 0);
  if (key === 'avgUnit') return Number(row.avg_unit || 0);
  if (key === 'discount') return detailDiscountValue(row);
  return '';
}}
function applyDetailSort(rows) {{
  const key = detailTableSort.key || 'payment';
  const dir = detailTableSort.dir === 'asc' ? 1 : -1;
  rows.sort((a,b)=>{{
    const av = detailSortValue(a, key);
    const bv = detailSortValue(b, key);
    if (typeof av === 'number' || typeof bv === 'number') return ((Number(av)||0) - (Number(bv)||0)) * dir;
    return String(av || '').localeCompare(String(bv || '')) * dir;
  }});
}}
function updateDetailSortMarks() {{
  document.querySelectorAll('[data-detail-sort]').forEach(th=>{{
    th.dataset.sortMark = th.dataset.detailSort === detailTableSort.key ? (detailTableSort.dir === 'asc' ? '▲' : '▼') : '↕';
  }});
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
    const groupInfo = detailGroupInfo(r);
    const key = groupInfo.key;
    if (!grouped[key]) grouped[key] = {{
      key,
      match_sku:r.match_sku || '',
      standard_name:groupInfo.label,
      product_name:groupInfo.product,
      detail_size:groupInfo.size,
      matchSkus:new Set(),
      searchParts:new Set(),
      stockKeys:new Set(),
      retailers:new Set(),
      seasons:new Set(),
      yearCodes:new Set(),
      categoryLarge:new Set(),
      categorySmall:new Set(),
      qty:0,
      gross:0,
      payment:0,
      orders:0,
      validGross:0,
      validPayment:0,
      avgQty:0,
      avgPayment:0,
      stock_qty:0,
      received_qty:0,
      stock_known:false,
      dailyMap:{{}}
    }};
    const g = grouped[key];
    const stockQty = Math.max(0, Number(r.stock_qty || 0));
    const stockKey = detailStockKey(r, key);
    const stockMatches = detailStockMatchesGroup(r, groupInfo);
    if (r.match_sku && stockMatches) g.matchSkus.add(r.match_sku);
    [groupInfo.label, groupInfo.product, groupInfo.size, r.name, r.color, r.size].forEach(v=>addDetailSearchPart(g, v));
    if (stockMatches) [r.standard_name, r.stock_name, r.stock_barcode, r.match_sku].forEach(v=>addDetailSearchPart(g, v));
    g.retailers.add(r.retailer || '-');
    g.seasons.add(r.season || '-');
    const yearCode = rowBarcodeYearCode(r);
    if (yearCode) g.yearCodes.add(yearCode);
    g.categoryLarge.add(r.category_large || '-');
    g.categorySmall.add(r.category_small || '-');
    g.qty += Number(r.qty || 0);
    g.gross += Number(r.gross || 0);
    g.payment += Number(r.payment || 0);
    g.orders += Number(r.orders || 0);
    const avgBase = unitPriceBase(r);
    g.avgQty += avgBase.qty;
    g.avgPayment += avgBase.payment;
    if (validDiscount(r.gross, r.payment) !== null) {{
      g.validGross += Number(r.gross || 0);
      g.validPayment += Number(r.payment || 0);
    }}
    if (stockMatches && stockQty > 0 && stockKey && !g.stockKeys.has(stockKey)) {{
      g.stock_qty += stockQty;
      g.stockKeys.add(stockKey);
    }}
    g.received_qty = Math.max(g.received_qty, Number(r.received_qty || 0));
    if (stockMatches && (r.stock_barcode || r.stock_name || stockQty > 0)) g.stock_known = true;
    (r.daily || []).forEach(d=>{{
      if (!g.dailyMap[d.date]) g.dailyMap[d.date] = {{date:d.date,qty:0,gross:0,payment:0,orders:0}};
      g.dailyMap[d.date].qty += Number(d.qty || 0);
      g.dailyMap[d.date].gross += Number(d.gross || 0);
      g.dailyMap[d.date].payment += Number(d.payment || 0);
      g.dailyMap[d.date].orders += Number(d.orders || 0);
    }});
  }});

  let rows = Object.values(grouped).map(g=>{{
    const avg_unit = g.avgQty ? g.avgPayment / g.avgQty : 0;
    const match_skus = Array.from(g.matchSkus);
    return {{
      ...g,
      match_skus,
      match_sku: match_skus.length > 1 ? `${{match_skus[0]}} 외 ${{match_skus.length - 1}}` : (match_skus[0] || ''),
      searchText:Array.from(g.searchParts).join(' ').toLowerCase(),
      retailers:Array.from(g.retailers),
      seasons:Array.from(g.seasons),
      yearCodes:Array.from(g.yearCodes),
      years:Array.from(g.yearCodes).map(code=>barcodeYearLabel(code)).filter(Boolean),
      categoryLarge:Array.from(g.categoryLarge),
      categorySmall:Array.from(g.categorySmall),
      daily:Object.values(g.dailyMap).sort((a,b)=>a.date.localeCompare(b.date)),
      avg_unit
    }};
  }});

  if (q) rows = rows.filter(r=>(r.searchText||'').includes(q));

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
    r.reorder_year_eligible = (r.yearCodes || []).some(code => REORDER_ALLOWED_YEAR_CODES.has(code));
    r.reorder_reasons = [];
    if (r.reorder_year_eligible && r.stock_known && r.weekly_rate >= 7) r.reorder_reasons.push('주간 판매율 7% 이상');
    if (r.reorder_year_eligible && r.stock_known && r.stock_rate >= 20) r.reorder_reasons.push('재고 소진율 20% 이상');
  }});

  const reorderRows = rows
    .filter(r=>r.reorder_reasons.length)
    .sort((a,b)=>(
      Math.max(...(b.years||[]).map(s=>parseInt(s,10)).filter(Number.isFinite), -1) -
      Math.max(...(a.years||[]).map(s=>parseInt(s,10)).filter(Number.isFinite), -1)
    ) || b.weekly_rate-a.weekly_rate || b.stock_rate-a.stock_rate);
  document.getElementById('reorderSummary').textContent = `${{rows.filter(r=>r.reorder_reasons.length).length}}개 상품`;
  document.getElementById('reorderAlerts').innerHTML = reorderRows.length
    ? reorderRows.map(r=>`<div class="rank-item"><div><div class="rank-name">${{r.standard_name}}</div><div class="rank-meta">${{r.reorder_reasons.join(' · ')}} · 현재고 ${{fmt(r.stock_qty)}} · 주간 ${{fmt(r.weekly_qty)}}개</div></div><div class="rank-value">${{pct(r.weekly_rate)}}</div></div>`).join('')
    : '<div style="padding:14px;color:var(--ink3)">현재 기준 리오더 알림 상품 없음</div>';

  applyDetailSort(rows);
  updateDetailSortMarks();

  const tbody = document.getElementById('detailRows');
  const filterBody = document.getElementById('detailFilterRows');
  filterBody.innerHTML = `
    <th><select id="tableAlertFilter"><option value="">전체</option><option value="확인">확인</option><option value="정상">정상</option><option value="재고확인">재고확인</option></select></th>
    <th><input id="tableSkuFilter" placeholder="바코드"></th>
    <th><input id="tableNameFilter" placeholder="상품명"></th>
    <th><select id="tableYearFilter"><option value="">전체</option>${{uniqSorted(rows.flatMap(r=>r.years)).map(v=>`<option value="${{v}}">${{v}}</option>`).join('')}}</select></th>
    <th><select id="tableSeasonFilter"><option value="">전체</option>${{uniqSorted(rows.flatMap(r=>r.seasons)).map(v=>`<option value="${{v}}">${{v}}</option>`).join('')}}</select></th>
    <th><select id="tableCategoryFilter"><option value="">전체</option>${{uniqSorted(rows.flatMap(r=>r.categoryLarge.concat(r.categorySmall))).map(v=>`<option value="${{v}}">${{v}}</option>`).join('')}}</select></th>
    <th colspan="9"></th>`;

  const tableFilters = {{
    alert:'',
    sku:'',
    name:'',
    year:'',
    season:'',
    category:''
  }};
  const applyTableFilters = () => {{
    tableFilters.alert = document.getElementById('tableAlertFilter').value;
    tableFilters.sku = (document.getElementById('tableSkuFilter').value||'').toLowerCase();
    tableFilters.name = (document.getElementById('tableNameFilter').value||'').toLowerCase();
    tableFilters.year = document.getElementById('tableYearFilter').value;
    tableFilters.season = document.getElementById('tableSeasonFilter').value;
    tableFilters.category = document.getElementById('tableCategoryFilter').value;
    drawDetailRows(rows.filter(r=>{{
      const alertText = detailAlertText(r);
      if (tableFilters.alert && alertText !== tableFilters.alert) return false;
      if (tableFilters.sku && !((r.match_skus||[]).join(' ').toLowerCase().includes(tableFilters.sku) || (r.searchText||'').includes(tableFilters.sku))) return false;
      if (tableFilters.name && !(r.searchText||'').includes(tableFilters.name)) return false;
      if (tableFilters.year && !r.years.includes(tableFilters.year)) return false;
      if (tableFilters.season && !r.seasons.includes(tableFilters.season)) return false;
      if (tableFilters.category && !r.categoryLarge.concat(r.categorySmall).includes(tableFilters.category)) return false;
      return true;
    }}));
  }};
  ['tableAlertFilter','tableSkuFilter','tableNameFilter','tableYearFilter','tableSeasonFilter','tableCategoryFilter'].forEach(id=>{{
    const el = document.getElementById(id);
    el.addEventListener('input', applyTableFilters);
    el.addEventListener('change', applyTableFilters);
  }});

  if (!rows.length) {{ tbody.innerHTML='<tr><td colspan="15" style="text-align:center;color:var(--ink3);padding:28px">검색 결과 없음</td></tr>'; return; }}

function drawDetailRows(displayRows) {{
  if (!displayRows.length) {{
    tbody.innerHTML='<tr><td colspan="15" style="text-align:center;color:var(--ink3);padding:28px">검색 결과 없음</td></tr>';
    return;
  }}
  tbody.innerHTML = displayRows.map(r=>{{
    const disc    = r.validGross > 0 && r.validPayment <= r.validGross ? (1 - r.validPayment / r.validGross) * 100 : null;
    const stockQty= r.stock_qty ?? '-';
    const weeklyQty= r.weekly_qty ?? '-';
    const alertB  = !r.stock_known ? 'badge-amber' : r.reorder_reasons.length ? 'badge-red' : 'badge-green';
    const alertT  = !r.stock_known ? '재고확인' : r.reorder_reasons.length ? '확인' : '정상';
    const alertTitle = !r.stock_known
      ? '재고 매칭 실패(바코드/상품명). WMS 출고상품명/바코드 확인 후 data.json 매칭키 정비가 필요합니다.'
      : r.reorder_reasons.length
        ? (r.reorder_reasons.join(' / ') + ' · 리오더/재고 보충 검토가 필요합니다.')
        : '';
    return `<tr>
      <td><span class="badge ${{alertB}}" title="${{alertTitle}}">${{alertT}}</span></td>
      <td class="td-mono" title="${{(r.match_skus||[]).join(', ')}}">${{r.match_sku || '-'}}</td>
      <td class="td-main standard-name-cell" title="${{r.standard_name}}">${{r.standard_name||'-'}}</td>
      <td><span class="badge badge-gray">${{(r.years||[]).join(', ') || '-'}}</span></td>
      <td><span class="badge badge-blue">${{r.seasons.join(', ')}}</span></td>
      <td><span class="badge badge-indigo">${{r.categoryLarge.concat(r.categorySmall).filter(v=>v && v !== '-').join(' / ') || '-'}}</span></td>
      <td class="num">${{fmt(r.qty)}}</td>
      <td class="num">${{weeklyQty!=='-'?fmt(weeklyQty):'-'}}</td>
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
['q','seasonFilter','categoryLargeFilter','categorySmallFilter'].forEach(id=>{{
  document.getElementById(id).addEventListener('input',renderDetail);
  document.getElementById(id).addEventListener('change',renderDetail);
}});
document.getElementById('sortBy').addEventListener('change', e=>{{
  const value = e.target.value;
  detailTableSort = {{
    key: value === 'weeklyRate' ? 'weeklyRate' : value === 'qty' ? 'qty' : value === 'name' ? 'name' : 'payment',
    dir: value === 'name' ? 'asc' : 'desc'
  }};
  renderDetail();
}});
document.querySelectorAll('[data-detail-sort]').forEach(th=>{{
  th.addEventListener('click', e=>{{
    if (e.target && e.target.classList && e.target.classList.contains('col-resizer')) return;
    const key = th.dataset.detailSort;
    detailTableSort = {{
      key,
      dir: detailTableSort.key === key && detailTableSort.dir === 'desc' ? 'asc' : 'desc'
    }};
    renderDetail();
  }});
}});
function initStandardNameResizer(){{
  const handle = document.getElementById('standardNameResizer');
  if (!handle || handle.dataset.ready) return;
  handle.dataset.ready = '1';
  handle.addEventListener('mousedown', e=>{{
    e.preventDefault();
    const startX = e.clientX;
    const startWidth = parseInt(getComputedStyle(document.documentElement).getPropertyValue('--standard-name-col-width')) || 280;
    const onMove = ev=>{{
      const width = Math.max(180, Math.min(760, startWidth + ev.clientX - startX));
      document.documentElement.style.setProperty('--standard-name-col-width', width + 'px');
    }};
    const onUp = ()=>{{
      window.removeEventListener('mousemove', onMove);
      window.removeEventListener('mouseup', onUp);
    }};
    window.addEventListener('mousemove', onMove);
    window.addEventListener('mouseup', onUp);
  }});
}}
initStandardNameResizer();

// TODO LIST -------------------------------------------------------------
const TODO_KEY = 'plaknitTodoProjects.v1';
let TODO_SUPABASE_REST = {json.dumps(TODO_SUPABASE_REST, ensure_ascii=False)};
let TODO_SUPABASE_PROJECT = {json.dumps(TODO_SUPABASE_PROJECT, ensure_ascii=False)};
let TODO_SUPABASE_KEY = {json.dumps(TODO_SUPABASE_KEY, ensure_ascii=False)};
let TODO_SUPABASE_BUCKET = {json.dumps(TODO_SUPABASE_BUCKET, ensure_ascii=False)};
let TODO_HEADERS = buildTodoHeaders();
let todoEditId = null;
let todoItems = [];
const todoStatusText = {{todo:'할 일', progress:'진행중', done:'완료', delay:'지연'}};

function buildTodoHeaders() {{
  return {{
    apikey: TODO_SUPABASE_KEY,
    Authorization: `Bearer ${{TODO_SUPABASE_KEY}}`,
    'Content-Type': 'application/json'
  }};
}}
function hasTodoSupabaseConfig() {{
  return Boolean(TODO_SUPABASE_REST && TODO_SUPABASE_PROJECT && TODO_SUPABASE_KEY);
}}
async function ensureTodoSupabaseConfig() {{
  if (hasTodoSupabaseConfig()) return true;
  try {{
    const res = await fetch('/api/dashboard-config', {{cache:'no-store'}});
    if (!res.ok) throw new Error(await res.text());
    const cfg = await res.json();
    TODO_SUPABASE_PROJECT = (cfg.url || cfg.projectUrl || '').replace(/\\/$/, '');
    TODO_SUPABASE_REST = cfg.rest || (TODO_SUPABASE_PROJECT ? `${{TODO_SUPABASE_PROJECT}}/rest/v1` : '');
    TODO_SUPABASE_KEY = cfg.anonKey || cfg.key || '';
    TODO_SUPABASE_BUCKET = cfg.bucket || TODO_SUPABASE_BUCKET || 'todo-files';
    TODO_HEADERS = buildTodoHeaders();
  }} catch (e) {{
    console.warn('Supabase config load failed.', e);
  }}
  return hasTodoSupabaseConfig();
}}

function todayISO() {{
  const d = new Date();
  return `${{d.getFullYear()}}-${{String(d.getMonth()+1).padStart(2,'0')}}-${{String(d.getDate()).padStart(2,'0')}}`;
}}
function addDaysISO(base, days) {{
  const d = new Date(base + 'T00:00:00');
  d.setDate(d.getDate() + days);
  return `${{d.getFullYear()}}-${{String(d.getMonth()+1).padStart(2,'0')}}-${{String(d.getDate()).padStart(2,'0')}}`;
}}
function isLegacySampleTodo(item) {{
  const project = String(item?.project || '');
  const task = String(item?.task || '');
  return String(item?.id || '').startsWith('sample-') ||
    project === 'REXY 2 개발' ||
    project === 'SEERSUCKER 개발' ||
    project === '개발편직(경편)' ||
    task.includes('원사 입고') ||
    task.includes('Seersucker 개발') ||
    task.includes('air hole 트리코트');
}}
function loadTodoItems() {{
  try {{
    const saved = JSON.parse(localStorage.getItem(TODO_KEY) || '[]');
    if (Array.isArray(saved) && saved.length) {{
      const cleaned = saved.filter(item => !isLegacySampleTodo(item));
      if (cleaned.length !== saved.length) localStorage.setItem(TODO_KEY, JSON.stringify(cleaned));
      return cleaned;
    }}
  }} catch(e) {{}}
  return [];
}}
function saveTodoItems() {{
  localStorage.setItem(TODO_KEY, JSON.stringify(todoItems));
}}
function todoFromServer(row) {{
  return {{
    id: row.id,
    project: row.project || '미지정 프로젝트',
    task: row.task || '미지정 태스크',
    owner: row.owner || '',
    status: row.status || 'todo',
    start: row.start_date || '',
    end: row.end_date || '',
    doneDate: row.done_date || '',
    memo: row.memo || '',
    fileName: row.file_name || '',
    filePath: row.file_path || ''
  }};
}}
function todoToServer(item) {{
  return {{
    project: item.project || '미지정 프로젝트',
    task: item.task || '미지정 태스크',
    owner: item.owner || '',
    status: item.status || 'todo',
    start_date: item.start || null,
    end_date: item.end || null,
    done_date: item.doneDate || null,
    memo: item.memo || '',
    file_name: item.fileName || '',
    file_path: item.filePath || ''
  }};
}}
async function fetchTodoRemote() {{
  if (!(await ensureTodoSupabaseConfig())) throw new Error('Supabase config is missing.');
  const res = await fetch(`${{TODO_SUPABASE_REST}}/todos?select=*&order=created_at.asc`, {{
    headers: TODO_HEADERS
  }});
  if (!res.ok) throw new Error(await res.text());
  return (await res.json()).map(todoFromServer);
}}
async function refreshTodoFromSupabase() {{
  try {{
    const remote = await fetchTodoRemote();
    todoItems = remote;
    saveTodoItems();
    renderTodo();
  }} catch (e) {{
    console.warn('Supabase todo load failed. Local fallback is active.', e);
  }}
}}
async function uploadTodoFile(file) {{
  if (!file) return {{fileName:'', filePath:''}};
  if (!(await ensureTodoSupabaseConfig())) throw new Error('Supabase config is missing.');
  const safeName = file.name.replace(/[^a-zA-Z0-9._-]/g, '_');
  const path = `${{Date.now()}}-${{safeName}}`;
  const res = await fetch(`${{TODO_SUPABASE_PROJECT}}/storage/v1/object/${{TODO_SUPABASE_BUCKET}}/${{path}}`, {{
    method: 'POST',
    headers: {{
      apikey: TODO_SUPABASE_KEY,
      Authorization: `Bearer ${{TODO_SUPABASE_KEY}}`,
      'x-upsert': 'true',
      'Content-Type': file.type || 'application/octet-stream'
    }},
    body: file
  }});
  if (!res.ok) throw new Error(await res.text());
  return {{
    fileName: file.name,
    filePath: `${{TODO_SUPABASE_PROJECT}}/storage/v1/object/public/${{TODO_SUPABASE_BUCKET}}/${{path}}`
  }};
}}
async function saveTodoRemote(item, isEdit) {{
  if (!(await ensureTodoSupabaseConfig())) throw new Error('Supabase config is missing.');
  const payload = todoToServer(item);
  const url = isEdit
    ? `${{TODO_SUPABASE_REST}}/todos?id=eq.${{encodeURIComponent(item.id)}}`
    : `${{TODO_SUPABASE_REST}}/todos`;
  const res = await fetch(url, {{
    method: isEdit ? 'PATCH' : 'POST',
    headers: {{...TODO_HEADERS, Prefer:'return=representation'}},
    body: JSON.stringify(payload)
  }});
  if (!res.ok) throw new Error(await res.text());
  const rows = await res.json();
  return rows[0] ? todoFromServer(rows[0]) : item;
}}
async function deleteTodoRemote(id) {{
  if (!(await ensureTodoSupabaseConfig())) throw new Error('Supabase config is missing.');
  const res = await fetch(`${{TODO_SUPABASE_REST}}/todos?id=eq.${{encodeURIComponent(id)}}`, {{
    method: 'DELETE',
    headers: TODO_HEADERS
  }});
  if (!res.ok) throw new Error(await res.text());
}}
async function deleteTodoProjectRemote(project) {{
  const remoteItems = (await fetchTodoRemote()).filter(item=>item.project === project);
  if (!remoteItems.length) return 0;
  const idFilter = `in.(${{remoteItems.map(item=>item.id).join(',')}})`;
  const res = await fetch(`${{TODO_SUPABASE_REST}}/todos?id=${{encodeURIComponent(idFilter)}}`, {{
    method: 'DELETE',
    headers: {{...TODO_HEADERS, Prefer:'return=representation'}}
  }});
  if (!res.ok) throw new Error(await res.text());
  const deleted = await res.json();
  if (!Array.isArray(deleted) || deleted.length !== remoteItems.length) {{
    throw new Error(`프로젝트 삭제 결과가 일치하지 않습니다. 요청 ${{remoteItems.length}}건, 삭제 ${{Array.isArray(deleted) ? deleted.length : 0}}건`);
  }}
  const remaining = (await fetchTodoRemote()).filter(item=>item.project === project);
  if (remaining.length) throw new Error(`프로젝트 태스크 ${{remaining.length}}건이 남아 있습니다.`);
  return deleted.length;
}}
async function deleteTodoProject(project, button) {{
  const projectItems = todoItems.filter(item=>item.project === project);
  if (!projectItems.length) return;
  if (!confirm(`'${{project}}' 프로젝트와 태스크 ${{projectItems.length}}개를 모두 삭제하시겠습니까?\n삭제 후 복구할 수 없습니다.`)) return;
  button.disabled = true;
  try {{
    await deleteTodoProjectRemote(project);
  }} catch(e) {{
    console.warn('Supabase project delete failed. Local project delete applied.', e);
  }}
  todoItems = todoItems.filter(item=>item.project !== project);
  saveTodoItems();
  renderTodo();
}}
function todoStatusClass(status) {{
  if (status === 'done') return 'done';
  if (status === 'delay') return 'delay';
  if (status === 'progress') return 'progress';
  return 'todo';
}}
function daysBetween(a,b) {{
  return Math.round((new Date(b+'T00:00:00') - new Date(a+'T00:00:00')) / 86400000);
}}
function todoRange() {{
  const dates = todoItems.flatMap(i=>[i.start,i.end]).filter(Boolean).sort();
  const start = dates[0] || todayISO();
  const end = dates[dates.length-1] || addDaysISO(start,30);
  return {{start:addDaysISO(start,-2), end:addDaysISO(end,5)}};
}}
function renderTodoStats() {{
  const counts = {{all:todoItems.length,todo:0,progress:0,done:0,delay:0}};
  todoItems.forEach(i=>counts[i.status] = (counts[i.status] || 0) + 1);
  document.getElementById('todoStatAll').textContent = counts.all;
  document.getElementById('todoStatProgress').textContent = counts.progress;
  document.getElementById('todoStatDone').textContent = counts.done;
  document.getElementById('todoStatDelay').textContent = counts.delay;
  document.getElementById('todoStatTodo').textContent = counts.todo;
  const today = new Date().toLocaleDateString('ko-KR', {{weekday:'long', year:'numeric', month:'long', day:'numeric'}});
  document.getElementById('todoToday').textContent = `오늘: ${{today}}`;
}}
function renderTodoGantt() {{
  const el = document.getElementById('todoGantt');
  const range = todoRange();
  const span = Math.max(1, daysBetween(range.start, range.end));
  const byProject = {{}};
  todoItems.forEach(item => {{
    if (!byProject[item.project]) byProject[item.project] = [];
    byProject[item.project].push(item);
  }});
  const projectEntries = Object.entries(byProject);
  el.innerHTML = projectEntries.map(([project,items],projectIndex)=>`
    <section class="gantt-project">
      <div class="gantt-head">
        <div class="gantt-head-main"><span>▾ ${{project}}</span><span class="gantt-meta">${{items.length}}개</span></div>
        <button class="todo-project-delete" type="button" data-delete-project="${{projectIndex}}" title="프로젝트와 모든 태스크 삭제">프로젝트 삭제</button>
      </div>
      ${{items.map(item=>{{
        const left = Math.max(0, Math.min(100, daysBetween(range.start, item.start || range.start) / span * 100));
        const width = Math.max(4, Math.min(100-left, (Math.max(1, daysBetween(item.start || range.start, item.end || item.start || range.end)+1) / span * 100)));
        return `<div class="gantt-row" data-open-todo="${{item.id}}">
          <div class="gantt-task-name"><span class="gantt-dot ${{todoStatusClass(item.status)}}"></span><span title="${{item.memo || ''}}">${{item.task}}</span></div>
          <div class="gantt-timeline">
            <div class="gantt-today"></div>
            <div class="gantt-bar ${{todoStatusClass(item.status)}}" style="left:${{left}}%;width:${{width}}%">${{item.owner || todoStatusText[item.status]}}</div>
          </div>
        </div>`;
      }}).join('')}}
    </section>
  `).join('') || '<div class="panel">등록된 일정이 없습니다.</div>';
  el.querySelectorAll('[data-open-todo]').forEach(node=>node.addEventListener('click',()=>openTodoForm(node.dataset.openTodo)));
  el.querySelectorAll('[data-delete-project]').forEach(button=>button.addEventListener('click',async event=>{{
    event.stopPropagation();
    const entry = projectEntries[Number(button.dataset.deleteProject)];
    if (entry) await deleteTodoProject(entry[0], button);
  }}));
}}
function renderTodoCalendar() {{
  const grid = document.getElementById('todoCalendarGrid');
  const start = todayISO();
  const days = Array.from({{length:35}}, (_,i)=>addDaysISO(start,i));
  grid.innerHTML = days.map(day=>{{
    const dayItems = todoItems.filter(i=>i.start <= day && i.end >= day);
    return `<div class="calendar-day">
      <div class="calendar-date">${{fmtD(day)}}</div>
      ${{dayItems.slice(0,4).map(i=>`<div class="calendar-chip ${{todoStatusClass(i.status)}}" data-open-todo="${{i.id}}" title="${{i.project}} / ${{i.task}}">${{i.task}}</div>`).join('')}}
    </div>`;
  }}).join('');
  grid.querySelectorAll('[data-open-todo]').forEach(node=>node.addEventListener('click',e=>{{e.stopPropagation();openTodoForm(node.dataset.openTodo);}}));
}}
function renderTodoList() {{
  const body = document.getElementById('todoRows');
  body.innerHTML = todoItems
    .slice()
    .sort((a,b)=>(a.end||'').localeCompare(b.end||''))
    .map(item=>`<tr data-open-todo="${{item.id}}">
      <td>${{item.project}}</td><td class="td-main">${{item.task}}</td><td>${{item.owner || '-'}}</td>
      <td><span class="badge ${{item.status==='delay'?'badge-red':item.status==='done'?'badge-blue':item.status==='progress'?'badge-green':'badge-gray'}}">${{todoStatusText[item.status] || item.status}}</span></td>
      <td>${{item.start || '-'}}</td><td>${{item.end || '-'}}</td><td>${{item.doneDate || '-'}}</td>
      <td>${{item.filePath ? `<a href="${{item.filePath}}" target="_blank" rel="noopener">${{item.fileName || 'file'}}</a>` : (item.fileName || '-')}}</td><td style="max-width:220px;white-space:normal">${{item.memo || '-'}}</td>
      <td><div class="todo-row-actions"><button data-edit="${{item.id}}">수정</button><button data-delete="${{item.id}}">삭제</button></div></td>
    </tr>`).join('');
  body.querySelectorAll('[data-open-todo]').forEach(row=>row.addEventListener('click',()=>openTodoForm(row.dataset.openTodo)));
  body.querySelectorAll('a').forEach(link=>link.addEventListener('click',e=>e.stopPropagation()));
  body.querySelectorAll('[data-edit]').forEach(btn=>btn.addEventListener('click',e=>{{e.stopPropagation();openTodoForm(btn.dataset.edit);}}));
  body.querySelectorAll('[data-delete]').forEach(btn=>btn.addEventListener('click',async e=>{{e.stopPropagation();
    const id = btn.dataset.delete;
    try {{
      await deleteTodoRemote(id);
    }} catch(e) {{
      console.warn('Supabase todo delete failed. Local delete applied.', e);
    }}
    todoItems = todoItems.filter(i=>i.id !== id);
    saveTodoItems();
    renderTodo();
  }}));
}}
function renderTodoDeadline() {{
  const target = document.getElementById('todoDeadline');
  const today = todayISO();
  const items = todoItems
    .filter(i=>i.status !== 'done')
    .map(i=>({{...i, dday:daysBetween(today, i.end || today)}}))
    .filter(i=>i.dday <= 7)
    .sort((a,b)=>a.dday-b.dday)
    .slice(0,8);
  target.innerHTML = items.length ? items.map(i=>`<div class="todo-deadline-item"><b>${{i.task}}</b><span>D${{i.dday>=0?'-'+i.dday:'+'+Math.abs(i.dday)}}</span><div>${{i.project}}</div></div>`).join('') : '<div class="todo-deadline-item">마감 임박 일정 없음</div>';
}}
function renderTodo() {{
  renderTodoStats();
  renderTodoGantt();
  renderTodoCalendar();
  renderTodoList();
  renderTodoDeadline();
}}
function openTodoForm(id) {{
  todoEditId = id || null;
  const item = todoEditId ? todoItems.find(i=>i.id===todoEditId) : null;
  document.getElementById('todoFormTitle').textContent = item ? '프로젝트 상세/수정' : '프로젝트 등록';
  document.getElementById('todoFormSub').textContent = item ? `${{item.project || '-'}} / ${{item.task || '-'}}` : '프로젝트 또는 태스크를 등록합니다.';
  document.getElementById('todoSubmit').textContent = item ? '수정 저장' : '등록';
  document.getElementById('todoProject').value = item?.project || '';
  document.getElementById('todoTask').value = item?.task || '';
  document.getElementById('todoOwner').value = item?.owner || '';
  document.getElementById('todoStatus').value = item?.status || 'todo';
  document.getElementById('todoStart').value = item?.start || todayISO();
  document.getElementById('todoEnd').value = item?.end || addDaysISO(todayISO(),7);
  document.getElementById('todoDoneDate').value = item?.doneDate || '';
  document.getElementById('todoMemo').value = item?.memo || '';
  document.getElementById('todoFileName').textContent = item?.fileName || '첨부 파일 없음';
  document.getElementById('todoForm').classList.add('active');
  document.getElementById('todoForm').scrollIntoView({{behavior:'smooth', block:'start'}});
}}
function closeTodoForm() {{
  todoEditId = null;
  document.getElementById('todoForm').classList.remove('active');
  document.getElementById('todoFile').value = '';
}}
function submitTodo() {{
  const file = document.getElementById('todoFile').files[0];
  const item = {{
    id: todoEditId || String(Date.now()),
    project: document.getElementById('todoProject').value.trim() || '미지정 프로젝트',
    task: document.getElementById('todoTask').value.trim() || '미지정 태스크',
    owner: document.getElementById('todoOwner').value.trim(),
    status: document.getElementById('todoStatus').value,
    start: document.getElementById('todoStart').value || todayISO(),
    end: document.getElementById('todoEnd').value || todayISO(),
    doneDate: document.getElementById('todoDoneDate').value,
    memo: document.getElementById('todoMemo').value.trim(),
    fileName: file ? file.name : (todoEditId ? (todoItems.find(i=>i.id===todoEditId)?.fileName || '') : '')
  }};
  if (todoEditId) todoItems = todoItems.map(i=>i.id===todoEditId ? item : i);
  else todoItems.push(item);
  saveTodoItems();
  closeTodoForm();
  renderTodo();
}}
function exportTodoCsv() {{
  const header = ['project','task','owner','status','start','end','doneDate','fileName','memo'];
  const lines = [header.join(',')].concat(todoItems.map(i=>header.map(k=>`"${{String(i[k] || '').replaceAll('"','""')}}"`).join(',')));
  const blob = new Blob([lines.join('\\n')], {{type:'text/csv;charset=utf-8'}});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'plaknit_todo.csv';
  a.click();
  URL.revokeObjectURL(a.href);
}}
function importTodoFile(file) {{
  const reader = new FileReader();
  reader.onload = () => {{
    try {{
      if (file.name.toLowerCase().endsWith('.json')) {{
        const parsed = JSON.parse(reader.result);
        if (Array.isArray(parsed)) todoItems = parsed;
      }} else {{
        const lines = String(reader.result).split(/\\r?\\n/).filter(Boolean);
        const header = lines.shift().split(',').map(v=>v.replaceAll('"',''));
        todoItems = lines.map(line=>{{
          const cols = line.match(/("([^"]|"")*"|[^,]+)/g) || [];
          const obj = {{id:String(Date.now()+Math.random())}};
          header.forEach((h,idx)=>obj[h]=String(cols[idx] || '').replace(/^"|"$/g,'').replaceAll('""','"'));
          return obj;
        }});
      }}
      saveTodoItems();
      renderTodo();
    }} catch(e) {{
      alert('가져오기 파일을 확인해 주세요.');
    }}
  }};
  reader.readAsText(file, 'utf-8');
}}
async function submitTodo() {{
  const file = document.getElementById('todoFile').files[0];
  const prevItem = todoEditId ? todoItems.find(i=>i.id===todoEditId) : null;
  const item = {{
    id: todoEditId || String(Date.now()),
    project: document.getElementById('todoProject').value.trim() || '미지정 프로젝트',
    task: document.getElementById('todoTask').value.trim() || '미지정 태스크',
    owner: document.getElementById('todoOwner').value.trim(),
    status: document.getElementById('todoStatus').value,
    start: document.getElementById('todoStart').value || todayISO(),
    end: document.getElementById('todoEnd').value || todayISO(),
    doneDate: document.getElementById('todoDoneDate').value,
    memo: document.getElementById('todoMemo').value.trim(),
    fileName: file ? file.name : (prevItem?.fileName || ''),
    filePath: prevItem?.filePath || ''
  }};
  let saved = item;
  try {{
    if (file) {{
      const uploaded = await uploadTodoFile(file);
      item.fileName = uploaded.fileName;
      item.filePath = uploaded.filePath;
    }}
    saved = await saveTodoRemote(item, Boolean(todoEditId));
  }} catch(e) {{
    console.warn('Supabase todo save failed. Local save applied.', e);
  }}
  if (todoEditId) todoItems = todoItems.map(i=>i.id===todoEditId ? saved : i);
  else todoItems.push(saved);
  saveTodoItems();
  closeTodoForm();
  renderTodo();
}}
document.querySelectorAll('[data-todo-view]').forEach(btn=>btn.addEventListener('click',()=>{{
  document.querySelectorAll('[data-todo-view]').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.todo-board').forEach(b=>b.classList.remove('active'));
  btn.classList.add('active');
  document.getElementById('todo' + btn.dataset.todoView.charAt(0).toUpperCase() + btn.dataset.todoView.slice(1)).classList.add('active');
}}));
document.getElementById('todoAddProject').addEventListener('click',()=>openTodoForm());
document.getElementById('todoAddTask').addEventListener('click',()=>openTodoForm());
document.getElementById('todoSave').addEventListener('click',async ()=>{{saveTodoItems(); await refreshTodoFromSupabase(); alert('저장 및 동기화가 완료되었습니다.');}});
document.getElementById('todoCsv').addEventListener('click',exportTodoCsv);
document.getElementById('todoImportBtn').addEventListener('click',()=>document.getElementById('todoImportFile').click());
document.getElementById('todoImportFile').addEventListener('change',e=>{{if(e.target.files[0]) importTodoFile(e.target.files[0]);}});
document.getElementById('todoSubmit').addEventListener('click',submitTodo);
document.getElementById('todoCancel').addEventListener('click',closeTodoForm);
document.getElementById('todoFile').addEventListener('change',e=>{{document.getElementById('todoFileName').textContent = e.target.files[0]?.name || '첨부 파일 없음';}});
todoItems = loadTodoItems();
renderTodo();
refreshTodoFromSupabase();
refreshSalesTargetsRemote();

// ?? NAV TABS ???????????????????????????????????????????????
function activateTab(tab) {{
  document.querySelectorAll('.nav-item').forEach(b=>b.classList.remove('active'));
  document.querySelectorAll('.tab-panel').forEach(p=>p.classList.remove('active'));
  const nav = document.querySelector(`.nav-item[data-tab="${{tab}}"]`);
  if (nav) nav.classList.add('active');
  const panel = document.getElementById(tab);
  if (panel) panel.classList.add('active');
  const mobileToggle = document.getElementById('mobileNavToggle');
  if (mobileToggle && nav) mobileToggle.textContent = nav.querySelector('.nav-label')?.textContent?.trim() || '메뉴';
  if (window.location.hash !== `#${{tab}}`) history.replaceState(null, '', `#${{tab}}`);
  if(tab==='compare')   {{renderCompare();renderSalesTargets();}}
  if(tab==='calendar')  renderDaily();
  if(tab==='retailer')  renderRetailer();
  if(tab==='detail')    renderDetail();
}}

document.querySelectorAll('.nav-item[data-tab]').forEach(btn=>{{
  btn.addEventListener('click',(event)=>{{
    event.preventDefault();
    activateTab(btn.dataset.tab);
    document.getElementById('mainNav')?.classList.remove('open');
    document.getElementById('mobileNavToggle')?.classList.remove('open');
    document.getElementById('mobileNavToggle')?.setAttribute('aria-expanded','false');
  }});
}});

document.getElementById('mobileNavToggle')?.addEventListener('click',()=>{{
  const nav = document.getElementById('mainNav');
  const toggle = document.getElementById('mobileNavToggle');
  const open = !nav.classList.contains('open');
  nav.classList.toggle('open', open);
  toggle.classList.toggle('open', open);
  toggle.setAttribute('aria-expanded', open ? 'true' : 'false');
}});

window.addEventListener('hashchange',()=>{{
  const tab = location.hash.replace('#','');
  if (document.getElementById(tab)) activateTab(tab);
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
const initialTab = location.hash.replace('#','');
if (document.getElementById(initialTab)) activateTab(initialTab);
renderDaily();
</script>
</body>
</html>"""

OUT_FILE.write_text(html, encoding="utf-8")
PUBLIC_OUT_FILE.parent.mkdir(parents=True, exist_ok=True)
PUBLIC_OUT_FILE.write_text(html, encoding="utf-8")
print(f"✅ index.html 생성 완료 ({OUT_FILE.stat().st_size:,} bytes)")
