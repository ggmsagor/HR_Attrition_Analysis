# Employee Attrition Analysis
### Tools Used: Microsoft Excel · Microsoft Power BI

---

## Project Overview

This project explores the factors driving employee attrition at a mid-to-large organisation using the IBM HR Analytics dataset. The goal was to move beyond a raw attrition rate and identify **which employee segments are most at risk, why they leave, and where the company should act first**.

The dataset contains records for **1,470 employees** across 35 variables — covering demographics, job satisfaction, compensation, work patterns, and career history.

All analysis was done entirely in **Microsoft Excel** (pivot tables, calculated fields, summary tables) and **Microsoft Power BI** (interactive dashboard). No coding or statistical software was used.

---

bash

cat << 'EOF' > /mnt/user-data/outputs/dashboard_preview.html
<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>HR Attrition Dashboard – Preview</title>
<link href="https://fonts.googleapis.com/css2?family=DM+Sans:wght@300;400;500;600;700&family=DM+Mono:wght@400;500&display=swap" rel="stylesheet">
<style>
  *, *::before, *::after { box-sizing: border-box; margin: 0; padding: 0; }

  :root {
    --navy:    #1F3864;
    --blue:    #2E75B6;
    --lblue:   #BDD7EE;
    --red:     #C00000;
    --amber:   #ED7D31;
    --green:   #375623;
    --bg:      #F0F4F8;
    --card:    #FFFFFF;
    --border:  #D9E3EE;
    --text:    #1A2540;
    --muted:   #6B7A99;
  }

  body {
    font-family: 'DM Sans', sans-serif;
    background: var(--bg);
    color: var(--text);
    min-height: 100vh;
    padding: 0;
  }

  /* ── TOP BANNER ── */
  .banner {
    background: var(--navy);
    color: white;
    padding: 14px 36px;
    display: flex;
    align-items: center;
    justify-content: space-between;
    border-bottom: 3px solid var(--blue);
  }
  .banner h1 { font-size: 18px; font-weight: 600; letter-spacing: 0.02em; }
  .banner p  { font-size: 11px; color: #8FAAD4; margin-top: 2px; }
  .badge { background: var(--blue); color: white; font-size: 10px;
           padding: 3px 10px; border-radius: 3px; font-weight: 600; }

  /* ── PAGE TABS ── */
  .tabs { background: #162B50; display: flex; gap: 2px; padding: 0 36px; }
  .tab { color: #8FAAD4; font-size: 12px; font-weight: 500; padding: 10px 20px;
         cursor: pointer; border-bottom: 3px solid transparent; transition: all .2s; }
  .tab.active { color: white; border-bottom-color: #2E75B6; background: rgba(255,255,255,.04); }

  /* ── MAIN LAYOUT ── */
  .content { padding: 24px 36px; display: grid; gap: 20px; }

  /* ── KPI ROW ── */
  .kpi-row { display: grid; grid-template-columns: repeat(4, 1fr); gap: 16px; }
  .kpi { background: var(--card); border-radius: 6px; padding: 18px 20px;
         border-left: 4px solid var(--blue);
         box-shadow: 0 1px 4px rgba(0,0,0,.07); }
  .kpi.red    { border-left-color: var(--red); }
  .kpi.amber  { border-left-color: var(--amber); }
  .kpi.green  { border-left-color: var(--green); }
  .kpi-label  { font-size: 10px; font-weight: 600; color: var(--muted); text-transform: uppercase;
                letter-spacing: .08em; margin-bottom: 6px; }
  .kpi-value  { font-size: 30px; font-weight: 700; color: var(--navy); line-height: 1; }
  .kpi-sub    { font-size: 10px; color: var(--muted); margin-top: 5px; }

  /* ── GRID ROWS ── */
  .row2 { display: grid; grid-template-columns: 1.8fr 1fr; gap: 16px; }
  .row3 { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 16px; }
  .row2b { display: grid; grid-template-columns: 1fr 1.4fr; gap: 16px; }

  /* ── CHART CARD ── */
  .card { background: var(--card); border-radius: 6px; padding: 18px 20px;
          box-shadow: 0 1px 4px rgba(0,0,0,.07); }
  .card-title { font-size: 11px; font-weight: 700; color: var(--navy);
                text-transform: uppercase; letter-spacing: .07em; margin-bottom: 14px;
                display: flex; align-items: center; gap: 8px; }
  .card-title::before { content:''; display:block; width:3px; height:12px;
                        background:var(--blue); border-radius:2px; }

  /* ── BAR CHARTS ── */
  .bar-chart { display: flex; flex-direction: column; gap: 9px; }
  .bar-row { display: flex; align-items: center; gap: 10px; font-size: 10px; }
  .bar-label { width: 145px; color: var(--muted); text-align: right; flex-shrink:0; white-space:nowrap; overflow:hidden; text-overflow:ellipsis; }
  .bar-track { flex: 1; background: #EEF2F8; border-radius: 3px; height: 16px; position: relative; overflow:hidden; }
  .bar-fill  { height: 100%; border-radius: 3px; display: flex; align-items:center;
               justify-content:flex-end; padding-right:6px;
               font-size:9px; font-weight:700; color:white; transition: width .8s ease; }
  .bar-fill.red   { background: var(--red); }
  .bar-fill.amber { background: var(--amber); }
  .bar-fill.green { background: var(--green); }
  .bar-fill.blue  { background: var(--blue); }
  .avg-line { position:absolute; top:0; bottom:0; width:2px; background:#1F3864;
              opacity:.35; z-index:2; }

  /* ── DONUT CHART ── */
  .donut-wrap { display:flex; align-items:center; justify-content:center; gap:20px; flex-wrap:wrap; padding:4px 0; }
  .donut-legend { display:flex; flex-direction:column; gap:8px; }
  .legend-item { display:flex; align-items:center; gap:8px; font-size:10px; }
  .legend-dot { width:10px; height:10px; border-radius:50%; flex-shrink:0; }

  /* ── COLUMN CHART ── */
  .col-chart { display:flex; align-items:flex-end; gap:8px; height:100px; padding-top:8px; }
  .col-bar-wrap { flex:1; display:flex; flex-direction:column; align-items:center; gap:4px; }
  .col-bar { width:100%; border-radius:3px 3px 0 0; display:flex; align-items:flex-start;
             justify-content:center; padding-top:3px;
             font-size:9px; font-weight:700; color:white; transition: height .8s ease; }
  .col-label { font-size:9px; color:var(--muted); text-align:center; }

  /* ── HEATMAP ── */
  .heatmap { font-size:9px; }
  .heatmap table { width:100%; border-collapse:collapse; }
  .heatmap th { background:#F5F8FC; color:var(--muted); font-weight:600; font-size:9px;
                padding:5px 8px; text-align:center; border:1px solid var(--border); }
  .heatmap td { padding:5px 8px; text-align:center; font-weight:600; font-size:10px;
                border:1px solid var(--border); }

  /* ── SCATTER (simplified) ── */
  .scatter { position:relative; height:110px; border-left:1px solid var(--border);
             border-bottom:1px solid var(--border); }
  .scatter-dot { position:absolute; border-radius:50%; display:flex; align-items:center;
                 justify-content:center; font-size:7px; font-weight:700; color:white;
                 cursor:default; transition: transform .2s; }
  .scatter-dot:hover { transform:scale(1.3); z-index:10; }
  .scatter-tooltip { position:absolute; background:var(--navy); color:white; font-size:9px;
                     padding:3px 7px; border-radius:3px; white-space:nowrap; pointer-events:none;
                     bottom:105%; left:50%; transform:translateX(-50%); display:none; }
  .scatter-dot:hover .scatter-tooltip { display:block; }
  .axis-label { font-size:8px; color:var(--muted); position:absolute; }

  /* ── SLICER BAR ── */
  .slicer-row { display:flex; gap:8px; flex-wrap:wrap; align-items:center;
                padding-bottom:4px; border-bottom:1px solid var(--border); margin-bottom:8px; }
  .slicer-chip { background:#EEF2F8; border:1px solid var(--border); border-radius:4px;
                 font-size:10px; padding:4px 12px; cursor:pointer; color:var(--navy);
                 font-weight:500; transition: all .15s; }
  .slicer-chip:hover, .slicer-chip.active { background:var(--blue); color:white; border-color:var(--blue); }
  .slicer-label { font-size:10px; color:var(--muted); font-weight:600; }

  /* ── PREVIEW NOTICE ── */
  .preview-notice { background:var(--lblue); border-radius:6px; padding:10px 16px;
                    font-size:10px; color:var(--navy); display:flex; gap:8px;
                    align-items:center; font-weight:500; }
  .notice-icon { font-size:14px; }
</style>
</head>
<body>

<!-- BANNER -->
<div class="banner">
  <div>
    <h1>HR Employee Attrition Analysis</h1>
    <p>IBM HR Dataset &nbsp;|&nbsp; 1,470 Employees &nbsp;|&nbsp; Power BI Dashboard</p>
  </div>
  <span class="badge">PORTFOLIO PROJECT</span>
</div>

<!-- TABS -->
<div class="tabs">
  <div class="tab active">Overview</div>
  <div class="tab">Demographics</div>
  <div class="tab">Job &amp; Compensation</div>
  <div class="tab">Risk Heatmap</div>
</div>

<!-- PAGE CONTENT -->
<div class="content">

  <div class="preview-notice">
    <span class="notice-icon">📊</span>
    This is a static preview of the Power BI dashboard. Open <strong>HR_Attrition_Dashboard.pbit</strong> in Power BI Desktop for the fully interactive version.
  </div>

  <!-- KPI CARDS -->
  <div class="kpi-row">
    <div class="kpi">
      <div class="kpi-label">Total Employees</div>
      <div class="kpi-value">1,470</div>
      <div class="kpi-sub">Full workforce</div>
    </div>
    <div class="kpi red">
      <div class="kpi-label">Employees Left</div>
      <div class="kpi-value" style="color:var(--red)">237</div>
      <div class="kpi-sub">Confirmed attrition</div>
    </div>
    <div class="kpi amber">
      <div class="kpi-label">Attrition Rate</div>
      <div class="kpi-value" style="color:var(--amber)">16.1%</div>
      <div class="kpi-sub">Company-wide</div>
    </div>
    <div class="kpi green">
      <div class="kpi-label">Avg Tenure (Stayed)</div>
      <div class="kpi-value" style="color:var(--green)">7.4 yrs</div>
      <div class="kpi-sub">vs 5.1 yrs for leavers</div>
    </div>
  </div>

  <!-- SLICER -->
  <div class="slicer-row">
    <span class="slicer-label">Department:</span>
    <span class="slicer-chip active">All</span>
    <span class="slicer-chip">Human Resources</span>
    <span class="slicer-chip">Research &amp; Development</span>
    <span class="slicer-chip">Sales</span>
    &nbsp;&nbsp;
    <span class="slicer-label">OverTime:</span>
    <span class="slicer-chip active">All</span>
    <span class="slicer-chip">Yes</span>
    <span class="slicer-chip">No</span>
  </div>

  <!-- ROW 1: Bar chart + pie -->
  <div class="row2">
    <div class="card">
      <div class="card-title">Attrition Rate by Job Role</div>
      <div class="bar-chart">
        <!-- avg line at 16.1% → position ~ 40.25% of max 40% track → actual pct of 40 max -->
        <!-- bar width = rate/40 * 100 -->
        <div class="bar-row">
          <div class="bar-label">Sales Representative</div>
          <div class="bar-track">
            <div class="avg-line" style="left:40.25%"></div>
            <div class="bar-fill red" style="width:99.5%">39.8%</div>
          </div>
        </div>
        <div class="bar-row">
          <div class="bar-label">Laboratory Technician</div>
          <div class="bar-track">
            <div class="avg-line" style="left:40.25%"></div>
            <div class="bar-fill red" style="width:59.75%">23.9%</div>
          </div>
        </div>
        <div class="bar-row">
          <div class="bar-label">Human Resources</div>
          <div class="bar-track">
            <div class="avg-line" style="left:40.25%"></div>
            <div class="bar-fill red" style="width:57.75%">23.1%</div>
          </div>
        </div>
        <div class="bar-row">
          <div class="bar-label">Sales Executive</div>
          <div class="bar-track">
            <div class="avg-line" style="left:40.25%"></div>
            <div class="bar-fill amber" style="width:43.75%">17.5%</div>
          </div>
        </div>
        <div class="bar-row">
          <div class="bar-label">Research Scientist</div>
          <div class="bar-track">
            <div class="avg-line" style="left:40.25%"></div>
            <div class="bar-fill amber" style="width:40.25%">16.1%</div>
          </div>
        </div>
        <div class="bar-row">
          <div class="bar-label">Healthcare Rep.</div>
          <div class="bar-track">
            <div class="avg-line" style="left:40.25%"></div>
            <div class="bar-fill green" style="width:17.25%">6.9%</div>
          </div>
        </div>
        <div class="bar-row">
          <div class="bar-label">Manufacturing Dir.</div>
          <div class="bar-track">
            <div class="avg-line" style="left:40.25%"></div>
            <div class="bar-fill green" style="width:17.25%">6.9%</div>
          </div>
        </div>
        <div class="bar-row">
          <div class="bar-label">Manager</div>
          <div class="bar-track">
            <div class="avg-line" style="left:40.25%"></div>
            <div class="bar-fill green" style="width:12.25%">4.9%</div>
          </div>
        </div>
        <div class="bar-row">
          <div class="bar-label">Research Director</div>
          <div class="bar-track">
            <div class="avg-line" style="left:40.25%"></div>
            <div class="bar-fill green" style="width:6.25%">2.5%</div>
          </div>
        </div>
        <div style="font-size:8px;color:var(--muted);margin-top:4px;padding-left:155px">
          ▏ Company average (16.1%)
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">Employees Left by Department</div>
      <div class="donut-wrap" style="flex-direction:column;align-items:flex-start;gap:12px">
        <!-- Inline SVG donut -->
        <svg width="150" height="150" viewBox="0 0 36 36" style="margin:0 auto;display:block">
          <!-- Sales: 92/237 = 38.8% → stroke-dasharray 38.8 61.2, offset 0 -->
          <circle cx="18" cy="18" r="15.915" fill="none" stroke="#EEF2F8" stroke-width="5"/>
          <circle cx="18" cy="18" r="15.915" fill="none" stroke="#ED7D31" stroke-width="5"
                  stroke-dasharray="24.4 75.6" stroke-dashoffset="25" transform="rotate(-90 18 18)"/>
          <circle cx="18" cy="18" r="15.915" fill="none" stroke="#C00000" stroke-width="5"
                  stroke-dasharray="4.2 95.8" stroke-dashoffset="0.6" transform="rotate(-90 18 18)"/>
          <circle cx="18" cy="18" r="15.915" fill="none" stroke="#2E75B6" stroke-width="5"
                  stroke-dasharray="56.1 43.9" stroke-dashoffset="-24" transform="rotate(-90 18 18)"/>
          <text x="18" y="18" text-anchor="middle" dominant-baseline="central"
                font-family="DM Sans" font-size="5" font-weight="700" fill="#1F3864">237</text>
          <text x="18" y="22.5" text-anchor="middle" dominant-baseline="central"
                font-family="DM Sans" font-size="3" fill="#6B7A99">left</text>
        </svg>
        <div class="donut-legend" style="margin:0 auto">
          <div class="legend-item"><div class="legend-dot" style="background:#2E75B6"></div>R&amp;D: 133 (56.1%)</div>
          <div class="legend-item"><div class="legend-dot" style="background:#ED7D31"></div>Sales: 92 (38.8%)</div>
          <div class="legend-item"><div class="legend-dot" style="background:#C00000"></div>HR: 10 (4.2%)</div>
        </div>
      </div>
      <div style="margin-top:14px">
        <div class="card-title">Department Attrition Rate</div>
        <div class="bar-chart">
          <div class="bar-row">
            <div class="bar-label" style="width:100px">Sales</div>
            <div class="bar-track"><div class="bar-fill red" style="width:83%">20.6%</div></div>
          </div>
          <div class="bar-row">
            <div class="bar-label" style="width:100px">Human Resources</div>
            <div class="bar-track"><div class="bar-fill amber" style="width:76%">19.0%</div></div>
          </div>
          <div class="bar-row">
            <div class="bar-label" style="width:100px">Research &amp; Dev.</div>
            <div class="bar-track"><div class="bar-fill green" style="width:55%">13.8%</div></div>
          </div>
        </div>
      </div>
    </div>
  </div>

  <!-- ROW 2: Three column charts -->
  <div class="row3">
    <div class="card">
      <div class="card-title">Attrition by Age Group</div>
      <div class="col-chart">
        <div class="col-bar-wrap">
          <div class="col-bar red" style="height:91px;background:var(--red)">34.8%</div>
          <div class="col-label">18-25</div>
        </div>
        <div class="col-bar-wrap">
          <div class="col-bar amber" style="height:50px;background:var(--amber)">19.1%</div>
          <div class="col-label">26-35</div>
        </div>
        <div class="col-bar-wrap">
          <div class="col-bar green" style="height:24px;background:var(--green)">9.2%</div>
          <div class="col-label">36-45</div>
        </div>
        <div class="col-bar-wrap">
          <div class="col-bar" style="height:33px;background:#5A7FA8">12.5%</div>
          <div class="col-label">46-60</div>
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">Overtime Impact</div>
      <div style="display:flex;flex-direction:column;gap:12px;padding-top:4px">
        <div>
          <div style="font-size:9px;color:var(--muted);margin-bottom:5px">Works Overtime</div>
          <div style="background:#EEF2F8;border-radius:4px;height:32px;position:relative;overflow:hidden">
            <div style="width:30.5%;height:100%;background:var(--red);border-radius:4px;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700;color:white">30.5%</div>
          </div>
        </div>
        <div>
          <div style="font-size:9px;color:var(--muted);margin-bottom:5px">No Overtime</div>
          <div style="background:#EEF2F8;border-radius:4px;height:32px;position:relative;overflow:hidden">
            <div style="width:10.4%;height:100%;background:var(--green);border-radius:4px;display:flex;align-items:center;padding-left:6px;font-size:11px;font-weight:700;color:white">10.4%</div>
          </div>
        </div>
        <div style="font-size:9px;color:var(--muted);font-style:italic;border-top:1px solid var(--border);padding-top:8px">
          Overtime workers leave at <strong>3× the rate</strong> of those without overtime.
        </div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">Business Travel</div>
      <div class="bar-chart" style="gap:12px;padding-top:4px">
        <div class="bar-row">
          <div class="bar-label" style="width:80px">Non-Travel</div>
          <div class="bar-track"><div class="bar-fill green" style="width:32%">8.0%</div></div>
        </div>
        <div class="bar-row">
          <div class="bar-label" style="width:80px">Travel Rarely</div>
          <div class="bar-track"><div class="bar-fill blue" style="width:60%">15.0%</div></div>
        </div>
        <div class="bar-row">
          <div class="bar-label" style="width:80px">Travel Frequently</div>
          <div class="bar-track"><div class="bar-fill red" style="width:99.6%">24.9%</div></div>
        </div>
        <div style="font-size:9px;color:var(--muted);font-style:italic;border-top:1px solid var(--border);padding-top:8px">
          Frequent travellers leave at <strong>3× the rate</strong> of non-travellers.
        </div>
      </div>
    </div>
  </div>

  <!-- ROW 3: Scatter + income -->
  <div class="row2b">
    <div class="card">
      <div class="card-title">Income vs Attrition Rate by Job Role</div>
      <div class="scatter">
        <div class="axis-label" style="bottom:-14px;left:50%;transform:translateX(-50%)">Monthly Income ($) →</div>
        <div class="axis-label" style="top:50%;left:-18px;transform:translateY(-50%) rotate(-90deg)">Attrition Rate →</div>
        <!-- dots: x = income/20000, y = inverse of rate/40 -->
        <!-- Sales Rep: income~3600, rate 39.8% → x=18%, y=0.5% (high up) -->
        <div class="scatter-dot" style="width:22px;height:22px;background:var(--red);left:16%;bottom:0%;font-size:6px">
          SRep<div class="scatter-tooltip">Sales Rep | $3,600 | 39.8%</div>
        </div>
        <div class="scatter-dot" style="width:18px;height:18px;background:var(--red);left:20%;bottom:40%;font-size:6px">
          LT<div class="scatter-tooltip">Lab Tech | $3,900 | 23.9%</div>
        </div>
        <div class="scatter-dot" style="width:14px;height:14px;background:var(--amber);left:30%;bottom:57%;font-size:6px">
          SE<div class="scatter-tooltip">Sales Exec | $6,900 | 17.5%</div>
        </div>
        <div class="scatter-dot" style="width:16px;height:16px;background:var(--amber);left:33%;bottom:60%;font-size:6px">
          RS<div class="scatter-tooltip">Research Sci | $6,200 | 16.1%</div>
        </div>
        <div class="scatter-dot" style="width:18px;height:18px;background:var(--green);left:55%;bottom:83%;font-size:6px">
          HR<div class="scatter-tooltip">Healthcare Rep | $7,500 | 6.9%</div>
        </div>
        <div class="scatter-dot" style="width:18px;height:18px;background:var(--green);left:58%;bottom:83%;font-size:6px">
          MD<div class="scatter-tooltip">Mfg Dir | $8,100 | 6.9%</div>
        </div>
        <div class="scatter-dot" style="width:20px;height:20px;background:var(--green);left:68%;bottom:88%;font-size:6px">
          Mgr<div class="scatter-tooltip">Manager | $10,700 | 4.9%</div>
        </div>
        <div class="scatter-dot" style="width:14px;height:14px;background:var(--green);left:80%;bottom:94%;font-size:6px">
          RD<div class="scatter-tooltip">Research Dir | $15,900 | 2.5%</div>
        </div>
        <div style="position:absolute;bottom:4px;right:8px;font-size:8px;color:var(--muted)">Bubble size = headcount</div>
      </div>
    </div>

    <div class="card">
      <div class="card-title">Compensation Gap</div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;margin-bottom:16px">
        <div style="background:#F5F8FC;border-radius:6px;padding:14px;text-align:center;border:1px solid var(--border)">
          <div style="font-size:9px;color:var(--muted);font-weight:600;margin-bottom:6px">STAYED</div>
          <div style="font-size:22px;font-weight:700;color:var(--green)">$6,833</div>
          <div style="font-size:9px;color:var(--muted)">avg/month</div>
        </div>
        <div style="background:#FFF5F5;border-radius:6px;padding:14px;text-align:center;border:1px solid #FCCFCF">
          <div style="font-size:9px;color:var(--muted);font-weight:600;margin-bottom:6px">LEFT</div>
          <div style="font-size:22px;font-weight:700;color:var(--red)">$4,787</div>
          <div style="font-size:9px;color:var(--muted)">avg/month</div>
        </div>
      </div>
      <div style="background:var(--navy);border-radius:6px;padding:12px;text-align:center;color:white">
        <div style="font-size:9px;color:#8FAAD4;margin-bottom:4px">INCOME GAP</div>
        <div style="font-size:24px;font-weight:700">$2,046</div>
        <div style="font-size:9px;color:#8FAAD4">Leavers earn 30% less on average</div>
      </div>
      <div style="margin-top:14px">
        <div class="card-title">Attrition by Stock Option Level</div>
        <div class="bar-chart" style="gap:8px">
          <div class="bar-row">
            <div class="bar-label" style="width:80px">Level 0 (None)</div>
            <div class="bar-track"><div class="bar-fill red" style="width:98%">24.4%</div></div>
          </div>
          <div class="bar-row">
            <div class="bar-label" style="width:80px">Level 1 (Basic)</div>
            <div class="bar-track"><div class="bar-fill green" style="width:37.6%">9.4%</div></div>
          </div>
          <div class="bar-row">
            <div class="bar-label" style="width:80px">Level 2 (Mid)</div>
            <div class="bar-track"><div class="bar-fill green" style="width:30.4%">7.6%</div></div>
          </div>
          <div class="bar-row">
            <div class="bar-label" style="width:80px">Level 3 (Senior)</div>
            <div class="bar-track"><div class="bar-fill amber" style="width:70.4%">17.6%</div></div>
          </div>
        </div>
      </div>
    </div>
  </div>

</div><!-- /content -->

<div style="text-align:center;font-size:10px;color:var(--muted);padding:16px 0 24px">
  HR Attrition Analysis &nbsp;|&nbsp; Source: IBM HR Analytics Dataset (Kaggle) &nbsp;|&nbsp; Tools: Microsoft Excel &amp; Power BI
</div>

</body>
</html>
EOF
echo "HTML file written: $(wc -c < /mnt/user-data/outputs/dashboard_preview.html) bytes"
Output

HTML file written: 23706 bytes

---

## Files in This Repository

| File | Description |
|------|-------------|
| `HR_Attrition_Analysis.xlsx` | Excel workbook with all summary tables, breakdowns, and analysis sheets |
| `HR_Attrition_Dashboard.pbix` | Power BI dashboard file (interactive) |
| `WA_Fn-UseC_-HR-Employee-Attrition.csv` | Raw dataset (source data) |
| `Data_Dictionary.md` | Field definitions and scale descriptions |
| `PowerBI_Build_Guide.md` | Step-by-step guide for recreating the dashboard |

---

## Key Findings

**Overall attrition rate: 16.1%** (237 of 1,470 employees left)

### 1. Job Role is the strongest predictor of attrition
Sales Representatives have a **39.8%** attrition rate — the highest across all roles. Laboratory Technicians (23.9%) and HR staff (23.1%) are next. In contrast, Research Directors leave at just 2.5%.

### 2. Overtime doubles the attrition rate
Employees who work overtime leave at **30.5%** vs **10.4%** for those who don't. This is the single largest binary split in the dataset.

### 3. Young employees leave at a disproportionately high rate
The 18–25 age group has a **34.8%** attrition rate. This drops to 9.2% for the 36–45 bracket. Early career experiences — compensation, growth opportunities, manager quality — appear to play a central role.

### 4. Compensation gap is significant
Employees who left earned an average of **$4,787/month** vs **$6,833** for those who stayed — a 30% gap. While some of this is explained by role and seniority, it points to compensation as a retention lever.

### 5. Single employees leave at 2.5× the rate of divorced employees
Single employees have a **25.5%** attrition rate vs 10.1% for divorced colleagues. This likely reflects different levels of financial commitments and mobility.

### 6. Stock options are a strong retention tool
Employees with no stock options leave at **24.4%**. Those at Level 1 drop to just **9.4%**. The relationship is not perfectly linear (Level 3 rises to 17.6%) but the signal is clear at the entry level.

### 7. Frequent business travel elevates risk
Employees who travel frequently leave at **24.9%** — three times the rate of non-travellers (8.0%). This compounds with job satisfaction and work-life balance scores.

### 8. Job satisfaction predicts attrition, but the gap is moderate
Low satisfaction employees (score 1/4) leave at **22.8%** vs **11.3%** for the most satisfied. Satisfaction alone is not the whole story — structural factors appear equally important.

---

## Analysis Structure (Excel Workbook)

**Sheet 1 – Executive Summary**
High-level KPIs and a summary of the eight key findings. Suitable for a one-page management read.

**Sheet 2 – By Department & Role**
Full breakdown of attrition counts and rates for all three departments and all nine job roles, with risk-level classifications.

**Sheet 3 – Workforce Demographics**
Six analysis tables covering age group, gender, marital status, business travel, overtime, and work-life balance.

**Sheet 4 – Job Factors**
Attrition breakdowns by job satisfaction, environment satisfaction, job involvement, stock option level, income comparison, and tenure.

**Sheet 5 – Raw Data (Sample)**
A 50-row sample of the original dataset for reference. The full CSV is provided separately.

---

## Power BI Dashboard Pages

1. **Overview** — Attrition rate KPIs, trend breakdown, headline filters
2. **Demographics** — Age, gender, marital status, business travel slicers
3. **Job & Compensation** — Role-level attrition, salary distribution, stock options
4. **Risk Heatmap** — Cross-tab of department × overtime × satisfaction

---

## Data Source

IBM HR Analytics Employee Attrition & Performance dataset, made publicly available on Kaggle. The data is fictional and was created by IBM data scientists for learning purposes.

**Dataset size:** 1,470 rows × 35 columns
**No missing values** in the original dataset.

---

## Methodology Notes

- Attrition rate is calculated as: `(Employees Left / Total Employees) × 100`
- Satisfaction scores (1–4) represent: 1 = Low, 2 = Medium, 3 = High, 4 = Very High
- Work-Life Balance scores (1–4): 1 = Bad, 2 = Good, 3 = Better, 4 = Best
- Job Involvement scores (1–4): 1 = Low, 2 = Medium, 3 = High, 4 = Very High
- Age groups were manually bucketed: 18–25, 26–35, 36–45, 46–60
- All figures are based on the full dataset of 1,470 records

---

## How to Use This Project

1. Open `HR_Attrition_Analysis.xlsx` in Excel — all summary tables are pre-built and formatted
2. Open `HR_Attrition_Dashboard.pbix` in Power BI Desktop — use slicers to explore segments interactively
3. To refresh or extend the analysis, load `WA_Fn-UseC_-HR-Employee-Attrition.csv` as the data source in both tools
4. Refer to `Data_Dictionary.md` for field definitions

---

*This project was completed as part of a data analytics portfolio. Feedback and questions are welcome via GitHub Issues.*
