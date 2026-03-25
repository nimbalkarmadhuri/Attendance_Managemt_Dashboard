"""
dashboard.py  ── Attendance Management System (Full Upgrade)
=============================================================
Run standalone:
    python dashboard.py
    python dashboard.py /path/to/attendance_cleaned.xlsx
Opens browser at http://localhost:8050
"""

import sys, os, threading, time, webbrowser


import pandas as pd
import numpy as np
import plotly.graph_objects as go
from dash import Dash, dcc, html, dash_table, Input, Output, State
import dash_bootstrap_components as dbc

# ─────────────────────────────────────────────────────────────
# FILE PATH
# ─────────────────────────────────────────────────────────────




if len(sys.argv) > 1:
    EXCEL_FILE = sys.argv[1]
else:
    EXCEL_FILE = os.path.join(os.path.dirname(os.path.abspath(__file__)),
                              "attendance_cleaned.xlsx")

if not os.path.exists(EXCEL_FILE):
    print(f"\n  ERROR: Cannot find '{EXCEL_FILE}'")
    print("  Run attendance_app.py first to generate the report.\n")
    sys.exit(1)

PORT = 8050

# ─────────────────────────────────────────────────────────────
# COLORS
# ─────────────────────────────────────────────────────────────
C = {
    "present": "#22c55e", "absent":  "#f43f5e",
    "late":    "#f97316", "partial": "#a855f7",
    "blue":    "#6366f1", "bg":      "#faf5ff",
    "card":    "#ffffff", "border":  "#e9d5ff",
    "muted":   "#9333ea", "grad1":   "#7c3aed",
    "grad2":   "#db2777",
}

# ─────────────────────────────────────────────────────────────
# CSS
# ─────────────────────────────────────────────────────────────
EXTRA_CSS = """
@import url('https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&display=swap');
* { font-family: 'Poppins', sans-serif; }

.ams-hero {
    background: linear-gradient(135deg, #7c3aed 0%, #a21caf 50%, #db2777 100%);
    border-radius: 18px; padding: 32px 36px 28px; margin-bottom: 28px;
    position: relative; overflow: hidden;
    box-shadow: 0 8px 32px rgba(124,58,237,0.25);
}
.ams-hero::before {
    content:''; position:absolute; top:-60px; right:-60px;
    width:220px; height:220px; border-radius:50%;
    background:rgba(255,255,255,0.08);
}
.ams-hero::after {
    content:''; position:absolute; bottom:-40px; left:40px;
    width:140px; height:140px; border-radius:50%;
    background:rgba(255,255,255,0.06);
}
.ams-title {
    font-size:2rem; font-weight:800; color:#fff;
    text-align:center; letter-spacing:-0.5px; margin:0;
    text-shadow:0 2px 12px rgba(0,0,0,0.18);
}
.ams-subtitle { text-align:center; color:rgba(255,255,255,0.82); font-size:13px; margin-top:6px; margin-bottom:0; }
.ams-badge {
    display:inline-block; background:rgba(255,255,255,0.18);
    border:1px solid rgba(255,255,255,0.3); color:#fff;
    font-size:12px; font-weight:600; padding:4px 14px;
    border-radius:20px; backdrop-filter:blur(4px); margin-top:2px;
}
.ams-icons-row { display:flex; justify-content:center; gap:18px; margin-bottom:14px; }
.ams-icon-bubble {
    width:48px; height:48px; border-radius:14px;
    background:rgba(255,255,255,0.18); border:1px solid rgba(255,255,255,0.3);
    display:flex; align-items:center; justify-content:center;
    font-size:22px; backdrop-filter:blur(4px);
}
.kpi-card {
    border-radius:14px !important; border:1.5px solid #e9d5ff !important;
    box-shadow:0 2px 12px rgba(124,58,237,0.08);
    transition:transform 0.2s, box-shadow 0.2s;
}
.kpi-card:hover { transform:translateY(-3px); box-shadow:0 8px 24px rgba(124,58,237,0.15); }
.chart-card { border-radius:14px !important; border:1.5px solid #e9d5ff !important; box-shadow:0 2px 8px rgba(124,58,237,0.06); }
.chart-card .card-header {
    border-radius:14px 14px 0 0 !important;
    background:linear-gradient(90deg,#f5f3ff,#fdf4ff) !important;
    border-bottom:1.5px solid #e9d5ff !important;
    font-weight:600 !important; color:#7c3aed !important;
}
.filter-card {
    border-radius:14px !important; border:1.5px solid #e9d5ff !important;
    background:linear-gradient(135deg,#fdf4ff,#f5f3ff) !important;
}
.insight-card {
    border-radius:14px; border:2px solid #f0abfc;
    background:linear-gradient(135deg,#fdf4ff 0%,#fff0f9 100%);
    box-shadow:0 4px 20px rgba(168,85,247,0.12);
}
.insight-item {
    display:flex; align-items:center; gap:10px;
    padding:10px 14px; border-radius:10px; margin-bottom:8px;
    background:rgba(255,255,255,0.7); border:1px solid #f3e8ff;
    font-size:13px; font-weight:500; color:#4c1d95;
}
.ai-card {
    border-radius:14px; border:2px solid #fbcfe8;
    background:linear-gradient(135deg,#fff0f9,#fdf4ff);
    box-shadow:0 4px 20px rgba(219,39,119,0.10);
}
.drill-card {
    border-radius:14px; border:2px dashed #d8b4fe;
    background:#faf5ff; min-height:200px;
}
.reset-btn {
    background:linear-gradient(135deg,#7c3aed,#db2777) !important;
    border:none !important; color:#fff !important;
    font-weight:600 !important; border-radius:8px !important;
    font-size:13px !important;
}
.reset-btn:hover { opacity:0.88 !important; }
.dash-table-container { border-radius:10px; overflow:hidden; }
"""

# ─────────────────────────────────────────────────────────────
# DATA LOADING  (lru_cache = load once, reuse always)
# ─────────────────────────────────────────────────────────────

def load_data(path):
    xl          = pd.ExcelFile(path)
    sheet_names = xl.sheet_names
    target      = "Attendance_Report" if "Attendance_Report" in sheet_names else sheet_names[0]
    print(f"  Using sheet: '{target}'  (available: {sheet_names})")
    raw          = pd.read_excel(path, sheet_name=target, header=None)
    current_date = None
    records      = []

    for _, row in raw.iterrows():
        cell = str(row.iloc[0]).strip()
        if cell.startswith("Date"):
            try:    current_date = cell.split(":", 1)[1].strip()
            except: pass
            continue
        if str(row.iloc[0]).strip() in ("Sr No", "nan", "None") or pd.isna(row.iloc[1]):
            continue

        row_vals = list(row)
        fixed    = 4
        num_sess = (len(row_vals) - fixed) // 3
        rec = {
            "Date": current_date, "Sr_No": row_vals[0],
            "PNR_Number": str(row_vals[1]), "Name": str(row_vals[2]),
            "Department": str(row_vals[3]),
        }
        for i in range(num_sess):
            b = fixed + i * 3
            rec[f"S{i+1}_In"]  = str(row_vals[b])   if b   < len(row_vals) else "NA"
            rec[f"S{i+1}_Out"] = str(row_vals[b+1]) if b+1 < len(row_vals) else "NA"
            try:    rec[f"S{i+1}_Pct"] = float(str(row_vals[b+2]).replace("%","").strip())
            except: rec[f"S{i+1}_Pct"] = 0.0
        records.append(rec)

    df       = pd.DataFrame(records)
    num_sess = (len(df.columns) - 5) // 3

    def derive_status(row):
        pcts = [row.get(f"S{i+1}_Pct", 0) for i in range(num_sess)]
        ins  = [row.get(f"S{i+1}_In",  "NA") for i in range(num_sess)]
        if all(v in ("NA","nan","None") for v in ins): return "Absent"
        if all(p == 0 for p in pcts):                  return "Absent"
        if any(p == 0 for p in pcts):                  return "Partial absent"
        if any(0 < p < 75 for p in pcts):              return "Present – low %"
        return "Present"

    df["Status"]      = df.apply(derive_status, axis=1)
    df["Overall_Pct"] = df[[f"S{i+1}_Pct" for i in range(num_sess)]].mean(axis=1)
    return df, num_sess




# ─────────────────────────────────────────────────────────────
# AI RISK PREDICTION  (linear regression slope per student)
# ─────────────────────────────────────────────────────────────
def predict_at_risk(df, num_sess):
    results = []
    for name, grp in df.groupby("Name"):
        pcts = [grp[f"S{i+1}_Pct"].mean() for i in range(num_sess)]
        avg  = np.mean(pcts)
        slope = np.polyfit(range(len(pcts)), pcts, 1)[0] if len(pcts) > 1 else 0
        if avg < 75:
            risk = "🔴 Critical"
        elif avg < 85 and slope < -2:
            risk = "🟠 At Risk"
        elif avg < 85:
            risk = "🟡 Watch"
        else:
            risk = "🟢 Safe"
        results.append({"Name": name, "Avg %": round(avg,1),
                         "Trend (slope)": round(slope,2), "Prediction": risk})
    return pd.DataFrame(results).sort_values("Avg %")


# ─────────────────────────────────────────────────────────────
# CHART BUILDERS
# ─────────────────────────────────────────────────────────────
def _base(height=260):
    return dict(
        height=height, margin=dict(t=10,b=30,l=10,r=10),
        paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        font=dict(size=12, family="Poppins,sans-serif"),
    )

def chart_donut(df):
    counts = df["Status"].value_counts().reset_index()
    counts.columns = ["Status","Count"]
    cmap = {"Present":C["present"],"Absent":C["absent"],
            "Partial absent":C["partial"],"Present – low %":C["late"]}
    fig = go.Figure(go.Pie(
        labels=counts["Status"], values=counts["Count"], hole=0.62,
        marker=dict(colors=[cmap.get(s,C["muted"]) for s in counts["Status"]],
                    line=dict(color="#fff",width=2)),
        textinfo="percent",
        hovertemplate="%{label}: %{value}<extra></extra>",
    ))
    fig.update_layout(**_base(260), legend=dict(orientation="h",y=-0.18,font=dict(size=11)))
    return fig

def chart_trend(df):
    daily = (df.groupby("Date")
               .apply(lambda g: pd.Series(
                   {"pct": round((g["Status"]=="Present").sum()/len(g)*100,1)}))
               .reset_index())
    daily["rolling"] = daily["pct"].rolling(3, min_periods=1).mean().round(1)
    fig = go.Figure()
    fig.add_trace(go.Scatter(
        x=daily["Date"], y=daily["pct"], name="Daily %",
        mode="lines+markers+text",
        line=dict(color="#a855f7",width=2),
        marker=dict(size=7,color="#db2777",line=dict(color="#fff",width=2)),
        text=daily["pct"].astype(str)+"%", textposition="top center",
        fill="tozeroy", fillcolor="rgba(168,85,247,0.08)",
        hovertemplate="Date: %{x}<br>%{y:.1f}%<extra></extra>",
    ))
    fig.add_trace(go.Scatter(
        x=daily["Date"], y=daily["rolling"], name="3-day avg",
        mode="lines", line=dict(color="#f97316",width=2,dash="dot"),
        hovertemplate="3-day avg: %{y:.1f}%<extra></extra>",
    ))
    fig.update_layout(**_base(240),
        xaxis=dict(showgrid=False),
        yaxis=dict(range=[0,115],ticksuffix="%",showgrid=True,gridcolor=C["border"]),
        legend=dict(orientation="h",y=1.12))
    return fig

def chart_student_bar(df, num_sess):
    grp = df.groupby("Name").agg(
        **{f"S{i+1}":(f"S{i+1}_Pct","mean") for i in range(num_sess)}
    ).reset_index()
    grp["Overall"] = grp[[f"S{i+1}" for i in range(num_sess)]].mean(axis=1).round(1)
    grp = grp.sort_values("Overall", ascending=True)
    colors = ["#f43f5e" if v<75 else "#f97316" if v<85 else "#a855f7" for v in grp["Overall"]]
    fig = go.Figure(go.Bar(
        x=grp["Overall"], y=grp["Name"], orientation="h",
        marker_color=colors,
        text=grp["Overall"].astype(str)+"%", textposition="outside",
        hovertemplate="%{y}: %{x:.1f}%<extra></extra>",
    ))
    fig.update_layout(**_base(max(280,len(grp)*34)),
        xaxis=dict(range=[0,115],ticksuffix="%",showgrid=True,gridcolor=C["border"]),
        yaxis=dict(title=""))
    return fig

def chart_session_compare(df, num_sess):
    dates  = sorted(df["Date"].dropna().unique())
    colors = ["#7c3aed","#db2777","#f97316","#a855f7"]
    fig    = go.Figure()
    for i in range(num_sess):
        avgs = [round(df[df["Date"]==d][f"S{i+1}_Pct"].mean(),1) for d in dates]
        fig.add_trace(go.Bar(
            name=f"Session {i+1}", x=list(dates), y=avgs,
            marker_color=colors[i%len(colors)],
            text=[f"{v:.0f}%" for v in avgs], textposition="outside",
        ))
    fig.update_layout(**_base(260), barmode="group",
        yaxis=dict(range=[0,115],ticksuffix="%",showgrid=True,gridcolor=C["border"]),
        xaxis=dict(showgrid=False), legend=dict(orientation="h",y=1.12))
    return fig

def chart_dept(df):
    dept = (df.groupby("Department")
              .agg(avg=("Overall_Pct","mean"),count=("Name","count"))
              .reset_index())
    dept["avg"] = dept["avg"].round(1)
    colors = ["#f43f5e" if v<75 else "#f97316" if v<85 else "#a855f7" for v in dept["avg"]]
    fig = go.Figure(go.Bar(
        x=dept["Department"], y=dept["avg"],
        marker_color=colors,
        text=dept["avg"].astype(str)+"%", textposition="outside",
        hovertemplate="%{x}: %{y:.1f}%<extra></extra>",
    ))
    fig.update_layout(**_base(240),
        yaxis=dict(range=[0,115],ticksuffix="%",showgrid=True,gridcolor=C["border"]),
        xaxis=dict(showgrid=False))
    return fig

def chart_heatmap(df, num_sess):
    pivot = df.pivot_table(index="Name", columns="Date",
                           values="Overall_Pct", aggfunc="mean")
    fig = go.Figure(go.Heatmap(
        z=pivot.values, x=list(pivot.columns), y=list(pivot.index),
        colorscale=[[0,"#f43f5e"],[0.5,"#f97316"],[0.75,"#facc15"],[1,"#a855f7"]],
        zmin=0, zmax=100,
        hovertemplate="Student: %{y}<br>Date: %{x}<br>%{z:.1f}%<extra></extra>",
        colorbar=dict(ticksuffix="%", len=0.8),
    ))
    fig.update_layout(**_base(max(300,len(pivot)*26)),
        xaxis=dict(showgrid=False), yaxis=dict(showgrid=False))
    return fig

def chart_histogram(df):
    bins   = [0,50,75,85,101]
    labels = ["0–50%","50–75%","75–85%","85–100%"]
    colors = ["#f43f5e","#f97316","#facc15","#a855f7"]
    student_avg = df.groupby("Name")["Overall_Pct"].mean()
    counts = pd.cut(student_avg, bins=bins, labels=labels,
                    right=False).value_counts().reindex(labels)
    fig = go.Figure(go.Bar(
        x=labels, y=counts.values, marker_color=colors,
        text=counts.values, textposition="outside",
        hovertemplate="%{x}: %{y} students<extra></extra>",
    ))
    fig.update_layout(**_base(230),
        yaxis=dict(showgrid=True,gridcolor=C["border"],title="Students"),
        xaxis=dict(showgrid=False))
    return fig

def chart_top_bottom(df, top=True):
    avg = df.groupby("Name")["Overall_Pct"].mean().round(1).reset_index()
    avg.columns = ["Name","Overall_Pct"]
    data   = avg.nlargest(5,"Overall_Pct")  if top else avg.nsmallest(5,"Overall_Pct")
    colors = ["#a855f7"]*5                  if top else ["#f43f5e"]*5
    fig = go.Figure(go.Bar(
        x=data["Overall_Pct"], y=data["Name"], orientation="h",
        marker_color=colors,
        text=data["Overall_Pct"].astype(str)+"%", textposition="outside",
        hovertemplate="%{y}: %{x:.1f}%<extra></extra>",
    ))
    fig.update_layout(**_base(220),
        xaxis=dict(range=[0,115],ticksuffix="%",showgrid=True,gridcolor=C["border"]),
        yaxis=dict(title=""))
    return fig

def chart_drill(df, student_name, num_sess):
    sdf = df[df["Name"]==student_name]
    if sdf.empty:
        return go.Figure()
    colors = ["#7c3aed","#db2777","#f97316","#a855f7"]
    fig = go.Figure()
    for i in range(num_sess):
        fig.add_trace(go.Bar(
            name=f"Session {i+1}",
            x=list(sdf["Date"]), y=list(sdf[f"S{i+1}_Pct"]),
            text=[f"{v:.0f}%" for v in sdf[f"S{i+1}_Pct"]],
            textposition="outside",
            marker_color=colors[i%len(colors)],
        ))
    fig.update_layout(**_base(260), barmode="group",
        title=dict(text=f"📋 {student_name}",font=dict(size=13,color="#7c3aed"),x=0.01),
        yaxis=dict(range=[0,115],ticksuffix="%",showgrid=True,gridcolor=C["border"]),
        xaxis=dict(showgrid=False), legend=dict(orientation="h",y=1.15))
    return fig


# ─────────────────────────────────────────────────────────────
# ACTIONABLE INSIGHTS
# ─────────────────────────────────────────────────────────────
def build_insights(df, num_sess):
    low_count      = df[df["Overall_Pct"] < 75]["Name"].nunique()
    dept_avg       = df.groupby("Department")["Overall_Pct"].mean()
    lowest_dept    = dept_avg.idxmin()     if not dept_avg.empty else "N/A"
    lowest_dept_v  = dept_avg.min()        if not dept_avg.empty else 0
    best_dept      = dept_avg.idxmax()     if not dept_avg.empty else "N/A"
    best_dept_v    = dept_avg.max()        if not dept_avg.empty else 0
    sess_avgs      = {f"S{i+1}": df[f"S{i+1}_Pct"].mean() for i in range(num_sess)}
    worst_sess     = min(sess_avgs, key=sess_avgs.get) if sess_avgs else "N/A"
    worst_sess_v   = sess_avgs.get(worst_sess,0)
    absent_days    = df[df["Status"]=="Absent"].groupby("Name").size()
    chronic        = int((absent_days >= 2).sum()) if not absent_days.empty else 0
    avg_overall    = df["Overall_Pct"].mean()

    items = [
        ("⚠️", f"{low_count} student(s) below 75% attendance threshold", "#f43f5e"),
        ("🏢", f"Lowest dept: {lowest_dept} ({lowest_dept_v:.1f}%) — needs urgent attention", "#f97316"),
        ("🏆", f"Best dept: {best_dept} ({best_dept_v:.1f}%) — excellent performance!", "#22c55e"),
        ("📅", f"Weakest session: {worst_sess} (avg {worst_sess_v:.1f}%) — check scheduling", "#a855f7"),
        ("😔", f"{chronic} student(s) with chronic absences (2+ absent days)", "#f43f5e"),
        ("📊", f"Overall avg attendance across all records: {avg_overall:.1f}%", "#6366f1"),
    ]
    return [
        html.Div([
            html.Span(icon, style={"fontSize":"18px","minWidth":"24px"}),
            html.Span(text, style={"color":color,"fontWeight":"600"}),
        ], className="insight-item")
        for icon, text, color in items
    ]


# ─────────────────────────────────────────────────────────────
# KPI CARD
# ─────────────────────────────────────────────────────────────
def kpi(title, value, sub="", color=C["blue"], icon="📊"):
    return dbc.Card(dbc.CardBody([
        html.Div([
            html.Span(icon, style={"fontSize":"28px","lineHeight":"1"}),
            html.Div([
                html.P(title, style={"fontSize":"11px","color":"#9333ea","marginBottom":"2px",
                                     "fontWeight":"500","textTransform":"uppercase","letterSpacing":"0.5px"}),
                html.H3(value, style={"fontSize":"22px","fontWeight":"700","color":color,
                                      "marginBottom":"2px","lineHeight":"1"}),
                html.P(sub, style={"fontSize":"11px","color":"#a78bfa","marginBottom":0}),
            ], style={"marginLeft":"12px"}),
        ], style={"display":"flex","alignItems":"center"}),
    ]), className="kpi-card")


# ─────────────────────────────────────────────────────────────
# BUILD APP
# ─────────────────────────────────────────────────────────────
def build_app():
    app = Dash(__name__,
               external_stylesheets=[dbc.themes.BOOTSTRAP,
                   "https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&display=swap"],
               title="Attendance Management System")

    app.server.config['SEND_FILE_MAX_AGE_DEFAULT'] = 0
    app.index_string = app.index_string.replace("</head>", f"<style>{EXTRA_CSS}</style></head>")

    dates      = sorted(df["Date"].dropna().unique())
    depts      = sorted(df["Department"].dropna().unique())
    all_status = sorted(df["Status"].unique())

    total_students = df["PNR_Number"].nunique()
    avg_pct        = df["Overall_Pct"].mean()
    absent_count   = (df["Status"]=="Absent").sum()
    flagged_count  = df["Status"].str.contains("low|partial",case=False,na=False).sum()
    good_count     = (df["Overall_Pct"] >= 85).sum()
    avg_count      = ((df["Overall_Pct"] >= 75) & (df["Overall_Pct"] < 85)).sum()
    critical_count = (df["Overall_Pct"] < 75).sum()

    ai_df = predict_at_risk(df, num_sess)

    hdr_style = {"fontSize":"13px","fontWeight":"600","background":"transparent","border":"none"}

    def section(title, chart_id, icon="📈"):
        return dbc.Card([
            dbc.CardHeader(f"{icon}  {title}", style=hdr_style),
            dbc.CardBody(dcc.Loading(
                dcc.Graph(id=chart_id, config={"displayModeBar":False}),
                color="#a855f7", type="circle")),
        ], className="chart-card")

    app.layout = dbc.Container([

        # ── Hero ────────────────────────────────────────────
        html.Div([
            html.Div([
                html.Div("🎓", className="ams-icon-bubble"),
                html.Div("📋", className="ams-icon-bubble"),
                html.Div("✅", className="ams-icon-bubble"),
                html.Div("📊", className="ams-icon-bubble"),
                html.Div("🖥️", className="ams-icon-bubble"),
                html.Div("👥", className="ams-icon-bubble"),
                html.Div("🔔", className="ams-icon-bubble"),
            ], className="ams-icons-row"),
            html.H1("🎓 Attendance Management System 🎓", className="ams-title"),
            html.P("Smart · Automated · Insightful · AI-Powered", className="ams-subtitle"),
            html.Div([
                html.Span(
                    f"📁 {os.path.basename(EXCEL_FILE)}  ·  📅 {len(dates)} day(s)  ·  "
                    f"⏱ {num_sess} session(s)/day  ·  👨‍🎓 {total_students} students",
                    className="ams-badge")
            ], style={"textAlign":"center","marginTop":"14px"}),
            html.Div([
                html.Span("🎯 75% Attendance Threshold Active",
                          style={"display":"inline-block","background":"rgba(255,255,255,0.15)",
                                 "border":"1px solid rgba(255,255,255,0.25)","color":"#fff",
                                 "fontSize":"12px","fontWeight":"600","padding":"4px 16px",
                                 "borderRadius":"20px","marginTop":"10px"}),
            ], style={"textAlign":"center"}),
        ], className="ams-hero"),

        # ── KPI Row 1 ────────────────────────────────────────
        dbc.Row([
            dbc.Col(kpi("Total Students",  str(total_students), "unique PNRs",     "#7c3aed", "👨‍🎓"), md=3),
            dbc.Col(kpi("Avg Attendance",  f"{avg_pct:.1f}%",   "all sessions",    "#059669", "📈"), md=3),
            dbc.Col(kpi("Absent Records",  str(absent_count),   "fully absent",    "#f43f5e", "❌"), md=3),
            dbc.Col(kpi("Below 75%",       str(flagged_count),  "flagged records", "#f97316", "⚠️"), md=3),
        ], className="g-3", style={"marginBottom":"16px"}),

        # ── KPI Row 2: Category ──────────────────────────────
        dbc.Row([
            dbc.Col(kpi("🟢 Good (≥ 85%)",     str(good_count),     "records", "#22c55e", "🟢"), md=4),
            dbc.Col(kpi("🟠 Average (75–85%)",  str(avg_count),      "records", "#f97316", "🟠"), md=4),
            dbc.Col(kpi("🔴 Critical (< 75%)",  str(critical_count), "records", "#f43f5e", "🔴"), md=4),
        ], className="g-3", style={"marginBottom":"24px"}),

        # ── Actionable Insights ──────────────────────────────
        dbc.Card([
            dbc.CardHeader("⚡  Actionable Insights — Dashboard Intelligence",
                           style={**hdr_style,"color":"#7c3aed","fontSize":"14px"}),
            dbc.CardBody(html.Div(id="insights-panel")),
        ], className="insight-card", style={"marginBottom":"24px"}),

        # ── Filters ─────────────────────────────────────────
        dbc.Card(dbc.CardBody([
            html.Div([
                html.P("🔍  Filter & Search",
                       style={"fontSize":"13px","fontWeight":"700","color":"#7c3aed",
                              "marginBottom":"12px","display":"inline-block"}),
                dbc.Button("🔄 Reset Filters", id="btn-reset", size="sm",
                           className="reset-btn",
                           style={"float":"right","marginTop":"-4px"}),
            ]),
            dbc.Row([
                dbc.Col([
                    html.Label("📅 Date", style={"fontSize":"12px","color":"#9333ea","fontWeight":"600"}),
                    dcc.Dropdown(id="f-date",
                        options=[{"label":"📅 All dates","value":""}]+[{"label":d,"value":d} for d in dates],
                        value="", clearable=False, style={"fontSize":"13px"}),
                ], md=3),
                dbc.Col([
                    html.Label("🏢 Department", style={"fontSize":"12px","color":"#9333ea","fontWeight":"600"}),
                    dcc.Dropdown(id="f-dept",
                        options=[{"label":"🏢 All departments","value":""}]+[{"label":d,"value":d} for d in depts],
                        value="", clearable=False, style={"fontSize":"13px"}),
                ], md=3),
                dbc.Col([
                    html.Label("✅ Status", style={"fontSize":"12px","color":"#9333ea","fontWeight":"600"}),
                    dcc.Dropdown(id="f-status",
                        options=[{"label":"✅ All statuses","value":""}]+[{"label":s,"value":s} for s in all_status],
                        value="", clearable=False, style={"fontSize":"13px"}),
                ], md=3),
                dbc.Col([
                    html.Label("🔎 Search name / PNR", style={"fontSize":"12px","color":"#9333ea","fontWeight":"600"}),
                    dbc.Input(id="f-search", placeholder="Type to search…", size="sm",
                              style={"fontSize":"13px","border":"1.5px solid #e9d5ff","borderRadius":"8px"}),
                ], md=3),
            ], className="g-2"),
        ]), className="filter-card", style={"marginBottom":"24px"}),

        # ── Charts Row 1 ─────────────────────────────────────
        dbc.Row([
            dbc.Col(section("Daily Attendance Trend + 3-Day Rolling Avg", "ch-trend", "📈"), md=8),
            dbc.Col(section("Status Distribution",                         "ch-donut", "🍩"), md=4),
        ], className="g-3", style={"marginBottom":"24px"}),

        # ── Charts Row 2 ─────────────────────────────────────
        dbc.Row([
            dbc.Col(section("Session Comparison by Date",   "ch-session", "📅"), md=6),
            dbc.Col(section("Avg Attendance by Department", "ch-dept",    "🏢"), md=6),
        ], className="g-3", style={"marginBottom":"24px"}),

        # ── Heatmap ──────────────────────────────────────────
        dbc.Card([
            dbc.CardHeader("🔥  Attendance Heatmap  (Student × Date)", style=hdr_style),
            dbc.CardBody(dcc.Loading(
                dcc.Graph(id="ch-heatmap", config={"displayModeBar":False}),
                color="#a855f7", type="circle")),
        ], className="chart-card", style={"marginBottom":"24px"}),

        # ── Histogram + Top/Bottom ────────────────────────────
        dbc.Row([
            dbc.Col(section("Attendance Distribution", "ch-hist",   "📊"), md=4),
            dbc.Col(section("🏆 Top 5 Students",       "ch-top",    "🟢"), md=4),
            dbc.Col(section("📉 Bottom 5 Students",    "ch-bottom", "🔴"), md=4),
        ], className="g-3", style={"marginBottom":"24px"}),

        # ── Student Bar ──────────────────────────────────────
        dbc.Card([
            dbc.CardHeader(
                "👨‍🎓  Attendance % per Student  —  🔴 < 75%  |  🟠 < 85%  |  🟣 ≥ 85%"
                "   ·  Click a bar to drill down ↓",
                style=hdr_style),
            dbc.CardBody(dcc.Loading(
                dcc.Graph(id="ch-bar", config={"displayModeBar":False}),
                color="#a855f7", type="circle")),
        ], className="chart-card", style={"marginBottom":"24px"}),

        # ── Drill-down ───────────────────────────────────────
        dbc.Card([
            dbc.CardHeader("🔍  Student Drill-Down  (click a bar above to explore)",
                           style={**hdr_style,"color":"#db2777"}),
            dbc.CardBody(dcc.Loading(
                html.Div(id="drill-panel"), color="#db2777", type="circle")),
        ], className="drill-card", style={"marginBottom":"24px"}),

        # ── AI Prediction ────────────────────────────────────
        dbc.Card([
            dbc.CardHeader("🤖  AI Risk Prediction  — Who may fall below 75%?",
                           style={**hdr_style,"color":"#db2777","fontSize":"14px"}),
            dbc.CardBody([
                html.P(
                    "Prediction uses linear regression on session-wise percentages. "
                    "Negative slope + avg near threshold = At Risk.",
                    style={"fontSize":"12px","color":"#9333ea","marginBottom":"12px","fontStyle":"italic"}),
                dcc.Loading(
                    dash_table.DataTable(
                        id="tbl-ai",
                        data=ai_df.to_dict("records"),
                        columns=[{"name":c,"id":c} for c in ai_df.columns],
                        style_table={"overflowX":"auto"},
                        style_header={"background":"#fdf4ff","fontWeight":"700","fontSize":"12px",
                                      "color":"#7c3aed","border":"1px solid #f3e8ff","padding":"8px 10px"},
                        style_cell={"fontSize":"12px","padding":"7px 10px",
                                    "border":"1px solid #f3e8ff","fontFamily":"Poppins,sans-serif"},
                        style_data_conditional=[
                            {"if":{"filter_query":"{Prediction} contains 'Critical'"},"color":"#f43f5e","fontWeight":"700"},
                            {"if":{"filter_query":"{Prediction} contains 'At Risk'"},"color":"#f97316","fontWeight":"600"},
                            {"if":{"filter_query":"{Prediction} contains 'Watch'"},  "color":"#ca8a04","fontWeight":"500"},
                            {"if":{"filter_query":"{Prediction} contains 'Safe'"},   "color":"#22c55e","fontWeight":"500"},
                            {"if":{"row_index":"odd"},"backgroundColor":"#faf5ff"},
                        ],
                        page_size=10, sort_action="native",
                    ), color="#db2777", type="circle"),
            ]),
        ], className="ai-card", style={"marginBottom":"24px"}),

        # ── Detail Table ─────────────────────────────────────
        dbc.Card([
            dbc.CardHeader(dbc.Row([
                dbc.Col("📋  Student Detail", md=5,
                        style={"fontSize":"13px","fontWeight":"700","color":"#7c3aed","alignSelf":"center"}),
                dbc.Col(html.Div(id="rec-count",
                                 style={"textAlign":"center","fontSize":"12px",
                                        "color":"#9333ea","fontWeight":"600","alignSelf":"center"}), md=4),
                dbc.Col(html.Div([
                    dcc.Download(id="download-data"),
                    dbc.Button("⬇️ Export CSV", id="btn-export", size="sm",
                               style={"background":"linear-gradient(135deg,#7c3aed,#db2777)",
                                      "border":"none","color":"#fff","fontWeight":"600",
                                      "borderRadius":"8px","fontSize":"12px"}),
                ], style={"textAlign":"right"}), md=3),
            ]), style={"background":"linear-gradient(90deg,#f5f3ff,#fdf4ff)",
                       "border":"none","borderRadius":"14px 14px 0 0"}),
            dbc.CardBody(dcc.Loading(
                dash_table.DataTable(
                    id="tbl",
                    style_table={"overflowX":"auto"},
                    style_header={"background":"#f5f3ff","fontWeight":"700","fontSize":"12px",
                                  "color":"#7c3aed","border":"1px solid #e9d5ff","padding":"8px 10px"},
                    style_cell={"fontSize":"12px","padding":"7px 10px","fontFamily":"Poppins,sans-serif",
                                "border":"1px solid #f3e8ff","maxWidth":"160px",
                                "overflow":"hidden","textOverflow":"ellipsis"},
                    style_data_conditional=[
                        {"if":{"row_index":"odd"},"backgroundColor":"#faf5ff"},
                        {"if":{"filter_query":"{Status} = 'Absent'"},       "color":"#f43f5e","fontWeight":"600"},
                        {"if":{"filter_query":"{Status} contains 'low'"},   "color":"#f97316","fontWeight":"600"},
                        {"if":{"filter_query":"{Status} contains 'Partial'"},"color":"#a855f7","fontWeight":"600"},
                    ],
                    page_size=20, sort_action="native",
                    filter_action="native", export_format="xlsx",
                ), color="#a855f7", type="circle")),
        ], className="chart-card", style={"marginBottom":"40px"}),

        # ── Footer ───────────────────────────────────────────
        html.Div([
            html.Hr(style={"borderColor":"#e9d5ff","marginBottom":"16px"}),
            html.Div([
                html.P("🎓 Attendance Management System  ·  Powered by Dash & Plotly",
                       style={"textAlign":"center","color":"#a855f7","fontSize":"12px",
                              "fontWeight":"500","marginBottom":"6px"}),
                html.P([
                    html.Span("© 2025 ", style={"color":"#db2777"}),
                    html.Span("Madhuri Nimbalkar", style={"color":"#7c3aed","fontWeight":"700","fontSize":"13px"}),
                    html.Span("  ·  All Rights Reserved", style={"color":"#a855f7"}),
                    html.Span("  🔒", style={"fontSize":"13px"}),
                ], style={"textAlign":"center","marginBottom":"4px","fontSize":"12px"}),
                html.P("Unauthorized use, reproduction or distribution is strictly prohibited.",
                       style={"textAlign":"center","color":"#c084fc","fontSize":"11px",
                              "marginBottom":"4px","fontStyle":"italic"}),
            ], style={"background":"linear-gradient(135deg,#fdf4ff,#f5f3ff)",
                      "border":"1.5px solid #e9d5ff","borderRadius":"12px",
                      "padding":"14px 20px","marginBottom":"30px"}),
        ]),

    ], fluid=True, style={
        "fontFamily":"'Poppins',sans-serif",
        "background":"linear-gradient(160deg,#fdf4ff 0%,#faf5ff 50%,#eff6ff 100%)",
        "minHeight":"100vh","padding":"24px",
    })

    # ── CALLBACK: Reset filters ──────────────────────────────
    @app.callback(
        Output("f-date",   "value"),
        Output("f-dept",   "value"),
        Output("f-status", "value"),
        Output("f-search", "value"),
        Input("btn-reset", "n_clicks"),
        prevent_initial_call=True,
    )
    def reset_filters(_):
        return "", "", "", ""

    # ── CALLBACK: Export CSV ─────────────────────────────────
    @app.callback(
        Output("download-data","data"),
        Input("btn-export","n_clicks"),
        State("tbl","data"),
        prevent_initial_call=True,
    )
    def export_csv(n, tbl_data):
        if not tbl_data:
            return None
        return dcc.send_data_frame(
            pd.DataFrame(tbl_data).to_csv,
            "attendance_export.csv", index=False)

    # ── CALLBACK: Drill-down ─────────────────────────────────
    @app.callback(
        Output("drill-panel","children"),
        Input("ch-bar","clickData"),
    )
    def drill_down(click_data):

        if not click_data:
            return html.P(
                "👆 Click any student bar above to see their session-wise daily breakdown.",
                style={"color": "#a78bfa", "textAlign": "center", "padding": "40px", "fontSize": "13px"}
            )

        try:
            student = click_data["points"][0]["y"]

        except Exception:
            return html.P("Could not parse selection.", style={"color": "#f43f5e"})

        # 🔥 LOAD DATA OUTSIDE EXCEPT (IMPORTANT)
        df, num_sess = load_data(EXCEL_FILE)

        return dcc.Graph(
            figure=chart_drill(df, student, num_sess),
            config={"displayModeBar": False}
        )


    # ── CALLBACK: Main charts + table + insights ─────────────
    @app.callback(
        Output("ch-trend",       "figure"),
        Output("ch-donut",       "figure"),
        Output("ch-session",     "figure"),
        Output("ch-dept",        "figure"),
        Output("ch-bar",         "figure"),
        Output("ch-heatmap",     "figure"),
        Output("ch-hist",        "figure"),
        Output("ch-top",         "figure"),
        Output("ch-bottom",      "figure"),
        Output("tbl",            "data"),
        Output("tbl",            "columns"),
        Output("rec-count",      "children"),
        Output("insights-panel", "children"),
        Input("f-date",   "value"),
        Input("f-dept",   "value"),
        Input("f-status", "value"),
        Input("f-search", "value"),
    )
    def update(f_date, f_dept, f_status, search):
        df, num_sess = load_data(EXCEL_FILE)  # 🔥 fresh load
        dff = df.copy()
        if f_date:   dff = dff[dff["Date"] == f_date]
        if f_dept:   dff = dff[dff["Department"] == f_dept]
        if f_status: dff = dff[dff["Status"] == f_status]
        if search:
            q   = search.strip().lower()
            dff = dff[dff["Name"].str.lower().str.contains(q, na=False) |
                      dff["PNR_Number"].str.lower().str.contains(q, na=False)]

        fixed  = ["Date","Sr_No","PNR_Number","Name","Department","Status","Overall_Pct"]
        s_cols = [c for i in range(num_sess) for c in [f"S{i+1}_In",f"S{i+1}_Out",f"S{i+1}_Pct"]]
        avail  = [c for c in fixed+s_cols if c in dff.columns]
        tbl_df = dff[avail].copy()
        tbl_df["Overall_Pct"] = tbl_df["Overall_Pct"].round(1).astype(str)+"%"
        for i in range(num_sess):
            col = f"S{i+1}_Pct"
            if col in tbl_df.columns:
                tbl_df[col] = tbl_df[col].round(1).astype(str)+"%"
        cols = [{"name":c.replace("_"," "),"id":c} for c in avail]

        return (
            chart_trend(dff),
            chart_donut(dff),
            chart_session_compare(dff, num_sess),
            chart_dept(dff),
            chart_student_bar(dff, num_sess),
            chart_heatmap(dff, num_sess),
            chart_histogram(dff),
            chart_top_bottom(dff, top=True),
            chart_top_bottom(dff, top=False),
            tbl_df.to_dict("records"),
            cols,
            f"📊 {len(dff):,} records",
            build_insights(dff, num_sess),
        )

    return app


# ─────────────────────────────────────────────────────────────
# AUTO-OPEN BROWSER
# ─────────────────────────────────────────────────────────────
def open_browser():
    import urllib.request
    url = f"http://localhost:{PORT}"
    for _ in range(20):
        try:
            urllib.request.urlopen(url, timeout=0.5)
            webbrowser.open(url)
            return
        except Exception:
            time.sleep(0.5)

# ─────────────────────────────────────────────────────────────
# ENTRY POINT
# ─────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print(f"\n  Loading {EXCEL_FILE} …")

    # 🔥 Only for initial logging (NOT used in app)
    df, num_sess = load_data(EXCEL_FILE)

    print(f"  {len(df)} records  ·  {df['PNR_Number'].nunique()} students  "
          f"·  {df['Date'].nunique()} dates  ·  {num_sess} sessions/day")

    print(f"\n  Dashboard → http://localhost:{PORT}\n")

    threading.Thread(target=open_browser, daemon=True).start()

    # 🔥 DO NOT PASS df (preloaded data)
    app = build_app()   # ✅ changed

    app.run(debug=False, host="0.0.0.0", port=PORT, use_reloader=False)