"""
WhatsApp Campaign ROI Dashboard - v4
Fixes: net revenue, professional email with Excel attachment, better optional column detection
"""

import os, re, smtplib, io, base64
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

import pandas as pd
import plotly.express as px
import plotly.io as pio
import streamlit as st

st.set_page_config(page_title="WhatsApp Campaign ROI Dashboard", page_icon="📊", layout="wide")

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
html, body, [class*="css"] { font-family: 'Inter', sans-serif; }
.stApp { background-color: #F8FAFC; }
[data-testid="stSidebar"] { background: #FFFFFF; border-right: 1px solid #E2E8F0; }
.kpi-card { background:#FFFFFF; border-radius:12px; padding:20px 24px;
    box-shadow:0 1px 4px rgba(0,0,0,0.07); border:1px solid #E2E8F0; text-align:center; margin-bottom:8px; }
.kpi-label { font-size:12px; font-weight:600; color:#64748B; letter-spacing:0.05em;
    text-transform:uppercase; margin-bottom:6px; }
.kpi-value { font-size:30px; font-weight:700; color:#0F172A; }
.kpi-icon { font-size:24px; margin-bottom:8px; }
.section-header { font-size:17px; font-weight:700; color:#0F172A; margin-top:28px; margin-bottom:10px; }
.stDownloadButton > button, .stButton > button {
    background-color:#2563EB !important; color:#FFFFFF !important;
    border-radius:8px !important; font-weight:600 !important; border:none !important; }
.dash-title { font-size:26px; font-weight:800; color:#0F172A; margin-bottom:2px; }
.dash-subtitle { font-size:13px; color:#64748B; margin-bottom:20px; }
</style>""", unsafe_allow_html=True)

MAX_BYTES = 25 * 1024 * 1024

# Columns that are "optional add-ons" — detected automatically if present in sheet
# Net revenue deduction columns (numeric, will be subtracted from revenue)
DEDUCTION_COLS = ["cpshare", "IHC", "tutorShare", "razorpayShare", "platformFee"]
# Info-only columns (shown in table but not deducted)
INFO_COLS      = ["transferAmount", "expiryDate", "transactionDate", "state", "email"]
ALL_OPTIONAL   = DEDUCTION_COLS + INFO_COLS

# ── Helpers ───────────────────────────────────────────────────────────────────

def normalize_phone(phone) -> str:
    if pd.isna(phone): return ""
    digits = re.sub(r"\D", "", str(phone))
    return digits[-10:] if len(digits) >= 10 else digits

def load_file(f):
    try:
        if f.name.lower().endswith(".csv"):
            return pd.read_csv(f)
        return pd.read_excel(f, engine="openpyxl")
    except Exception as e:
        st.error(f"Error reading file: {e}"); return None

def check_size(f):
    f.seek(0, 2); sz = f.tell(); f.seek(0)
    if sz > MAX_BYTES:
        st.error("❌ Upload Blocked: File exceeds 25 MB limit."); return False
    return True

def fmt_inr(v):
    try:
        p = str(int(round(float(v))))
        if len(p) > 3:
            last3, rest, groups = p[-3:], p[:-3], []
            while len(rest) > 2: groups.insert(0, rest[-2:]); rest = rest[:-2]
            if rest: groups.insert(0, rest)
            return f"₹{','.join(groups)},{last3}"
        return f"₹{p}"
    except: return f"₹{v:,.0f}"

def detect_mobile_col(df):
    best_col, best_count = None, 0
    for col in df.columns:
        count = df[col].apply(lambda x: len(re.sub(r"\D","",str(x))) >= 10).sum()
        if count > best_count: best_count, best_col = count, col
    return best_col

def find_col(df_cols, candidates):
    """Case-insensitive column finder."""
    lower_map = {c.lower().strip(): c for c in df_cols}
    for c in candidates:
        if c.lower() in lower_map:
            return lower_map[c.lower()]
    return None

def make_excel_bytes(df):
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Verified Conversions")
    return buf.getvalue()

def make_chart_image(df_matched):
    """Return base64 PNG of the course chart for email embedding."""
    try:
        course_stats = (df_matched.groupby("CourseName")
            .agg(Revenue=("Amount_Paid","sum"), Students=("mobile","nunique"))
            .reset_index().sort_values("Revenue", ascending=True))
        fig = px.bar(course_stats, x="Revenue", y="CourseName", orientation="h",
            title="Course Sales Performance",
            text=course_stats["Students"].apply(lambda s: f"👤 {s}"),
            color="Revenue", color_continuous_scale=[[0,"#93C5FD"],[1,"#2563EB"]],
            labels={"Revenue":"Revenue (₹)","CourseName":"Course"})
        fig.update_traces(textposition="outside", marker_line_width=0)
        fig.update_layout(plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
            coloraxis_showscale=False, height=max(300, len(course_stats)*52+80),
            margin=dict(l=20,r=40,t=50,b=20))
        img_bytes = pio.to_image(fig, format="png", width=700, height=max(300, len(course_stats)*52+80))
        return base64.b64encode(img_bytes).decode()
    except:
        return None

def build_insights(df_matched, total_revenue, total_conversions, roi, campaign_cost):
    """Generate smart bullet-point insights from the data."""
    course_stats = (df_matched.groupby("CourseName")
        .agg(Revenue=("Amount_Paid","sum"), Enrollments=("mobile","nunique"))
        .reset_index())

    top_rev   = course_stats.loc[course_stats["Revenue"].idxmax()]
    top_enrol = course_stats.loc[course_stats["Enrollments"].idxmax()]
    avg_order = total_revenue / total_conversions if total_conversions else 0
    conv_rate = (total_conversions / df_matched["mobile"].nunique() * 100) if total_conversions else 0

    insights = [
        f"🏆 <b>Highest Revenue Course:</b> {top_rev['CourseName']} — {fmt_inr(top_rev['Revenue'])} from {int(top_rev['Enrollments'])} students",
        f"👥 <b>Most Enrolled Course:</b> {top_enrol['CourseName']} — {int(top_enrol['Enrollments'])} enrollments",
        f"💳 <b>Average Order Value:</b> {fmt_inr(avg_order)} per conversion",
        f"✅ <b>Total Verified Conversions:</b> {total_conversions:,} students",
    ]
    return insights


def send_report_email(recipient, total_revenue, roi, total_conversions,
                      campaign_cost, total_deductions, net_revenue,
                      df, df_matched, include_net_revenue=True):
    try:
        sender   = st.secrets.get("EMAIL_SENDER")   or os.environ.get("EMAIL_SENDER","")
        password = st.secrets.get("EMAIL_PASSWORD") or os.environ.get("EMAIL_PASSWORD","")
        if not sender or not password:
            return False, "❌ Email credentials not set in secrets.toml"

        # Chart image
        chart_b64   = make_chart_image(df_matched)
        chart_html  = f'<img src="data:image/png;base64,{chart_b64}" style="width:100%;border-radius:10px;margin:8px 0 20px;" />' if chart_b64 else "<p style='color:#94A3B8;font-size:13px;'>Chart not available.</p>"

        # Smart insights
        insights     = build_insights(df_matched, total_revenue, total_conversions, roi, campaign_cost)
        insights_html = "".join(f'<li style="margin-bottom:8px;">{i}</li>' for i in insights)

        # Net revenue block
        nr_color = "#16A34A" if net_revenue >= 0 else "#DC2626"
        net_rev_kpi = f'<div class="kpi"><div class="icon">🟢</div><div class="val" style="color:{nr_color};">{fmt_inr(net_revenue)}</div><div class="lbl">Net Revenue</div></div>' if include_net_revenue else ""
        breakdown_html = f"""
        <div class="section-title">📐 Revenue Breakdown</div>
        <div class="breakdown"><table>
          <tr><td>💰 Gross Revenue</td><td>{fmt_inr(total_revenue)}</td></tr>
          <tr><td>📡 Broadcast Cost</td><td style="color:#DC2626;">− {fmt_inr(campaign_cost)}</td></tr>
          <tr><td>🔻 Platform Deductions</td><td style="color:#DC2626;">− {fmt_inr(total_deductions)}</td></tr>
          <tr style="border-top:1px solid #CBD5E1;font-weight:700;">
            <td>🟢 Net Revenue</td><td style="color:{nr_color};">{fmt_inr(net_revenue)}</td></tr>
        </table></div>""" if include_net_revenue else ""

        roi_color = "#16A34A" if roi >= 0 else "#DC2626"

        html = f"""<!DOCTYPE html>
<html><head>
<link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;600;700&display=swap" rel="stylesheet">
<style>
  body {{ margin:0; padding:0; background:#F1F5F9; font-family:'Inter',Arial,sans-serif; color:#0F172A; }}
  .wrap {{ max-width:680px; margin:32px auto; background:#fff; border-radius:16px;
           padding:36px 40px; box-shadow:0 4px 24px rgba(0,0,0,0.08); }}
  .logo {{ font-size:22px; font-weight:800; color:#2563EB; margin-bottom:2px; }}
  .sub  {{ font-size:13px; color:#64748B; margin-bottom:20px; }}
  hr    {{ border:none; border-top:1px solid #E2E8F0; margin:18px 0; }}
  .kpi-row {{ display:flex; gap:10px; margin:18px 0; }}
  .kpi {{ flex:1; background:#F8FAFC; border:1px solid #E2E8F0; border-radius:10px; padding:14px 8px; text-align:center; }}
  .kpi .icon {{ font-size:18px; }}
  .kpi .val  {{ font-size:19px; font-weight:700; color:#0F172A; margin:4px 0; }}
  .kpi .lbl  {{ font-size:10px; color:#64748B; text-transform:uppercase; letter-spacing:0.04em; }}
  .section-title {{ font-size:15px; font-weight:700; color:#0F172A; margin:22px 0 10px; }}
  .insight-box {{ background:#F0F9FF; border-left:4px solid #2563EB; border-radius:8px;
                  padding:16px 20px; margin-bottom:20px; }}
  .insight-box ul {{ margin:0; padding-left:18px; font-size:13px; line-height:1.8; color:#1E3A5F; }}
  .breakdown {{ background:#F8FAFC; border-radius:8px; padding:14px 18px; font-size:13px; }}
  .breakdown table {{ width:100%; border-collapse:collapse; }}
  .breakdown td {{ padding:6px 0; }}
  .breakdown td:last-child {{ text-align:right; font-weight:600; }}
  .excel-note {{ background:#F0FDF4; border:1px solid #86EFAC; border-radius:8px;
                 padding:12px 16px; font-size:13px; color:#166534; margin-top:20px; }}
  .footer {{ font-size:11px; color:#94A3B8; margin-top:28px; text-align:center; }}
</style></head><body>
<div class="wrap">
  <div class="logo">📊 WhatsApp Campaign ROI Report</div>
  <div class="sub">Auto-generated · {pd.Timestamp.now().strftime("%d %b %Y, %I:%M %p")}</div>
  <hr>

  <!-- KPI Row -->
  <div class="kpi-row">
    <div class="kpi"><div class="icon">💰</div><div class="val">{fmt_inr(total_revenue)}</div><div class="lbl">Total Revenue</div></div>
    <div class="kpi"><div class="icon">📈</div><div class="val" style="color:{roi_color};">{roi:.2f}%</div><div class="lbl">ROI</div></div>
    <div class="kpi"><div class="icon">✅</div><div class="val">{total_conversions:,}</div><div class="lbl">Conversions</div></div>
    {net_rev_kpi}
  </div>

  <!-- Smart Insights -->
  <div class="section-title">💡 Campaign Insights</div>
  <div class="insight-box">
    <ul>{insights_html}</ul>
  </div>

  {breakdown_html}

  <!-- Chart -->
  <div class="section-title">📊 Course Sales Performance</div>
  {chart_html}

  <!-- Excel note -->
  <div class="excel-note">
    📎 <strong>Full student list attached</strong> as Excel file (<em>verified_conversions.xlsx</em>).
    It contains all {total_conversions} verified conversions with name, mobile, course and amount details.
  </div>

  <div class="footer">Sent by WhatsApp Campaign ROI Dashboard</div>
</div>
</body></html>"""

        msg = MIMEMultipart("mixed")
        msg["Subject"] = f"📊 WhatsApp ROI — {total_conversions} Conversions | {fmt_inr(total_revenue)} | ROI {roi:.1f}%"
        msg["From"] = "WhatsApp ROI Dashboard <noreply@classplus.co>"
        msg["To"]   = recipient

        alt = MIMEMultipart("alternative")
        alt.attach(MIMEText(html, "html"))
        msg.attach(alt)

        # Excel attachment (student details — NOT in email body)
        excel_bytes = make_excel_bytes(df)
        part = MIMEBase("application", "octet-stream")
        part.set_payload(excel_bytes)
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", 'attachment; filename="verified_conversions.xlsx"')
        msg.attach(part)

        with smtplib.SMTP("smtp.gmail.com", 587) as s:
            s.ehlo(); s.starttls(); s.login(sender, password)
            s.sendmail(sender, recipient, msg.as_string())
        return True, "✅ Report sent! Student list attached as Excel."
    except smtplib.SMTPAuthenticationError:
        return False, "❌ Authentication failed. Check EMAIL_SENDER and EMAIL_PASSWORD."
    except Exception as e:
        return False, f"❌ Error: {e}"

# ── Sidebar ───────────────────────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### ⚙️ Campaign Settings")
    st.markdown("---")
    broadcast_cost = st.number_input(
        "📡 Broadcast Cost per Message (₹)",
        min_value=0.0, value=1.18, step=0.01, format="%.2f"
    )
    st.markdown("---")
    st.caption("Net Revenue = Total Revenue − Broadcast Cost − Platform Deductions (cpshare, IHC etc. from your sheet)")

# ── Header ────────────────────────────────────────────────────────────────────
st.markdown("""
<div class="dash-title">📊 WhatsApp Campaign ROI Dashboard</div>
<div class="dash-subtitle">Classplus Internal Tool — Measure the ROI of your WhatsApp campaigns.</div>
""", unsafe_allow_html=True)

# ── File Uploads ──────────────────────────────────────────────────────────────
col_a, col_b = st.columns(2)
with col_a:
    st.markdown('<div class="section-header">📋 Block A – Broadcast List</div>', unsafe_allow_html=True)
    st.caption("Mobile numbers — any column, auto-detected")
    file_a = st.file_uploader("Upload Broadcast List", type=["csv","xlsx"], key="broadcast", label_visibility="collapsed")
with col_b:
    st.markdown('<div class="section-header">🧾 Block B – Sales Transactions</div>', unsafe_allow_html=True)
    st.caption("Required: **name, mobile, CourseName, Amount_Paid**")
    file_b = st.file_uploader("Upload Sales Transactions", type=["csv","xlsx"], key="sales", label_visibility="collapsed")

if file_a is None or file_b is None:
    st.info("⬆️ Upload both files to generate the dashboard.", icon="ℹ️")
    st.stop()

if not check_size(file_a) or not check_size(file_b): st.stop()

df_a = load_file(file_a)
df_b = load_file(file_b)
if df_a is None or df_b is None: st.stop()
if df_a.empty: st.error("❌ Block A is empty."); st.stop()
if df_b.empty: st.error("❌ Block B is empty."); st.stop()

# ── Block A: extract only mobile numbers ──────────────────────────────────────
mobile_col_a = detect_mobile_col(df_a)
if not mobile_col_a:
    st.error("❌ Could not find a mobile number column in Block A."); st.stop()
df_a = df_a[[mobile_col_a]].copy()
df_a.columns = ["mobile"]
df_a["mobile"] = df_a["mobile"].apply(normalize_phone)
df_a = df_a[df_a["mobile"].str.len() == 10].drop_duplicates()

# ── Block B: strip column names, validate required ────────────────────────────
df_b.columns = [c.strip() for c in df_b.columns]

# Case-insensitive required column detection
col_name   = find_col(df_b.columns, ["name","Name","NAME","student_name"])
col_mobile = find_col(df_b.columns, ["mobile","Mobile","MOBILE","phone","Phone","number"])
col_course = find_col(df_b.columns, ["CourseName","coursename","course_name","Course","course"])
col_amount = find_col(df_b.columns, ["Amount_Paid","amount_paid","AmountPaid","amount","Amount","fees"])

missing = []
if not col_name:   missing.append("name")
if not col_mobile: missing.append("mobile")
if not col_course: missing.append("CourseName")
if not col_amount: missing.append("Amount_Paid")

if missing:
    st.error(f"❌ Block B is missing columns: **{', '.join(missing)}**"); st.stop()

# Standardise required column names
df_b = df_b.rename(columns={
    col_name: "name", col_mobile: "mobile",
    col_course: "CourseName", col_amount: "Amount_Paid"
})

df_b["mobile"] = df_b["mobile"].apply(normalize_phone)
df_b = df_b[df_b["mobile"].str.len() == 10]
df_b["Amount_Paid"] = pd.to_numeric(df_b["Amount_Paid"], errors="coerce").fillna(0)

# Detect optional columns (case-insensitive) — include ALL extra columns not in required set
required_set = {"name","mobile","CourseName","Amount_Paid"}
extra_cols = [c for c in df_b.columns if c not in required_set]

# ── Match ─────────────────────────────────────────────────────────────────────
keep_b = ["name","mobile","CourseName","Amount_Paid"] + extra_cols
df_matched = pd.merge(df_a[["mobile"]], df_b[keep_b], on="mobile", how="inner")

if df_matched.empty:
    st.warning("⚠️ No matching records found between Block A and Block B."); st.stop()

# ── Course filter ─────────────────────────────────────────────────────────────
all_courses = sorted(df_matched["CourseName"].unique().tolist())
selected_courses = st.multiselect("🎓 Filter by Course (leave blank for all)", all_courses)
if selected_courses:
    df_matched = df_matched[df_matched["CourseName"].isin(selected_courses)]

# ── KPI calculations ──────────────────────────────────────────────────────────
total_revenue     = df_matched["Amount_Paid"].sum()
total_conversions = len(df_matched)
campaign_cost     = df_a["mobile"].nunique() * broadcast_cost

# Deductions: sum all numeric columns that are deduction-type
total_deductions = 0.0
found_deduction_cols = [c for c in df_matched.columns if c.lower() in [d.lower() for d in DEDUCTION_COLS]]
for c in found_deduction_cols:
    total_deductions += pd.to_numeric(df_matched[c], errors="coerce").fillna(0).sum()

net_revenue = total_revenue - campaign_cost - total_deductions
roi = ((total_revenue - campaign_cost) / campaign_cost * 100) if campaign_cost > 0 else 0.0

# ── KPI Cards ─────────────────────────────────────────────────────────────────
st.markdown("---")
k1, k2, k3, k4 = st.columns(4)

with k1:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-icon">💰</div><div class="kpi-label">Total Revenue</div>
        <div class="kpi-value">{fmt_inr(total_revenue)}</div></div>""", unsafe_allow_html=True)
with k2:
    color = "#22C55E" if roi >= 0 else "#EF4444"
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-icon">📈</div><div class="kpi-label">ROI</div>
        <div class="kpi-value" style="color:{color};">{roi:.2f}%</div></div>""", unsafe_allow_html=True)
with k3:
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-icon">✅</div><div class="kpi-label">Total Conversions</div>
        <div class="kpi-value">{total_conversions:,}</div></div>""", unsafe_allow_html=True)
with k4:
    nc = "#22C55E" if net_revenue >= 0 else "#EF4444"
    st.markdown(f"""<div class="kpi-card">
        <div class="kpi-icon">🟢</div><div class="kpi-label">Net Revenue</div>
        <div class="kpi-value" style="color:{nc};">{fmt_inr(net_revenue)}</div></div>""", unsafe_allow_html=True)

# Net revenue breakdown expander
with st.expander("📐 Net Revenue Breakdown"):
    st.markdown(f"""
    | | Amount |
    |---|---|
    | 💰 Gross Revenue | {fmt_inr(total_revenue)} |
    | 📡 Broadcast Cost | − {fmt_inr(campaign_cost)} |
    | 🔻 Platform Deductions ({', '.join(found_deduction_cols) if found_deduction_cols else 'none detected'}) | − {fmt_inr(total_deductions)} |
    | **🟢 Net Revenue** | **{fmt_inr(net_revenue)}** |
    """)

# ── Pre-define display_cols so email section can use them ─────────────────────
_fixed = ["name","mobile","CourseName","Amount_Paid"]
_extra_preview = extra_cols  # will be refined in Verified Conversions section below

# ── Email Report ──────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-header">📧 Share Dashboard Report</div>', unsafe_allow_html=True)

include_net_rev = st.checkbox("🟢 Include Net Revenue in email", value=False,
    help="Tick to include Net Revenue KPI and breakdown in the email. Untick to hide it.")

ec, bc = st.columns([3,1])
with ec:
    recipient_email = st.text_input("Email", placeholder="example@domain.com", label_visibility="collapsed")
with bc:
    st.markdown("<br>", unsafe_allow_html=True)
    if st.button("📧 Email Report", use_container_width=True):
        if not recipient_email or "@" not in recipient_email:
            st.warning("⚠️ Enter a valid email address.")
        else:
            # Email uses fixed cols + any optional cols user selected in table
            email_cols = ["name","mobile","CourseName","Amount_Paid"] + st.session_state.get("selected_opt_cols", [])
            with st.spinner("Sending…"):
                ok, msg = send_report_email(
                    recipient_email, total_revenue, roi, total_conversions,
                    campaign_cost, total_deductions, net_revenue,
                    df_matched[email_cols], df_matched,
                    include_net_revenue=include_net_rev
                )
            if ok: st.success(msg)
            else:  st.error(msg)

# ── Chart ─────────────────────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-header">📊 Course Sales Performance</div>', unsafe_allow_html=True)

course_stats = (df_matched.groupby("CourseName")
    .agg(Revenue=("Amount_Paid","sum"), Students=("mobile","nunique"))
    .reset_index().sort_values("Revenue", ascending=True))

fig = px.bar(course_stats, x="Revenue", y="CourseName", orientation="h",
    title="Course Sales Performance",
    text=course_stats["Students"].apply(lambda s: f"👤 {s} student{'s' if s!=1 else ''}"),
    color="Revenue", color_continuous_scale=[[0,"#93C5FD"],[1,"#2563EB"]],
    labels={"Revenue":"Revenue (₹)","CourseName":"Course"})
fig.update_traces(textposition="outside", textfont_size=12, marker_line_width=0,
    hovertemplate="<b>%{y}</b><br>Revenue: ₹%{x:,.0f}<extra></extra>")
fig.update_layout(title_font=dict(size=16,family="Inter",color="#0F172A"),
    plot_bgcolor="#FFFFFF", paper_bgcolor="#FFFFFF",
    font=dict(family="Inter",color="#0F172A"), coloraxis_showscale=False,
    xaxis=dict(title="Revenue (₹)",gridcolor="#F1F5F9",tickprefix="₹",tickformat=","),
    yaxis=dict(title="",automargin=True), margin=dict(l=20,r=20,t=50,b=20),
    height=max(300, len(course_stats)*52+80))
st.plotly_chart(fig, use_container_width=True)

# ── Verified Conversions ──────────────────────────────────────────────────────
st.markdown("---")
st.markdown('<div class="section-header">✅ Verified Conversions</div>', unsafe_allow_html=True)

fixed_cols = ["name","mobile","CourseName","Amount_Paid"]

if extra_cols:
    selected_opt = st.multiselect(
        "➕ Add optional columns to view (also included in email & download)",
        options=extra_cols,
        default=[],
        key="opt_cols_select",
        help="Extra columns found in your sheet — select any to display. These will also be added to the email and Excel file."
    )
else:
    selected_opt = []

# Save to session state so email button (rendered above) can use it
st.session_state["selected_opt_cols"] = selected_opt

display_cols = fixed_cols + selected_opt
display_df = df_matched[display_cols].copy()
st.dataframe(display_df, use_container_width=True, hide_index=True)

# ── Downloads ─────────────────────────────────────────────────────────────────
dl1, dl2 = st.columns(2)
with dl1:
    csv_data = df_matched[display_cols].to_csv(index=False)
    st.download_button("⬇️ Download as CSV", data=csv_data,
        file_name="verified_conversions.csv", mime="text/csv")
with dl2:
    excel_data = make_excel_bytes(df_matched[display_cols])
    st.download_button("📥 Download as Excel", data=excel_data,
        file_name="verified_conversions.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
