import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from io import BytesIO
from datetime import datetime

# ============================================================
# إعدادات الصفحة
# ============================================================
st.set_page_config(
    page_title="AFM Support",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
    <style>
    .main-title {font-size: 2.2rem; font-weight: 700; color: #1f4e79;}
    .sub-title  {font-size: 1rem; color: #666; margin-bottom: 1.5rem;}
    div[data-testid="stMetricValue"] {font-size: 1.6rem;}
    </style>
""", unsafe_allow_html=True)

st.markdown('<div class="main-title">📊 AFM Support — تحليل تذاكر الصيانة</div>',
            unsafe_allow_html=True)
st.markdown('<div class="sub-title">ارفع ملف Excel واحصل على تحليل كامل لحظياً</div>',
            unsafe_allow_html=True)


# ============================================================
# دوال مساعدة
# ============================================================
CLOSURE_ORDER = ["<1 Hour", "1-2 Hours", "2-4 Hours", "4-8 Hours", "8-24 Hours", ">1 Day", "مفتوح"]


def closure_range(h):
    """تصنيف زمن الإغلاق."""
    if pd.isna(h):
        return "مفتوح"
    if h < 1:   return "<1 Hour"
    if h <= 2:  return "1-2 Hours"
    if h <= 4:  return "2-4 Hours"
    if h <= 8:  return "4-8 Hours"
    if h <= 24: return "8-24 Hours"
    return ">1 Day"


@st.cache_data(show_spinner="⏳ جاري قراءة وتحليل الملف...")
def load_data(file_bytes: bytes) -> pd.DataFrame:
    """قراءة شيت report وتنظيفه وإعادة حساب الأعمدة المشتقة."""
    xls = pd.ExcelFile(BytesIO(file_bytes))

    # نحاول نقرأ شيت report، لو مش موجود ناخد أول شيت فيه بيانات
    sheet_name = "report" if "report" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet_name, header=1)

    # تنظيف أسماء الأعمدة
    df.columns = [str(c).strip() for c in df.columns]

    # تحويل التواريخ
    for col in ["Request Date", "Close Date"]:
        if col in df.columns:
            df[col] = pd.to_datetime(df[col], errors="coerce")

    # إعادة حساب Closure Hours من التواريخ (لأن الملف فيه معادلات)
    if "Request Date" in df.columns and "Close Date" in df.columns:
        df["Closure Hours"] = (
            df["Close Date"] - df["Request Date"]
        ).dt.total_seconds() / 3600

    # إعادة حساب Closure Range
    if "Closure Hours" in df.columns:
        df["Closure Range"] = df["Closure Hours"].apply(closure_range)

    # عمود التاريخ فقط (للتحليل الزمني)
    if "Request Date" in df.columns:
        df["Request Day"]   = df["Request Date"].dt.date
        df["Request Week"]  = df["Request Date"].dt.to_period("W").astype(str)
        df["Request Month"] = df["Request Date"].dt.to_period("M").astype(str)

    # تنظيف النصوص
    for col in ["Project Name", "Status", "Sub-category", "From", "Created By", "Request Type"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip().replace({"nan": None})

    # حذف الصفوف الفاضية تماماً
    df = df.dropna(how="all")

    # التأكد من عدم تكرار Request NO
    if "Request NO" in df.columns:
        df = df.drop_duplicates(subset=["Request NO"], keep="first")

    return df.reset_index(drop=True)


def multiselect_filter(label, col, df):
    """إنشاء فلتر متعدد الاختيارات."""
    if col in df.columns:
        opts = sorted([x for x in df[col].dropna().unique() if str(x).strip()])
        return st.sidebar.multiselect(label, opts)
    return []


def style_bar(fig):
    fig.update_layout(
        margin=dict(l=10, r=10, t=40, b=10),
        font=dict(size=12),
        template="plotly_white",
    )
    return fig


def to_excel_bytes(sheets: dict) -> bytes:
    """تحويل عدة DataFrames إلى ملف Excel واحد."""
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
        for name, df_ in sheets.items():
            df_.to_excel(writer, sheet_name=name[:31], index=False)
    return buf.getvalue()


# ============================================================
# رفع الملف
# ============================================================
uploaded = st.sidebar.file_uploader("📁 ارفع ملف Excel (q3.xlsx)", type=["xlsx", "xls"])

if uploaded is None:
    st.info("⬆️ ارفع ملف التذاكر من الشريط الجانبي للبدء")
    st.stop()

df_full = load_data(uploaded.getvalue())
st.sidebar.success(f"✅ تم تحميل {len(df_full):,} تذكرة")


# ============================================================
# الفلاتر
# ============================================================
st.sidebar.header("🔎 الفلاتر")

sel_project   = multiselect_filter("المشروع",         "Project Name",  df_full)
sel_status    = multiselect_filter("الحالة",           "Status",        df_full)
sel_subcat    = multiselect_filter("التصنيف الفرعي",   "Sub-category",  df_full)
sel_from      = multiselect_filter("الجهة (From)",     "From",          df_full)
sel_reqtype   = multiselect_filter("نوع الطلب",        "Request Type",  df_full)

# فلتر التاريخ
date_filter = None
if "Request Date" in df_full.columns and df_full["Request Date"].notna().any():
    st.sidebar.markdown("**📅 نطاق التاريخ**")
    min_d = df_full["Request Date"].min().date()
    max_d = df_full["Request Date"].max().date()
    date_filter = st.sidebar.date_input(
        "من - إلى", value=(min_d, max_d),
        min_value=min_d, max_value=max_d,
    )

# فلتر SLA
st.sidebar.markdown("**⏱️ حد SLA (ساعات)**")
sla_hours = st.sidebar.slider("التذاكر المغلقة خلال", 1, 72, 24)

# تطبيق الفلاتر
df = df_full.copy()
if sel_project: df = df[df["Project Name"].isin(sel_project)]
if sel_status:  df = df[df["Status"].isin(sel_status)]
if sel_subcat:  df = df[df["Sub-category"].isin(sel_subcat)]
if sel_from:    df = df[df["From"].isin(sel_from)]
if sel_reqtype: df = df[df["Request Type"].isin(sel_reqtype)]

if date_filter and len(date_filter) == 2:
    start, end = pd.to_datetime(date_filter[0]), pd.to_datetime(date_filter[1])
    df = df[(df["Request Date"] >= start) & (df["Request Date"] <= end + pd.Timedelta(days=1))]

if len(df) == 0:
    st.warning("⚠️ لا توجد تذاكر مطابقة للفلاتر الحالية")
    st.stop()

st.caption(f"عدد التذاكر بعد الفلترة: **{len(df):,}** من أصل **{len(df_full):,}**")


# ============================================================
# KPI Cards
# ============================================================
closed_mask   = df["Status"] == "Closed"
progress_mask = df["Status"] == "In-Progress"
pending_mask  = df["Status"].isin(["Pending", "Waiting Client Approval", "Schedule"])
canceled_mask = df["Status"] == "Canceled"

# SLA Compliance
closed_df = df[closed_mask & df["Closure Hours"].notna()]
sla_ok    = (closed_df["Closure Hours"] <= sla_hours).sum()
sla_pct   = (sla_ok / len(closed_df) * 100) if len(closed_df) else 0

c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("إجمالي التذاكر",     f"{len(df):,}")
c2.metric("مغلقة",              f"{closed_mask.sum():,}")
c3.metric("قيد التنفيذ",        f"{progress_mask.sum():,}")
c4.metric("معلّقة/مجدولة",      f"{pending_mask.sum():,}")
c5.metric("ملغية",              f"{canceled_mask.sum():,}")
c6.metric("متوسط زمن الإغلاق",  f"{df['Closure Hours'].mean():.1f} س")

st.markdown("---")

# ============================================================
# التبويبات
# ============================================================
tab1, tab2, tab3, tab4, tab5, tab6, tab7, tab8 = st.tabs([
    "📌 حسب التصنيف",
    "📈 حسب الحالة",
    "⏱️ زمن الإغلاق",
    "🏢 حسب المشروع",
    "📅 التحليل الزمني",
    "🎯 SLA Compliance",
    "👥 أداء الفريق",
    "📋 البيانات والتصدير",
])


# ---------- 1) حسب التصنيف ----------
with tab1:
    st.subheader("عدد التذاكر حسب التصنيف الفرعي")
    vc = df["Sub-category"].value_counts().reset_index()
    vc.columns = ["التصنيف", "العدد"]
    fig = px.bar(vc, x="العدد", y="التصنيف", orientation="h",
                 text="العدد", color="العدد", color_continuous_scale="Blues")
    fig.update_traces(textposition="outside")
    fig.update_layout(height=max(400, len(vc) * 28), yaxis=dict(autorange="reversed"))
    st.plotly_chart(style_bar(fig), use_container_width=True)

    st.subheader("جدول التصنيف × الحالة")
    cross = pd.crosstab(df["Sub-category"], df["Status"])
    cross["الإجمالي"] = cross.sum(axis=1)
    cross = cross.sort_values("الإجمالي", ascending=False)
    st.dataframe(cross, use_container_width=True)


# ---------- 2) حسب الحالة ----------
with tab2:
    st.subheader("توزيع الحالات")
    vc = df["Status"].value_counts().reset_index()
    vc.columns = ["الحالة", "العدد"]
    col_a, col_b = st.columns([1, 1])
    with col_a:
        fig = px.pie(vc, names="الحالة", values="العدد", hole=0.45,
                     color_discrete_sequence=px.colors.qualitative.Set2)
        fig.update_traces(textinfo="percent+label")
        st.plotly_chart(fig, use_container_width=True)
    with col_b:
        st.dataframe(vc, use_container_width=True, hide_index=True)

    st.subheader("الحالة × الجهة (From)")
    cross = pd.crosstab(df["From"], df["Status"], margins=True, margins_name="الإجمالي")
    st.dataframe(cross, use_container_width=True)


# ---------- 3) زمن الإغلاق ----------
with tab3:
    st.subheader("توزيع زمن الإغلاق")
    vc = df["Closure Range"].value_counts().reindex(CLOSURE_ORDER).dropna().reset_index()
    vc.columns = ["الفترة", "العدد"]
    fig = px.bar(vc, x="الفترة", y="العدد", text="العدد",
                 color="الفترة", color_discrete_sequence=px.colors.sequential.Viridis)
    fig.update_traces(textposition="outside")
    st.plotly_chart(style_bar(fig), use_container_width=True)

    st.subheader("إحصائيات زمن الإغلاق (ساعة)")
    stats = df["Closure Hours"].describe()
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("المتوسط", f"{stats['mean']:.1f}")
    c2.metric("الوسيط", f"{stats['50%']:.1f}")
    c3.metric("الأقصى", f"{stats['max']:.1f}")
    c4.metric("عدد التذاكر المغلقة", f"{closed_df.shape[0]:,}")

    st.subheader("زمن الإغلاق حسب التصنيف الفرعي")
    by_cat = df.groupby("Sub-category")["Closure Hours"].agg(
        ["count", "mean", "median", "max"]
    ).round(2).sort_values("mean", ascending=False)
    by_cat.columns = ["العدد", "المتوسط", "الوسيط", "الأقصى"]
    st.dataframe(by_cat, use_container_width=True)


# ---------- 4) حسب المشروع ----------
with tab4:
    st.subheader("التذاكر حسب المشروع")
    vc = df["Project Name"].value_counts().reset_index()
    vc.columns = ["المشروع", "العدد"]
    fig = px.bar(vc, x="العدد", y="المشروع", orientation="h",
                 text="العدد", color="العدد", color_continuous_scale="Teal")
    fig.update_traces(textposition="outside")
    fig.update_layout(height=max(350, len(vc) * 32), yaxis=dict(autorange="reversed"))
    st.plotly_chart(style_bar(fig), use_container_width=True)

    st.subheader("المشروع × الحالة")
    cross = pd.crosstab(df["Project Name"], df["Status"], margins=True, margins_name="الإجمالي")
    st.dataframe(cross, use_container_width=True)


# ---------- 5) التحليل الزمني ----------
with tab5:
    st.subheader("عدد التذاكر يومياً")
    daily = df.groupby("Request Day").size().reset_index(name="العدد")
    fig = px.line(daily, x="Request Day", y="العدد", markers=True)
    fig.update_layout(xaxis_title="التاريخ", yaxis_title="عدد التذاكر")
    st.plotly_chart(style_bar(fig), use_container_width=True)

    st.subheader("عدد التذاكر أسبوعياً")
    weekly = df.groupby("Request Week").size().reset_index(name="العدد")
    fig = px.bar(weekly, x="Request Week", y="العدد", text="العدد")
    fig.update_traces(textposition="outside")
    st.plotly_chart(style_bar(fig), use_container_width=True)

    st.subheader("عدد التذاكر شهرياً")
    monthly = df.groupby("Request Month").size().reset_index(name="العدد")
    fig = px.bar(monthly, x="Request Month", y="العدد", text="العدد",
                 color="العدد", color_continuous_scale="Oranges")
    fig.update_traces(textposition="outside")
    st.plotly_chart(style_bar(fig), use_container_width=True)

    st.subheader("حسب ساعة اليوم")
    if "Request Time" in df.columns:
        hours = pd.to_datetime(df["Request Time"].astype(str), format="%H:%M", errors="coerce").dt.hour
        by_hour = hours.value_counts().sort_index().reset_index()
        by_hour.columns = ["الساعة", "العدد"]
        fig = px.bar(by_hour, x="الساعة", y="العدد", text="العدد",
                     color="العدد", color_continuous_scale="Purples")
        fig.update_traces(textposition="outside")
        st.plotly_chart(style_bar(fig), use_container_width=True)


# ---------- 6) SLA Compliance ----------
with tab6:
    st.subheader(f"SLA Compliance (إغلاق خلال {sla_hours} ساعة)")

    c1, c2, c3 = st.columns(3)
    c1.metric("تذاكر مغلقة", f"{len(closed_df):,}")
    c2.metric(f"في حدود SLA (≤{sla_hours}س)", f"{sla_ok:,}")
    c3.metric("نسبة الالتزام", f"{sla_pct:.1f}%")

    fig = go.Figure(go.Indicator(
        mode="gauge+number",
        value=sla_pct,
        number={"suffix": "%"},
        gauge={
            "axis": {"range": [0, 100]},
            "bar": {"color": "#2ca02c" if sla_pct >= 90 else "#ff7f0e" if sla_pct >= 75 else "#d62728"},
            "steps": [
                {"range": [0, 75],  "color": "#f8d7da"},
                {"range": [75, 90], "color": "#fff3cd"},
                {"range": [90, 100], "color": "#d4edda"},
            ],
            "threshold": {"line": {"color": "black", "width": 3},
                          "thickness": 0.8, "value": 90},
        },
    ))
    fig.update_layout(height=300, margin=dict(l=20, r=20, t=30, b=10))
    st.plotly_chart(fig, use_container_width=True)

    st.subheader("SLA حسب الجهة (From)")
    sla_by_from = df.groupby("From").apply(
        lambda g: pd.Series({
            "العدد الكلي": len(g),
            "مغلق": g["Closure Hours"].notna().sum(),
            f"ضمن SLA (≤{sla_hours}س)": (g["Closure Hours"] <= sla_hours).sum(),
            "نسبة الالتزام %": round(
                (g["Closure Hours"] <= sla_hours).sum() /
                max(g["Closure Hours"].notna().sum(), 1) * 100, 1
            ),
        }),
        include_groups=False,
    ).sort_values("نسبة الالتزام %", ascending=False)
    st.dataframe(sla_by_from, use_container_width=True)

    st.subheader("SLA حسب التصنيف الفرعي")
    sla_by_cat = df.groupby("Sub-category").apply(
        lambda g: pd.Series({
            "مغلق": g["Closure Hours"].notna().sum(),
            f"ضمن SLA": (g["Closure Hours"] <= sla_hours).sum(),
            "نسبة الالتزام %": round(
                (g["Closure Hours"] <= sla_hours).sum() /
                max(g["Closure Hours"].notna().sum(), 1) * 100, 1
            ),
        }),
        include_groups=False,
    ).sort_values("نسبة الالتزام %", ascending=False)
    st.dataframe(sla_by_cat, use_container_width=True)

    st.subheader("تذاكر تجاوزت SLA (أعلى 50)")
    breach = closed_df[closed_df["Closure Hours"] > sla_hours].sort_values(
        "Closure Hours", ascending=False
    ).head(50)
    cols = [c for c in ["Request NO", "Ticket No", "Project Name", "Sub-category",
                        "Status", "Request Date", "Close Date", "Closure Hours"] if c in breach.columns]
    st.dataframe(breach[cols], use_container_width=True, hide_index=True)


# ---------- 7) أداء الفريق ----------
with tab7:
    st.subheader("أداء الموظفين (Created By)")
    perf = df.groupby("Created By").apply(
        lambda g: pd.Series({
            "عدد التذاكر": len(g),
            "مغلقة": g["Status"].eq("Closed").sum(),
            "متوسط الإغلاق (س)": round(g["Closure Hours"].mean(), 1) if g["Closure Hours"].notna().any() else None,
            f"ضمن SLA (≤{sla_hours}س)": (g["Closure Hours"] <= sla_hours).sum(),
            "نسبة الإغلاق %": round(g["Status"].eq("Closed").mean() * 100, 1),
        }),
        include_groups=False,
    ).sort_values("عدد التذاكر", ascending=False)
    st.dataframe(perf, use_container_width=True)

    st.subheader("أعلى 15 موظف بعدد التذاكر")
    top = perf.head(15).reset_index()
    fig = px.bar(top, x="عدد التذاكر", y="Created By", orientation="h",
                 text="عدد التذاكر", color="عدد التذاكر", color_continuous_scale="Blues")
    fig.update_traces(textposition="outside")
    fig.update_layout(yaxis=dict(autorange="reversed"), height=500)
    st.plotly_chart(style_bar(fig), use_container_width=True)

    st.subheader("أفضل 15 موظف بمتوسط زمن إغلاق (الأقل)")
    fastest = perf.dropna(subset=["متوسط الإغلاق (س)"]).sort_values(
        "متوسط الإغلاق (س)"
    ).head(15).reset_index()
    fig = px.bar(fastest, x="متوسط الإغلاق (س)", y="Created By", orientation="h",
                 text="متوسط الإغلاق (س)", color="متوسط الإغلاق (س)",
                 color_continuous_scale="Greens")
    fig.update_traces(textposition="outside")
    fig.update_layout(yaxis=dict(autorange="reversed"), height=500)
    st.plotly_chart(style_bar(fig), use_container_width=True)


# ---------- 8) البيانات والتصدير ----------
with tab8:
    st.subheader("البيانات الخام (بعد الفلترة)")
    st.dataframe(df, use_container_width=True, height=450)

    st.markdown("### ⬇️ تصدير")

    col_a, col_b = st.columns(2)

    with col_a:
        csv = df.to_csv(index=False).encode("utf-8-sig")
        st.download_button(
            "📄 تحميل CSV (البيانات المفلترة)",
            csv, "afm_filtered.csv", "text/csv",
            use_container_width=True,
        )

    with col_b:
        status_tbl = df["Status"].value_counts().reset_index()
        status_tbl.columns = ["الحالة", "العدد"]

        subcat_tbl = df["Sub-category"].value_counts().reset_index()
        subcat_tbl.columns = ["التصنيف", "العدد"]

        project_tbl = df["Project Name"].value_counts().reset_index()
        project_tbl.columns = ["المشروع", "العدد"]

        closure_tbl = df["Closure Range"].value_counts().reindex(CLOSURE_ORDER).dropna().reset_index()
        closure_tbl.columns = ["الفترة", "العدد"]

        perf_tbl = df.groupby("Created By").agg(
            عدد_التذاكر=("Request NO", "count"),
            مغلقة=("Status", lambda x: (x == "Closed").sum()),
            متوسط_الإغلاق_ساعة=("Closure Hours", "mean"),
        ).round(1).reset_index()

        sheets = {
            "البيانات": df,
            "الحالة": status_tbl,
            "التصنيف": subcat_tbl,
            "المشروع": project_tbl,
            "زمن الإغلاق": closure_tbl,
            "أداء الفريق": perf_tbl,
        }

        excel_bytes = to_excel_bytes(sheets)
        st.download_button(
            "📊 تحميل Excel شامل (كل التحليلات)",
            excel_bytes, "afm_report.xlsx",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    st.markdown("---")
    st.markdown("### 📌 ملخص سريع للبيانات")
    summary = pd.DataFrame({
        "المؤشر": [
            "إجمالي التذاكر", "مغلقة", "قيد التنفيذ", "معلّقة/مجدولة", "ملغية",
            "متوسط زمن الإغلاق (س)", "نسبة SLA",
        ],
        "القيمة": [
            f"{len(df):,}",
            f"{closed_mask.sum():,}",
            f"{progress_mask.sum():,}",
            f"{pending_mask.sum():,}",
            f"{canceled_mask.sum():,}",
            f"{df['Closure Hours'].mean():.1f}",
            f"{sla_pct:.1f}%",
        ],
    })
    st.table(summary)
