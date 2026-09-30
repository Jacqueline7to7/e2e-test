import streamlit as st
import pandas as pd
import io
import os
from datetime import datetime
import openpyxl
from openpyxl.styles import Font, Alignment, PatternFill, Border
from copy import copy

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
TEMPLATE_DIR = os.path.join(BASE_DIR, "templates")

st.set_page_config(
    page_title="E2E 自动化交付平台",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown("""
<style>
    .stApp { background: #F9FAFB; }
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    .main .block-container { max-width: 1400px; padding: 24px 32px; }
    .page-title { font-size: 24px; font-weight: 700; color: #111827; margin: 0 0 4px 0; }
    .page-subtitle { font-size: 13px; color: #6B7280; margin: 0 0 24px 0; }
    .kpi-card {
        background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 12px;
        padding: 20px; box-shadow: 0 1px 2px rgba(0,0,0,0.05);
    }
    .kpi-label { font-size: 13px; color: #6B7280; margin-bottom: 8px; }
    .kpi-value { font-size: 30px; font-weight: 700; color: #111827; line-height: 1.2; }
    .kpi-trend { font-size: 12px; margin-top: 8px; }
    .e2e-card {
        background: #FFFFFF; border: 1px solid #E5E7EB; border-radius: 12px;
        padding: 20px; box-shadow: 0 1px 2px rgba(0,0,0,0.05); margin-bottom: 16px;
    }
    .check-pass { color: #16A34A; font-size: 14px; padding: 4px 0; }
    .check-fail { color: #DC2626; font-size: 14px; padding: 4px 0; }
    .check-warn { color: #D97706; font-size: 14px; padding: 4px 0; }
</style>
""", unsafe_allow_html=True)

if "uploaded_data" not in st.session_state:
    st.session_state["uploaded_data"] = []
if "fetch_df" not in st.session_state:
    st.session_state["fetch_df"] = None
if "backfill_df" not in st.session_state:
    st.session_state["backfill_df"] = None
if "result_df" not in st.session_state:
    st.session_state["result_df"] = None
if "qc_issues" not in st.session_state:
    st.session_state["qc_issues"] = []

# ============ 校验函数 ============
RED_REQUIRED = [
    "Brand", "Data Period", "Product", "Category", "Sub-category",
    "Campaign Type", "Campaign Period", "跨域项目名称",
    "计划有效周期", "笔记发布日期", "达人名称", "笔记ID",
    "达人Spending", "合约Spending", "竞价Spending",
    "品线天猫商品ID List", "TM单品主链客单价", "TM全店客单价",
]
DY_REQUIRED = [
    "Data Period", "Category", "Sub-category", "Campaign type", "Product",
    "星图订单ID", "达人名称", "发布日期", "KOL花费",
    "CPC Paid Media花费", "CPM Paid Media花费（热推）",
    "竞价投流花费（种草通）", "品线抖音商品ID List",
    "DY单品主链客单价", "DY全店客单价",
]

def validate_dataframe(df, platform):
    results = []
    required = RED_REQUIRED if platform == "RED" else DY_REQUIRED
    missing_cols = [c for c in required if c not in df.columns]
    if missing_cols:
        return "FAIL", [("FAIL", f"缺少必填列：{', '.join(missing_cols)}")]
    results.append(("PASS", f"必填列完整（共 {len(required)} 列）"))

    empty_cols = [c for c in required if df[c].isnull().any() or (df[c].astype(str).str.strip() == "").any()]
    if empty_cols:
        results.append(("FAIL", f"以下列存在空值：{', '.join(empty_cols)}"))
    else:
        results.append(("PASS", "所有必填列无空值"))

    spend_cols = [c for c in df.columns if "Spending" in c or "花费" in c]
    neg_cols = [c for c in spend_cols if (pd.to_numeric(df[c], errors="coerce") < 0).any()]
    if neg_cols:
        results.append(("FAIL", f"以下花费列存在负值：{', '.join(neg_cols)}"))
    else:
        results.append(("PASS", "花费列均为非负"))

    aov_cols = [c for c in df.columns if "客单价" in c]
    bad_aov = [c for c in aov_cols if (pd.to_numeric(df[c], errors="coerce") <= 0).any()]
    if bad_aov:
        results.append(("WARN", f"以下客单价列存在 <=0 的值：{', '.join(bad_aov)}"))
    else:
        results.append(("PASS", "客单价列均 > 0"))

    id_col = "笔记ID" if platform == "RED" else "星图订单ID"
    if id_col in df.columns:
        dup = df[df.duplicated(subset=[id_col], keep=False)]
        if len(dup) > 0:
            results.append(("WARN", f"{id_col} 存在重复：{dup[id_col].nunique()} 个重复值"))
        else:
            results.append(("PASS", f"{id_col} 无重复"))

    results.append(("PASS", f"共 {len(df)} 行数据，{len(df.columns)} 列"))
    if any(r[0] == "FAIL" for r in results):
        return "FAIL", results
    elif any(r[0] == "WARN" for r in results):
        return "WARN", results
    return "PASS", results

# ============ 取数清单生成 ============
def build_fetch_list(uploaded_data):
    rows = []
    idx = 1
    for item in uploaded_data:
        df = item.get("df")
        if df is None:
            continue
        plat = item["平台"]
        for _, r in df.iterrows():
            if plat == "RED":
                brand = r.get("Brand", "")
                product = r.get("Product", "")
                project = r.get("跨域项目名称", product)
                order = r.get("笔记ID", "")
                period = r.get("计划有效周期", "") or r.get("Campaign Period", "")
                read_dim = "红书阅读数"
            else:
                brand = r.get("Brand", r.get("Product", ""))
                product = r.get("Product", "")
                project = product
                order = r.get("星图订单ID", "")
                period = r.get("Data Period", "")
                read_dim = "视频观看数"

            for dim in [read_dim, "进店UV", "购买人数", "渠道新客数"]:
                rows.append({
                    "#": idx,
                    "品牌": str(brand),
                    "平台": plat,
                    "项目名称": str(project),
                    "订单名称": str(order),
                    "时间范围": str(period),
                    "维度": dim,
                    "窗口": "R30",
                    "状态": "待查询",
                    "数值": None,
                })
                idx += 1
    return pd.DataFrame(rows)

# ============ 回填匹配 ============
MATCH_KEYS = ["品牌", "平台", "项目名称", "订单名称", "时间范围", "维度"]

def match_backfill(fetch_df, backfill_df):
    if fetch_df is None or backfill_df is None:
        return None
    missing = [k for k in MATCH_KEYS + ["数值"] if k not in backfill_df.columns]
    if missing:
        return {"error": f"回填文件缺少列：{', '.join(missing)}"}

    merged = fetch_df.merge(
        backfill_df[MATCH_KEYS + ["数值"]],
        on=MATCH_KEYS, how="left", suffixes=("", "_回填"),
    )
    if "数值_回填" in merged.columns:
        merged["数值"] = merged["数值_回填"].combine_first(merged["数值"])
        merged = merged.drop(columns=["数值_回填"])
    merged["状态"] = merged.apply(
        lambda r: "已完成" if pd.notna(r["数值"]) else r["状态"], axis=1
    )
    return merged

# ============ 聚合上传数据 ============
def aggregate_uploaded(uploaded_data):
    """按 品牌+平台+项目 聚合，提取花费、客单价"""
    rows = []
    for item in uploaded_data:
        df = item.get("df")
        plat = item["平台"]
        if df is None:
            continue
        for _, r in df.iterrows():
            if plat == "RED":
                brand = r.get("Brand", "")
                product = r.get("Product", "")
                project = r.get("跨域项目名称", product)
                spend = sum([
                    pd.to_numeric(r.get("达人Spending", 0), errors="coerce") or 0,
                    pd.to_numeric(r.get("合约Spending", 0), errors="coerce") or 0,
                    pd.to_numeric(r.get("竞价Spending", 0), errors="coerce") or 0,
                ])
                aov = pd.to_numeric(r.get("TM全店客单价", 0), errors="coerce") or 0
                item_aov = pd.to_numeric(r.get("TM单品主链客单价", 0), errors="coerce") or 0
            else:
                brand = r.get("Brand", r.get("Product", ""))
                product = r.get("Product", "")
                project = product
                spend = sum([
                    pd.to_numeric(r.get("KOL花费", 0), errors="coerce") or 0,
                    pd.to_numeric(r.get("CPC Paid Media花费", 0), errors="coerce") or 0,
                    pd.to_numeric(r.get("CPM Paid Media花费（热推）", 0), errors="coerce") or 0,
                    pd.to_numeric(r.get("竞价投流花费（种草通）", 0), errors="coerce") or 0,
                ])
                aov = pd.to_numeric(r.get("DY全店客单价", 0), errors="coerce") or 0
                item_aov = pd.to_numeric(r.get("DY单品主链客单价", 0), errors="coerce") or 0
            rows.append({
                "品牌": str(brand),
                "平台": plat,
                "项目名称": str(project),
                "产品": str(product),
                "媒体花费": spend,
                "全店客单价": aov,
                "单品客单价": item_aov,
            })
    if not rows:
        return pd.DataFrame(columns=["品牌", "平台", "项目名称", "产品", "媒体花费", "全店客单价", "单品客单价"])
    return pd.DataFrame(rows).groupby(["品牌", "平台", "项目名称"], as_index=False).agg({
        "产品": "first",
        "媒体花费": "sum",
        "全店客单价": "first",
        "单品客单价": "first",
    })

# ============ 计算 E2E 指标 ============
def calculate_metrics(uploaded_data, backfill_df):
    agg = aggregate_uploaded(uploaded_data)
    if len(agg) == 0:
        return None, []

    bf = backfill_df.copy()
    bf["数值"] = pd.to_numeric(bf["数值"], errors="coerce").fillna(0)

    if bf["数值"].sum() == 0:
        return None, [("FAIL", "全局", "回填数据全部为空或为 0，无法计算")]

    pivoted = (
        bf.groupby(["品牌", "平台", "项目名称", "维度"])["数值"]
        .sum()
        .unstack(fill_value=0)
        .reset_index()
    )

    if "红书阅读数" in pivoted.columns:
        pivoted["阅读数"] = pivoted["红书阅读数"]
    elif "视频观看数" in pivoted.columns:
        pivoted["阅读数"] = pivoted["视频观看数"]
    else:
        pivoted["阅读数"] = 0

    for col in ["进店UV", "购买人数", "渠道新客数", "阅读数"]:
        if col not in pivoted.columns:
            pivoted[col] = 0

    result = agg.merge(
        pivoted[["品牌", "平台", "项目名称", "阅读数", "进店UV", "购买人数", "渠道新客数"]],
        on=["品牌", "平台", "项目名称"], how="left",
    ).fillna(0)

    # ===== 关键修复：强制所有数值列转成 numeric =====
    numeric_cols = ["媒体花费", "全店客单价", "单品客单价",
                    "阅读数", "进店UV", "购买人数", "渠道新客数"]
    for c in numeric_cols:
        if c in result.columns:
            result[c] = pd.to_numeric(result[c], errors="coerce").fillna(0).astype(float)

    def safe_div(a, b):
        return (a / b).where(b != 0, 0.0)

    result["进店率"] = safe_div(result["进店UV"], result["阅读数"])
    result["购买转化率"] = safe_div(result["购买人数"], result["进店UV"])
    result["渠道新客占比"] = safe_div(result["渠道新客数"], result["购买人数"])
    result["获客成本"] = safe_div(result["媒体花费"], result["购买人数"])
    result["新客成本"] = safe_div(result["媒体花费"], result["渠道新客数"])
    result["预估新客GMV"] = result["全店客单价"] * result["渠道新客数"]
    result["预估整体新客ROI"] = safe_div(result["预估新客GMV"], result["媒体花费"])

    issues = []
    for _, r in result.iterrows():
        key = f"{r['品牌']} / {r['项目名称']}"
        if r["阅读数"] == 0:
            issues.append(("FAIL", key, "阅读数为 0，无法计算进店率"))
        if r["阅读数"] > 0 and r["进店UV"] > r["阅读数"]:
            issues.append(("FAIL", key, f"进店UV（{int(r['进店UV'])}）> 阅读数（{int(r['阅读数'])}）"))
        if r["进店UV"] > 0 and r["购买人数"] > r["进店UV"]:
            issues.append(("FAIL", key, f"购买人数（{int(r['购买人数'])}）> 进店UV（{int(r['进店UV'])}）"))
        if r["购买人数"] > 0 and r["渠道新客数"] > r["购买人数"]:
            issues.append(("FAIL", key, f"渠道新客数（{int(r['渠道新客数'])}）> 购买人数（{int(r['购买人数'])}）"))
        if r["媒体花费"] <= 0:
            issues.append(("WARN", key, "媒体花费为 0 或负"))
        if r["全店客单价"] <= 0:
            issues.append(("WARN", key, "全店客单价为 0 或负"))
    return result, issues

# ============ 侧边栏 ============
with st.sidebar:
    st.markdown("### 📊 E2E 自动化交付平台")
    st.markdown("---")
    page = st.radio(
        "导航",
        ["📊 概览", "📋 信息收集", "📥 取数清单", "📤 数据回填",
         "🧮 计算质检", "📄 报告输出", "📈 历史趋势"],
        label_visibility="collapsed",
    )
    st.markdown("---")
    st.markdown("**全局筛选**")
    period_filter = st.selectbox("周期", ["2026-08", "2026-07", "2026-06"])
    platform_filter = st.selectbox("平台", ["全部", "RED", "DY"])
    brand_filter = st.selectbox("品牌", ["全部", "品牌A", "品牌B", "品牌C", "品牌D", "品牌E"])
    st.markdown("---")
    st.caption("版本 v0.5")

# ============ 页面：概览 ============
if page == "📊 概览":
    st.markdown('<div class="page-title">E2E 交付概览</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period_filter} · 平台：{platform_filter} · 品牌：{brand_filter}</div>', unsafe_allow_html=True)

    cols = st.columns(4)
    kpis = [
        ("品牌总数", "12", "全部平台", "#111827"),
        ("已完成", "8", "🟢 67%", "#16A34A"),
        ("待处理", "3", "🟡 25%", "#D97706"),
        ("异常", "1", "🔴 8%", "#DC2626"),
    ]
    for col, (label, value, trend, color) in zip(cols, kpis):
        with col:
            st.markdown(f"""
            <div class="kpi-card">
                <div class="kpi-label">{label}</div>
                <div class="kpi-value" style="color: {color};">{value}</div>
                <div class="kpi-trend">{trend}</div>
            </div>
            """, unsafe_allow_html=True)

    st.markdown("<br>", unsafe_allow_html=True)
    st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
    st.markdown("#### 各品牌交付状态")
    st.info("待接入真实数据后，这里显示各品牌的交付状态表。")
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
    st.markdown("#### DDL 提醒")
    st.warning("品牌C DMP 取数未完成，剩余 3 项，DDL：09-05 18:00")
    st.error("品牌E 数据异常：新客数 > 购买人数，需复核")
    st.markdown('</div>', unsafe_allow_html=True)

# ============ 页面：信息收集 ============
elif page == "📋 信息收集":
    st.markdown('<div class="page-title">信息收集与校验</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period_filter} · 平台：{platform_filter} · 品牌：{brand_filter}</div>', unsafe_allow_html=True)

    st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
    st.markdown("#### 下载示例模板")
    st.caption("如果你还没有收集表，可以先下载示例模板，填完后上传测试。")

    col1, col2 = st.columns([1, 3])
    with col1:
        template_type = st.selectbox("模板类型", ["红书", "抖音"])
    with col2:
        st.write("")

    if st.button("生成示例模板"):
        if template_type == "红书":
            sample = pd.DataFrame({
                "Brand": ["品牌A"],
                "Data Period": ["2026.07.01-08.31"],
                "Product": ["产品A"],
                "Category": ["Make up"],
                "Sub-category": ["Lip"],
                "Campaign Type": ["RP+CVD"],
                "Campaign Period": ["7.1-8.19"],
                "跨域项目名称": ["产品A"],
                "计划有效周期": ["7.1-8.19"],
                "笔记发布日期": ["2026-07-15"],
                "达人名称": ["达人小王"],
                "笔记ID": ["note_001"],
                "达人Spending": [1830860],
                "合约Spending": [519374],
                "竞价Spending": [0],
                "品线天猫商品ID List": ["123456789"],
                "TM单品主链客单价": [380],
                "TM全店客单价": [1015],
            })
        else:
            sample = pd.DataFrame({
                "Data Period": ["2026.08"],
                "Category": ["Skincare"],
                "Sub-category": ["Lotion"],
                "Campaign type": ["Repush"],
                "Product": ["产品O"],
                "星图订单ID": ["order_001"],
                "达人名称": ["达人小李"],
                "发布日期": ["2026-08-15"],
                "KOL花费": [200000],
                "CPC Paid Media花费": [50000],
                "CPM Paid Media花费（热推）": [30000],
                "竞价投流花费（种草通）": [10000],
                "品线抖音商品ID List": ["987654321"],
                "DY单品主链客单价": [152],
                "DY全店客单价": [991],
            })

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
            sample.to_excel(writer, index=False, sheet_name="Sheet1")
        buffer.seek(0)

        st.download_button(
            label="⬇ 下载示例模板 Excel",
            data=buffer,
            file_name=f"{template_type}_示例模板.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
    st.markdown('</div>', unsafe_allow_html=True)

    st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
    st.markdown("#### 上传收集表")
    uploaded_files = st.file_uploader(
        "上传红书/抖音信息收集表（支持多文件）",
        type=["xlsx", "csv"],
        accept_multiple_files=True,
    )

    if uploaded_files:
        all_results = []
        for f in uploaded_files:
            try:
                if f.name.endswith(".csv"):
                    df = pd.read_csv(f)
                else:
                    df = pd.read_excel(f)

                if "红书" in f.name or "RED" in f.name.upper():
                    plat = "RED"
                elif "抖音" in f.name or "DY" in f.name.upper():
                    plat = "DY"
                else:
                    plat = "RED" if "笔记ID" in df.columns else "DY"

                status, details = validate_dataframe(df, plat)
                all_results.append({
                    "文件名": f.name, "平台": plat, "行数": len(df),
                    "状态": status, "明细": details, "df": df,
                })
            except Exception as e:
                all_results.append({
                    "文件名": f.name, "平台": "未知", "行数": 0,
                    "状态": "FAIL", "明细": [("FAIL", f"文件解析失败：{str(e)}")],
                    "df": None,
                })

        st.session_state["uploaded_data"] = all_results

        total = len(all_results)
        passed = sum(1 for r in all_results if r["状态"] == "PASS")
        warned = sum(1 for r in all_results if r["状态"] == "WARN")
        failed = sum(1 for r in all_results if r["状态"] == "FAIL")

        c1, c2, c3, c4 = st.columns(4)
        c1.metric("文件总数", total)
        c2.metric("通过", passed)
        c3.metric("警告", warned)
        c4.metric("失败", failed)

        st.markdown("---")
        st.markdown("#### 校验结果")
        for r in all_results:
            status_icon = {"PASS": "🟢", "WARN": "🟡", "FAIL": "🔴"}[r["状态"]]
            with st.expander(f"{status_icon} {r['文件名']}（{r['平台']}，{r['行数']} 行）"):
                for level, msg in r["明细"]:
                    if level == "PASS":
                        st.markdown(f'<div class="check-pass">✅ {msg}</div>', unsafe_allow_html=True)
                    elif level == "WARN":
                        st.markdown(f'<div class="check-warn">⚠ {msg}</div>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="check-fail">❌ {msg}</div>', unsafe_allow_html=True)

    st.markdown('</div>', unsafe_allow_html=True)

# ============ 页面：取数清单 ============
elif page == "📥 取数清单":
    st.markdown('<div class="page-title">DMP / 云图取数清单</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period_filter} · 平台：{platform_filter} · 品牌：{brand_filter}</div>', unsafe_allow_html=True)

    if not st.session_state["uploaded_data"]:
        st.warning("请先到「📋 信息收集」页上传收集表，系统会自动生成取数清单。")
    else:
        st.info("💡 请按以下清单到 DMP 后台逐项查询，查询完成后到「📤 数据回填」页上传或粘贴结果。")

        fetch_df = build_fetch_list(st.session_state["uploaded_data"])

        if len(fetch_df) == 0:
            st.warning("未能从收集表中提取到有效数据，请检查收集表格式。")
        else:
            total = len(fetch_df)
            done = len(fetch_df[fetch_df["状态"] == "已完成"])
            pending = len(fetch_df[fetch_df["状态"] == "待查询"])
            error = len(fetch_df[fetch_df["状态"] == "异常"])

            c1, c2, c3, c4 = st.columns(4)
            c1.metric("任务总数", total)
            c2.metric("已完成", done)
            c3.metric("待查询", pending)
            c4.metric("异常", error)

            st.progress(done / total if total > 0 else 0, text=f"取数进度：{done} / {total}")

            st.markdown("#### 取数任务清单")
            st.caption("每个活动对应 4 个维度：阅读数 / 进店UV / 购买人数 / 渠道新客数。")

            edited_df = st.data_editor(
                fetch_df,
                column_config={
                    "状态": st.column_config.SelectboxColumn(
                        "状态",
                        options=["待查询", "已完成", "异常"],
                        required=True, default="待查询",
                    ),
                },
                use_container_width=True, hide_index=True, key="fetch_editor",
            )
            st.session_state["fetch_df"] = edited_df

            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
                edited_df.to_excel(writer, index=False, sheet_name="取数清单")
            buffer.seek(0)

            st.download_button(
                label="⬇ 导出取数清单 Excel",
                data=buffer,
                file_name=f"取数清单_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# ============ 页面：数据回填 ============
elif page == "📤 数据回填":
    st.markdown('<div class="page-title">数据回填</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period_filter} · 平台：{platform_filter} · 品牌：{brand_filter}</div>', unsafe_allow_html=True)

    if st.session_state["fetch_df"] is None:
        st.warning("请先到「📥 取数清单」页生成取数清单。")
    else:
        fetch_df = st.session_state["fetch_df"]

        st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
        st.markdown("#### 下载回填模板")
        st.caption("下载模板后，把 DMP 查询结果填入「数值」列，再上传回来。")

        template = fetch_df[MATCH_KEYS].copy()
        template["数值"] = None

        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            template.to_excel(writer, index=False, sheet_name="回填模板")
        buf.seek(0)

        st.download_button(
            label="⬇ 下载回填模板 Excel",
            data=buf,
            file_name=f"回填模板_{datetime.now().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.markdown('</div>', unsafe_allow_html=True)

        tab1, tab2 = st.tabs(["上传文件", "手动粘贴"])

        with tab1:
            backfill_file = st.file_uploader("上传回填文件（Excel / CSV）", type=["xlsx", "csv"], key="backfill_upload")
            if backfill_file:
                try:
                    if backfill_file.name.endswith(".csv"):
                        bdf = pd.read_csv(backfill_file)
                    else:
                        bdf = pd.read_excel(backfill_file)
                    result = match_backfill(fetch_df, bdf)
                    if isinstance(result, dict) and "error" in result:
                        st.error(result["error"])
                    else:
                        st.session_state["backfill_df"] = result
                        st.success(f"回填文件解析成功，共 {len(bdf)} 行数据。")
                except Exception as e:
                    st.error(f"文件解析失败：{str(e)}")

        with tab2:
            pasted = st.text_area(
                "粘贴数据（第一行为表头，列用 Tab 分隔）",
                height=200,
                placeholder="品牌\t平台\t项目名称\t订单名称\t时间范围\t维度\t数值\n品牌A\tRED\t产品A\tnote_001\t7.1-8.19\t进店UV\t544798",
            )
            if pasted:
                try:
                    from io import StringIO
                    bdf = pd.read_csv(StringIO(pasted), sep="\t")
                    result = match_backfill(fetch_df, bdf)
                    if isinstance(result, dict) and "error" in result:
                        st.error(result["error"])
                    else:
                        st.session_state["backfill_df"] = result
                        st.success(f"粘贴数据解析成功，共 {len(bdf)} 行数据。")
                except Exception as e:
                    st.error(f"粘贴数据解析失败：{str(e)}")

        if st.session_state["backfill_df"] is not None:
            st.markdown("---")
            st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
            st.markdown("#### 匹配结果")
            result_df = st.session_state["backfill_df"]
            matched = result_df["数值"].notna().sum()
            total = len(result_df)

            c1, c2, c3 = st.columns(3)
            c1.metric("任务总数", total)
            c2.metric("已匹配", matched)
            c3.metric("未匹配", total - matched)

            display_df = result_df.copy()
            display_df["匹配状态"] = display_df["数值"].apply(lambda v: "✅ 已匹配" if pd.notna(v) else "❌ 未匹配")
            st.dataframe(
                display_df[["品牌", "平台", "项目名称", "订单名称", "时间范围", "维度", "数值", "匹配状态"]],
                use_container_width=True, hide_index=True,
            )

            unmatched = display_df[display_df["匹配状态"] == "❌ 未匹配"]
            if len(unmatched) > 0:
                st.warning(f"有 {len(unmatched)} 项未匹配到数据。")

            buf2 = io.BytesIO()
            with pd.ExcelWriter(buf2, engine="openpyxl") as writer:
                display_df.to_excel(writer, index=False, sheet_name="匹配结果")
            buf2.seek(0)
            st.download_button(
                label="⬇ 导出匹配结果 Excel", data=buf2,
                file_name=f"匹配结果_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
            st.markdown('</div>', unsafe_allow_html=True)

# ============ 页面：计算质检 ============
elif page == "🧮 计算质检":
    st.markdown('<div class="page-title">计算与质检</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period_filter} · 平台：{platform_filter} · 品牌：{brand_filter}</div>', unsafe_allow_html=True)

    if st.session_state["uploaded_data"] is None or len(st.session_state["uploaded_data"]) == 0:
        st.warning("请先到「📋 信息收集」页上传收集表。")
    elif st.session_state["backfill_df"] is None:
        st.warning("请先到「📤 数据回填」页完成数据回填。")
    else:
        col1, col2, col3 = st.columns([1, 1, 4])
        with col1:
            if st.button("开始计算", type="primary"):
                with st.spinner("正在计算 E2E 指标..."):
                    result_df, issues = calculate_metrics(
                        st.session_state["uploaded_data"],
                        st.session_state["backfill_df"],
                    )
                    st.session_state["result_df"] = result_df
                    st.session_state["qc_issues"] = issues
                st.success("计算完成！")
        with col2:
            if st.button("重新计算"):
                st.session_state["result_df"] = None
                st.session_state["qc_issues"] = []
                st.info("已清除计算结果，请点击「开始计算」。")

        if st.session_state["result_df"] is not None:
            result_df = st.session_state["result_df"]
            issues = st.session_state["qc_issues"]

            # 质检汇总
            fail_count = sum(1 for i in issues if i[0] == "FAIL")
            warn_count = sum(1 for i in issues if i[0] == "WARN")

            st.markdown("---")
            st.markdown("#### 质检结果")
            c1, c2, c3 = st.columns(3)
            c1.metric("活动总数", len(result_df))
            c2.metric("失败项", fail_count, delta_color="inverse" if fail_count > 0 else "normal")
            c3.metric("警告项", warn_count)

            if fail_count > 0:
                st.error(f"存在 {fail_count} 项失败，需复核后再生成报告。")
            elif warn_count > 0:
                st.warning(f"存在 {warn_count} 项警告，可继续但建议复核。")
            else:
                st.success("全部通过，可以进入报告输出。")

            if issues:
                st.markdown("##### 异常明细")
                for level, key, msg in issues:
                    if level == "FAIL":
                        st.markdown(f'<div class="check-fail">❌ {key}：{msg}</div>', unsafe_allow_html=True)
                    else:
                        st.markdown(f'<div class="check-warn">⚠ {key}：{msg}</div>', unsafe_allow_html=True)

            # 关键指标预览
            st.markdown("---")
            st.markdown("#### 关键指标预览")

            display_cols = ["品牌", "平台", "项目名称", "媒体花费", "阅读数", "进店UV", "购买人数",
                            "渠道新客数", "进店率", "购买转化率", "新客成本", "预估整体新客ROI"]
            display_df = result_df[display_cols].copy()

            # 格式化
            display_df["媒体花费"] = display_df["媒体花费"].apply(lambda x: f"{x:,.0f}")
            display_df["进店率"] = display_df["进店率"].apply(lambda x: f"{x:.2%}")
            display_df["购买转化率"] = display_df["购买转化率"].apply(lambda x: f"{x:.2%}")
            display_df["新客成本"] = display_df["新客成本"].apply(lambda x: f"¥{x:,.0f}")
            display_df["预估整体新客ROI"] = display_df["预估整体新客ROI"].apply(lambda x: f"{x:.2f}")

            st.dataframe(display_df, use_container_width=True, hide_index=True)

            # 导出
            buf = io.BytesIO()
            with pd.ExcelWriter(buf, engine="openpyxl") as writer:
                result_df.to_excel(writer, index=False, sheet_name="E2E指标")
            buf.seek(0)
            st.download_button(
                label="⬇ 导出计算结果 Excel",
                data=buf,
                file_name=f"E2E指标_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

# ============ 页面：报告输出 ============
elif page == "📄 报告输出":
    import re
    from openpyxl.styles import Font, PatternFill, Border, Side, Alignment

    st.markdown('<div class="page-title">报告输出</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period_filter} · 平台：{platform_filter} · 品牌：{brand_filter}</div>', unsafe_allow_html=True)

    if st.session_state["result_df"] is None:
        st.warning("请先到「🧮 计算质检」页完成计算。")
    else:
        st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
        st.markdown("#### 生成报告")

        col1, col2 = st.columns(2)
        with col1:
            platform_report = st.selectbox("平台", ["RED", "DY"])
        with col2:
            period_label = st.text_input("周期标签", value="2026.09.30")

        st.caption("新块简化为灰底 + 边框；历史数据样式/合并/公式完全保留。")

        if st.button("生成报告", type="primary"):
            try:
                result_df = st.session_state["result_df"]
                uploaded_data = st.session_state["uploaded_data"]

                # ---------- Category 映射 ----------
                cat_map = {}
                for item in uploaded_data:
                    d = item.get("df")
                    plat = item["平台"]
                    if d is None:
                        continue
                    for _, rr in d.iterrows():
                        cat = str(rr.get("Category", "")).strip()
                        if plat == "RED":
                            key = (plat, str(rr.get("Brand", "")), str(rr.get("Product", "")))
                        else:
                            key = (plat, str(rr.get("Product", "")), str(rr.get("Product", "")))
                        cat_map[key] = cat

                def get_category(plat, brand, product):
                    if (plat, brand, product) in cat_map:
                        return cat_map[(plat, brand, product)]
                    for (p, b, pr), c in cat_map.items():
                        if p == plat and pr == product:
                            return c
                    return ""

                # ---------- 数据 ----------
                if platform_report == "RED":
                    df_to_write = result_df[result_df["平台"] == "RED"].copy()
                    template_path = os.path.join(TEMPLATE_DIR, "RED_template.xlsx")
                else:
                    df_to_write = result_df[result_df["平台"] == "DY"].copy()
                    template_path = os.path.join(TEMPLATE_DIR, "DY_template.xlsx")

                if len(df_to_write) == 0:
                    st.error(f"没有 {platform_report} 数据。")
                    st.stop()

                wb = openpyxl.load_workbook(template_path)

                # ---------- 简化样式 ----------
                THIN = Side(style="thin", color="D0D0D0")
                BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
                FILL_HEADER = PatternFill("solid", fgColor="F2F2F2")
                FILL_TITLE = PatternFill("solid", fgColor="E7E6E6")
                FONT_BOLD = Font(bold=True, size=11)
                FONT_NORMAL = Font(size=11)
                ALIGN_CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
                ALIGN_LEFT = Alignment(horizontal="left", vertical="center")

                # ---------- 工具 ----------
                def snapshot_formulas(ws):
                    m = {}
                    for row in ws.iter_rows():
                        for cell in row:
                            if isinstance(cell.value, str) and cell.value.startswith("="):
                                m[(cell.row, cell.column)] = cell.value
                    return m

                pat = re.compile(r'(\$?)([A-Z]{1,3})(\$?)(\d+)')

                def shift_formulas(ws, formula_map, insert_row, n_rows):
                    def repl(m):
                        cd, col, rd, rn = m.groups()
                        rn = int(rn)
                        if rn >= insert_row:
                            rn += n_rows
                        return f"{cd}{col}{rd}{rn}"
                    for (old_r, old_c), f in formula_map.items():
                        nf = pat.sub(repl, f)
                        new_r = old_r + n_rows if old_r >= insert_row else old_r
                        ws.cell(row=new_r, column=old_c).value = nf

                def snapshot_merges(ws):
                    return [(str(m)) for m in ws.merged_cells.ranges]

                def remap_merges(ws, _unused, insert_row, n_rows):
                    for mr in list(ws.merged_cells.ranges):
                        r1 = mr.min_row
                        r2 = mr.max_row
                        if r1 >= insert_row:
                            mr.shift(row_shift=n_rows)
                        elif r2 >= insert_row:
                            mr.max_row = r2 + n_rows

                # ============================================================
                # 品牌 sheet
                # ============================================================
                brand_block_info = {}

                for brand in df_to_write["品牌"].unique().tolist():
                    if brand not in wb.sheetnames:
                        continue
                    ws_b = wb[brand]

                    old_title_row = None
                    for row in range(1, ws_b.max_row + 1):
                        for c in (1, 2, 3):
                            v = ws_b.cell(row=row, column=c).value
                            if v and "Performance Tracking" in str(v):
                                old_title_row = row
                                break
                        if old_title_row:
                            break
                    if old_title_row is None:
                        continue

                    old_header_row = None
                    for row in range(old_title_row + 1, min(old_title_row + 30, ws_b.max_row + 1)):
                        vals = [ws_b.cell(row=row, column=c).value for c in range(1, 12)]
                        if any(v == "Data Period" for v in vals if v):
                            old_header_row = row
                            break
                    if old_header_row is None:
                        continue

                    b_df = df_to_write[df_to_write["品牌"] == brand].copy()
                    n_data = len(b_df)
                    n_rows = 11 + n_data
                    insert_row = old_title_row

                    fmap = snapshot_formulas(ws_b)
                    merges = snapshot_merges(ws_b)
                    ws_b.insert_rows(insert_row, n_rows)
                    shift_formulas(ws_b, fmap, insert_row, n_rows)
                    remap_merges(ws_b, merges, insert_row, n_rows)

                    ref_h1 = insert_row + n_rows + 5
                    ref_h2 = insert_row + n_rows + 6

                    r = insert_row
                    ws_b.cell(row=r, column=1).value = f"E2E Monthly Performance Tracking _ {brand}（{period_label}）"
                    ws_b.cell(row=r, column=1).font = FONT_BOLD
                    ws_b.cell(row=r, column=1).fill = FILL_TITLE
                    r += 1
                    ws_b.cell(row=r, column=1).value = "Updated:"
                    ws_b.cell(row=r, column=1).font = FONT_BOLD
                    ws_b.cell(row=r, column=2).value = "待填写"
                    r += 1
                    ws_b.cell(row=r, column=1).value = "Data Period:"
                    ws_b.cell(row=r, column=1).font = FONT_BOLD
                    ws_b.cell(row=r, column=2).value = period_label
                    r += 1
                    r += 1
                    ws_b.cell(row=r, column=1).value = "*本次跨域计划名称包含：产品，故计算为产品合计"
                    r += 1

                    for src_h in (ref_h1, ref_h2):
                        for c in range(1, 41):
                            v = ws_b.cell(row=src_h, column=c).value
                            cell = ws_b.cell(row=r, column=c)
                            cell.value = v
                            cell.font = FONT_BOLD
                            cell.fill = FILL_HEADER
                            cell.border = BORDER
                            cell.alignment = ALIGN_CENTER
                        r += 1
                    data_start = r

                    for _, row in b_df.iterrows():
                        product = row.get("产品", "")
                        cat = get_category(platform_report, brand, product)
                        vals = {
                            1: str(row.get("项目名称", product)),
                            2: product,
                            3: cat,
                            7: float(row.get("媒体花费", 0)),
                            11: float(row.get("阅读数", 0)),
                            13: float(row.get("进店UV", 0)),
                            16: float(row.get("购买人数", 0)),
                            17: float(row.get("渠道新客数", 0)),
                        }
                        for c in range(1, 41):
                            cell = ws_b.cell(row=r, column=c)
                            cell.border = BORDER
                            cell.font = FONT_NORMAL
                            if c in vals:
                                cell.value = vals[c]
                        r += 1
                    data_end = r - 1

                    brand_block_info[brand] = {
                        "sheet": brand,
                        "data_start": data_start,
                        "data_end": data_end,
                        "categories": b_df["产品"].apply(
                            lambda p: get_category(platform_report, brand, p)
                        ).tolist(),
                    }

                # ============================================================
                # Summary
                # ============================================================
                ws = wb["Summary"]

                old_title_row = None
                for row in range(1, ws.max_row + 1):
                    v = ws.cell(row=row, column=2).value
                    if v and "Campaign起始日期" in str(v):
                        old_title_row = row
                        break
                if old_title_row is None:
                    st.error("找不到 Summary 原 Campaign 标题行。")
                    st.stop()

                n_data = len(df_to_write)
                n_ave = 4
                n_rows = 2 + n_data + n_ave

                fmap = snapshot_formulas(ws)
                merges = snapshot_merges(ws)
                ws.insert_rows(old_title_row, n_rows)
                shift_formulas(ws, fmap, old_title_row, n_rows)
                remap_merges(ws, merges, old_title_row, n_rows)

                # 新块标题（B:R 合并）
                r = old_title_row
                ws.cell(row=r, column=2).value = f"Campaign起始日期-{period_label}"
                ws.cell(row=r, column=2).font = FONT_BOLD
                for c in range(2, 19):
                    ws.cell(row=r, column=c).fill = FILL_TITLE
                    ws.cell(row=r, column=c).border = BORDER
                ws.merge_cells(start_row=r, start_column=2, end_row=r, end_column=18)
                ws.cell(row=r, column=2).alignment = ALIGN_LEFT
                ws.row_dimensions[r].height = 22
                r += 1

                # 表头
                read_label = "红书阅读数" if platform_report == "RED" else "视频观看数"
                header_map = {
                    2: "Brand",
                    3: "Media Format",
                    4: "媒体花费\n包含KOL+投流",
                    5: read_label,
                    6: "归因逻辑",
                    7: "进店UV\n（=商品页浏览）",
                    8: "进店UV vs 目标",
                    9: "进店率\n（=进店UV/阅读数)",
                    10: "购买转化率\n(=购买人数/进店数)",
                    11: "购买人数",
                    12: "渠道新客数",
                    13: "渠道新客数 vs 目标",
                    14: "渠道新客占比\n(=渠道新客/购买人数)",
                    15: "新客成本\n(=媒体花费/渠道新客数)",
                    16: "预估新客GMV\n(=客单价*新客数)",
                    17: "预估整体新客ROI\n(=新客GMV/媒体花费)",
                    18: "全店客单价",
                }
                for c in range(2, 19):
                    cell = ws.cell(row=r, column=c)
                    cell.value = header_map.get(c, "")
                    cell.font = FONT_BOLD
                    cell.fill = FILL_HEADER
                    cell.border = BORDER
                    cell.alignment = ALIGN_CENTER
                ws.row_dimensions[r].height = 45
                r += 1

                # 数据行
                for _, row in df_to_write.iterrows():
                    brand = row.get("品牌", "")
                    for c in range(2, 19):
                        cell = ws.cell(row=r, column=c)
                        cell.border = BORDER
                        cell.font = FONT_NORMAL
                    ws.cell(row=r, column=2).value = brand
                    ws.cell(row=r, column=3).value = "RED KFS" if platform_report == "RED" else "DY"
                    ws.cell(row=r, column=4).value = float(row.get("媒体花费", 0))
                    ws.cell(row=r, column=4).number_format = "#,##0"
                    ws.cell(row=r, column=5).value = float(row.get("阅读数", 0))
                    ws.cell(row=r, column=5).number_format = "#,##0"
                    ws.cell(row=r, column=7).value = float(row.get("进店UV", 0))
                    ws.cell(row=r, column=7).number_format = "#,##0"
                    ws.cell(row=r, column=9).value = f"=IFERROR(G{r}/E{r},\"-\")"
                    ws.cell(row=r, column=9).number_format = "0.0%"
                    ws.cell(row=r, column=10).value = f"=IFERROR(K{r}/G{r},\"-\")"
                    ws.cell(row=r, column=10).number_format = "0.0%"
                    ws.cell(row=r, column=11).value = float(row.get("购买人数", 0))
                    ws.cell(row=r, column=11).number_format = "#,##0"
                    ws.cell(row=r, column=12).value = float(row.get("渠道新客数", 0))
                    ws.cell(row=r, column=12).number_format = "#,##0"
                    ws.cell(row=r, column=14).value = f"=IFERROR(L{r}/K{r},\"-\")"
                    ws.cell(row=r, column=14).number_format = "0.0%"
                    ws.cell(row=r, column=15).value = f"=IFERROR(D{r}/L{r},\"-\")"
                    ws.cell(row=r, column=15).number_format = "#,##0"
                    ws.cell(row=r, column=16).value = f"=IFERROR(L{r}*R{r},\"-\")"
                    ws.cell(row=r, column=16).number_format = "#,##0"
                    ws.cell(row=r, column=17).value = f"=IFERROR(P{r}/D{r},\"-\")"
                    ws.cell(row=r, column=17).number_format = "0.0"
                    ws.cell(row=r, column=18).value = float(row.get("全店客单价", 0))
                    ws.cell(row=r, column=18).number_format = "#,##0"
                    ws.row_dimensions[r].height = 20
                    r += 1

                # Ave. 4 行（无边框）
                def _ratio(subset, num_col, den_col):
                    num_parts, den_parts = [], []
                    for b in subset:
                        info = brand_block_info.get(b)
                        if not info:
                            continue
                        ds, de = info["data_start"], info["data_end"]
                        sh = info["sheet"]
                        num_parts.append(f"'{sh}'!{num_col}{ds}:{num_col}{de}")
                        den_parts.append(f"'{sh}'!{den_col}{ds}:{den_col}{de}")
                    if not num_parts:
                        return "-"
                    return f"=IFERROR(SUM({','.join(num_parts)})/SUM({','.join(den_parts)}),\"-\")"

                skincare, makeup, fragrance, all_brands = [], [], [], []
                for brand in df_to_write["品牌"].unique():
                    info = brand_block_info.get(brand)
                    cats = set(info["categories"]) if info else set()
                    all_brands.append(brand)
                    if any(c == "Skincare" for c in cats):
                        skincare.append(brand)
                    if any(c in ("Make up", "Makeup") for c in cats):
                        makeup.append(brand)
                    if any(c == "Fragrance" for c in cats):
                        fragrance.append(brand)

                ave_defs = [
                    ("品牌组合 护肤 Ave.", skincare),
                    ("品牌组合 FRAG. Ave.", fragrance),
                    ("品牌组合 彩妆 Ave.", makeup),
                    ("品牌组合 Ave.", all_brands),
                ]
                for label, subset in ave_defs:
                    ws.cell(row=r, column=7).value = label
                    ws.cell(row=r, column=7).font = FONT_NORMAL
                    ws.cell(row=r, column=9).value = _ratio(subset, "G", "E")
                    ws.cell(row=r, column=9).number_format = "0.0%"
                    ws.cell(row=r, column=10).value = _ratio(subset, "K", "G")
                    ws.cell(row=r, column=10).number_format = "0.0%"
                    ws.cell(row=r, column=15).value = _ratio(subset, "D", "L")
                    ws.cell(row=r, column=15).number_format = "#,##0"
                    r += 1

                # 保存
                output_path = f"output_{platform_report}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                wb.save(output_path)
                with open(output_path, "rb") as f:
                    st.download_button(
                        label="⬇ 下载报告 Excel",
                        data=f,
                        file_name=f"E2E_{platform_report}_报告.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    )
                st.success(f"报告已生成：Summary {n_data} 品牌 + 4 Ave.；品牌 sheet {len(brand_block_info)} 个已更新。")

            except Exception as e:
                st.error(f"生成失败：{str(e)}")
                import traceback
                st.code(traceback.format_exc())

        st.markdown('</div>', unsafe_allow_html=True)
    
        
# ============ 页面：历史趋势 ============
elif page == "📈 历史趋势":
    st.markdown('<div class="page-title">历史数据与趋势</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period_filter} · 平台：{platform_filter} · 品牌：{brand_filter}</div>', unsafe_allow_html=True)
    st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
    st.markdown("#### 新客成本趋势")
    st.info("待接入历史数据后，这里显示趋势图。")
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
    st.markdown("#### 自动结论")
    st.info("待接入历史数据后，这里自动生成周期结论。")
    st.markdown('</div>', unsafe_allow_html=True)
