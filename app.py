import streamlit as st
import pandas as pd
import io
from datetime import datetime

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

# ============ 初始化 session_state ============
if "uploaded_data" not in st.session_state:
    st.session_state["uploaded_data"] = []

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
        results.append(("FAIL", f"缺少必填列：{', '.join(missing_cols)}"))
        return "FAIL", results
    else:
        results.append(("PASS", f"必填列完整（共 {len(required)} 列）"))

    empty_cols = []
    for c in required:
        if df[c].isnull().any() or (df[c].astype(str).str.strip() == "").any():
            empty_cols.append(c)
    if empty_cols:
        results.append(("FAIL", f"以下列存在空值：{', '.join(empty_cols)}"))
    else:
        results.append(("PASS", "所有必填列无空值"))

    spend_cols = [c for c in df.columns if "Spending" in c or "花费" in c]
    neg_cols = []
    for c in spend_cols:
        try:
            if (pd.to_numeric(df[c], errors="coerce") < 0).any():
                neg_cols.append(c)
        except Exception:
            pass
    if neg_cols:
        results.append(("FAIL", f"以下花费列存在负值：{', '.join(neg_cols)}"))
    else:
        results.append(("PASS", "花费列均为非负"))

    aov_cols = [c for c in df.columns if "客单价" in c]
    bad_aov = []
    for c in aov_cols:
        try:
            if (pd.to_numeric(df[c], errors="coerce") <= 0).any():
                bad_aov.append(c)
        except Exception:
            pass
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
    else:
        return "PASS", results

# ============ 取数清单生成函数 ============
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
                project = r.get("跨域项目名称", "")
                order = r.get("笔记ID", "")
                period = r.get("计划有效周期", "") or r.get("Campaign Period", "")
            else:
                brand = r.get("Brand", r.get("Product", ""))
                product = r.get("Product", "")
                project = r.get("星图订单ID", "")
                order = r.get("星图订单ID", "")
                period = r.get("Data Period", "")

            for dim in ["进店UV", "购买人数", "渠道新客数"]:
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
                })
                idx += 1
    return pd.DataFrame(rows)

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
    st.caption("版本 v0.3")

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
                    "文件名": f.name,
                    "平台": plat,
                    "行数": len(df),
                    "状态": status,
                    "明细": details,
                    "df": df,
                })
            except Exception as e:
                all_results.append({
                    "文件名": f.name,
                    "平台": "未知",
                    "行数": 0,
                    "状态": "FAIL",
                    "明细": [("FAIL", f"文件解析失败：{str(e)}")],
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
            st.caption("你可以直接在表格中修改「状态」列，标记每项任务的进度。")

            edited_df = st.data_editor(
                fetch_df,
                column_config={
                    "状态": st.column_config.SelectboxColumn(
                        "状态",
                        options=["待查询", "已完成", "异常"],
                        required=True,
                        default="待查询",
                    ),
                },
                use_container_width=True,
                hide_index=True,
                key="fetch_editor",
            )

            st.session_state["fetch_df"] = edited_df

            # 导出清单
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
    tab1, tab2 = st.tabs(["上传文件", "手动粘贴"])
    with tab1:
        files = st.file_uploader("上传 DMP / 云图导出文件", type=["xlsx", "csv"], accept_multiple_files=True, key="upload_tab")
        if files:
            st.success(f"已上传 {len(files)} 个文件，解析功能待接入。")
    with tab2:
        pasted = st.text_area("粘贴数据（Tab 分隔）", height=200)
        if pasted:
            st.success("已接收粘贴数据，解析功能待接入。")

# ============ 页面：计算质检 ============
elif page == "🧮 计算质检":
    st.markdown('<div class="page-title">计算与质检</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period_filter} · 平台：{platform_filter} · 品牌：{brand_filter}</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 4])
    with col1:
        if st.button("开始计算", type="primary"):
            st.success("计算任务已触发（逻辑待接入）。")
    with col2:
        if st.button("重新计算"):
            st.info("已重新触发计算。")
    st.markdown("#### 质检结果")
    st.info("待接入计算逻辑后，这里显示每个活动的质检结果。")

# ============ 页面：报告输出 ============
elif page == "📄 报告输出":
    st.markdown('<div class="page-title">报告输出</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period_filter} · 平台：{platform_filter} · 品牌：{brand_filter}</div>', unsafe_allow_html=True)
    st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
    report_type = st.radio("报告类型", ["Summary（全品牌汇总）", "品牌明细（单品牌 sheet）", "完整报告（Summary + 所有品牌 sheet）"])
    brands = st.multiselect("选择品牌", ["品牌A", "品牌B", "品牌C", "品牌D", "品牌E"], default=["品牌A", "品牌B"])
    col1, col2 = st.columns([1, 1])
    with col1:
        if st.button("生成 Excel", type="primary"):
            st.success("Excel 生成任务已触发（逻辑待接入）。")
    with col2:
        if st.button("预览"):
            st.info("预览功能待接入。")
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
