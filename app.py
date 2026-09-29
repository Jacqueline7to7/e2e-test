import streamlit as st

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
</style>
""", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 📊 E2E 自动化交付平台")
    st.markdown("---")
    page = st.radio(
        "导航",
        [
            "📊 概览",
            "📋 信息收集",
            "📥 取数清单",
            "📤 数据回填",
            "🧮 计算质检",
            "📄 报告输出",
            "📈 历史趋势",
        ],
        label_visibility="collapsed",
    )
    st.markdown("---")
    st.markdown("**全局筛选**")
    period = st.selectbox("周期", ["2026-08", "2026-07", "2026-06"])
    platform = st.selectbox("平台", ["全部", "RED", "DY"])
    brand = st.selectbox("品牌", ["全部", "品牌A", "品牌B", "品牌C", "品牌D", "品牌E"])
    st.markdown("---")
    st.caption("版本 v0.1")

if page == "📊 概览":
    st.markdown('<div class="page-title">E2E 交付概览</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period} · 平台：{platform} · 品牌：{brand}</div>', unsafe_allow_html=True)

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

elif page == "📋 信息收集":
    st.markdown('<div class="page-title">信息收集与校验</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period} · 平台：{platform} · 品牌：{brand}</div>', unsafe_allow_html=True)
    st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
    uploaded = st.file_uploader("上传红书/抖音信息收集表", type=["xlsx", "csv"], accept_multiple_files=True)
    if uploaded:
        st.success(f"已上传 {len(uploaded)} 个文件，校验功能待接入。")
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
    st.markdown("#### 校验结果")
    st.info("待接入校验逻辑后，这里显示每个品牌的校验状态。")
    st.markdown('</div>', unsafe_allow_html=True)

elif page == "📥 取数清单":
    st.markdown('<div class="page-title">DMP / 云图取数清单</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period} · 平台：{platform} · 品牌：{brand}</div>', unsafe_allow_html=True)
    st.info("💡 请按以下清单到 DMP 后台逐项查询，查询完成后上传导出文件或手动回填。")
    st.progress(0.75, text="取数进度：18 / 24")
    st.markdown("#### 取数任务清单")
    st.info("待接入收集表解析后，这里自动生成取数任务清单。")

elif page == "📤 数据回填":
    st.markdown('<div class="page-title">数据回填</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period} · 平台：{platform} · 品牌：{brand}</div>', unsafe_allow_html=True)
    tab1, tab2 = st.tabs(["上传文件", "手动粘贴"])
    with tab1:
        files = st.file_uploader("上传 DMP / 云图导出文件", type=["xlsx", "csv"], accept_multiple_files=True, key="upload_tab")
        if files:
            st.success(f"已上传 {len(files)} 个文件，解析功能待接入。")
    with tab2:
        pasted = st.text_area("粘贴数据（Tab 分隔）", height=200)
        if pasted:
            st.success("已接收粘贴数据，解析功能待接入。")

elif page == "🧮 计算质检":
    st.markdown('<div class="page-title">计算与质检</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period} · 平台：{platform} · 品牌：{brand}</div>', unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 1, 4])
    with col1:
        if st.button("开始计算", type="primary"):
            st.success("计算任务已触发（逻辑待接入）。")
    with col2:
        if st.button("重新计算"):
            st.info("已重新触发计算。")
    st.markdown("#### 质检结果")
    st.info("待接入计算逻辑后，这里显示每个活动的质检结果。")

elif page == "📄 报告输出":
    st.markdown('<div class="page-title">报告输出</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period} · 平台：{platform} · 品牌：{brand}</div>', unsafe_allow_html=True)
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

elif page == "📈 历史趋势":
    st.markdown('<div class="page-title">历史数据与趋势</div>', unsafe_allow_html=True)
    st.markdown(f'<div class="page-subtitle">周期：{period} · 平台：{platform} · 品牌：{brand}</div>', unsafe_allow_html=True)
    st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
    st.markdown("#### 新客成本趋势")
    st.info("待接入历史数据后，这里显示趋势图。")
    st.markdown('</div>', unsafe_allow_html=True)
    st.markdown('<div class="e2e-card">', unsafe_allow_html=True)
    st.markdown("#### 自动结论")
    st.info("待接入历史数据后，这里自动生成周期结论。")
    st.markdown('</div>', unsafe_allow_html=True)
