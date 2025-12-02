import streamlit as st
import pandas as pd
import numpy as np
import plotly.graph_objects as go
import math
import itertools
import os
import shutil
import ast
from datetime import datetime

# --- 0. 基本設定 ---
st.set_page_config(page_title="製造系統可靠性戰情室", page_icon="🏭", layout="wide", initial_sidebar_state="expanded")

# 預設 Excel 路徑
DEFAULT_EXCEL_PATH = "station_data.xlsx"

# --- 1. 全局 CSS (針對按鈕做了強力強化) ---
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;600;700&display=swap');

    .stApp {
        background: #23395B !important;
        color: #e6eef6;
        font-family: 'Inter', sans-serif;
    }
    
    .block-container {
        padding-top: 2rem !important;
        padding-bottom: 2rem !important;
    }

    /* --- 重點修改：讓 Browse files 按鈕超級顯眼 --- */
    
    /* 1. 整個上傳區域的虛線框 */
    [data-testid='stFileUploader'] {
        background-color: rgba(243, 162, 26, 0.1); 
        border: 2px dashed #f3a21a;
        border-radius: 12px;
        padding: 20px;
    }

    /* 2. 鎖定裡面的 "Browse files" 按鈕 */
    [data-testid='stFileUploader'] button {
        background-color: #f3a21a !important; /* 亮橘色實心背景 */
        color: #12223A !important;             /* 深藍色文字 */
        border: 2px solid #ffffff !important;  /* 白色邊框 */
        font-size: 20px !important;            /* 字體加大 */
        font-weight: 900 !important;           /* 特粗體 */
        padding: 12px 30px !important;         /* 按鈕尺寸加大 */
        border-radius: 10px !important;        /* 圓角 */
        cursor: pointer !important;
        transition: transform 0.2s ease, box-shadow 0.2s ease;
        box-shadow: 0 4px 10px rgba(0,0,0,0.3);
    }

    /* 3. 滑鼠移上去的效果 */
    [data-testid='stFileUploader'] button:hover {
        background-color: #ffca28 !important;  /* 變更亮 */
        transform: scale(1.05);                /* 稍微放大 */
        box-shadow: 0 0 15px rgba(243, 162, 26, 0.8); /* 發光效果 */
    }
    
    /* 4. 修改提示文字顏色 */
    [data-testid='stFileUploader'] .stMarkdown p {
        color: #ffca28 !important;
        font-size: 1.1rem !important;
    }
    
    /* ------------------------------------------- */

    /* KPI 樣式 */
    .kpi-row { display:flex; gap:18px; align-items:stretch; width:100%; }
    .kpi-box {
        flex:1; border-radius:10px; padding:18px;
        background: linear-gradient(180deg, rgba(255,255,255,0.02), rgba(255,255,255,0.01));
        box-shadow: 0 6px 18px rgba(2,8,23,0.5);
        border: 2px solid rgba(255,255,255,0.06);
        min-height:92px;
        transition: transform 0.18s ease, box-shadow 0.18s ease;
    }
    .kpi-label { color:#f3a21a; font-weight:700; font-size:18px; margin-bottom:8px; }
    .kpi-value { color:#3fe6ff; font-weight:800; font-size:26px; letter-spacing:1px; }
    
    .kpi-border-green { border-color: #4cd37a !important; }
    .kpi-border-yellow { border-color: #ffd86b !important; }
    .kpi-border-red { border-color: #ff6b6b !important; }

    /* Alert 樣式 */
    .alert-full {
        width:100%; border-radius:10px; padding:16px; margin-top:18px;
        display:flex; align-items:center; justify-content:center; gap:12px;
        border:2px solid rgba(255,255,255,0.06);
        background: rgba(255,255,255,0.03); min-height:56px;
    }
    .alert-text { font-weight:700; color:#f6d89a; }
    .alert-green { border-color: #4cd37a; background: linear-gradient(90deg, rgba(76,211,122,0.08), rgba(255,255,255,0.01)); }
    .alert-yellow { border-color: #ffd86b; background: linear-gradient(90deg, rgba(255,216,107,0.06), rgba(255,255,255,0.01)); }
    .alert-red { border-color: #ff6b6b; background: linear-gradient(90deg, rgba(255,107,107,0.06), rgba(255,255,255,0.01)); }

    /* 動畫 */
    @keyframes kpiPulse { 0% { transform: scale(1); } 50% { transform: scale(0.92); } 100% { transform: scale(1); } }
    .kpi-pulse { animation: kpiPulse 1s ease-in-out infinite; transform-origin: center; }
    
    @keyframes kpiShake {
        0% { transform: translateX(0); } 10% { transform: translateX(-10px) rotate(-1deg); }
        20% { transform: translateX(10px) rotate(1deg); } 30% { transform: translateX(-8px) rotate(-1deg); }
        40% { transform: translateX(8px) rotate(1deg); } 50% { transform: translateX(-6px) rotate(-0.5deg); }
        60% { transform: translateX(6px) rotate(0.5deg); } 70% { transform: translateX(-4px); }
        80% { transform: translateX(4px); } 90% { transform: translateX(-2px); } 100% { transform: translateX(0); }
    }
    .kpi-shake { animation: kpiShake 0.9s cubic-bezier(.36,.07,.19,.97) infinite; box-shadow: 0 18px 40px rgba(255,107,107,0.18); }

    /* Sidebar & Plotly */
    section[data-testid="stSidebar"] { background-color: #12223A !important; }
    section[data-testid="stSidebar"] label, section[data-testid="stSidebar"] .stMarkdown p { color: #f3a21a !important; font-weight: 600 !important; }
    [data-testid="stPlotlyChart"] { background-color: #ffffff; border-radius: 18px; box-shadow: 0 8px 24px rgba(0,0,0,0.20); padding: 10px; margin-bottom: 20px; }
    
    /* 變數表 */
    .var-table { width: 100%; border-collapse: collapse; background-color: rgba(255, 255, 255, 0.02); border-radius: 8px; margin-bottom: 20px; }
    .var-table th { background-color: rgba(63, 230, 255, 0.15); color: #3fe6ff; padding: 12px; border-bottom: 2px solid #3fe6ff; }
    .var-table td { padding: 12px; border-bottom: 1px solid rgba(255, 255, 255, 0.1); color: #e6eef6; }

    /* Tabs 優化 */
    .stTabs [data-baseweb="tab-list"] { gap: 10px; background-color: transparent; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; background-color: rgba(255,255,255,0.05); border-radius: 8px 8px 0 0; color: #fff; border: none; }
    .stTabs [aria-selected="true"] { background-color: #f3a21a !important; color: #12223A !important; font-weight: bold; }
    </style>
    """,
    unsafe_allow_html=True
)

# --- 2. 輔助函式與核心計算邏輯 ---

def parse_list_from_string(s):
    """解析 Excel 中的字串列表"""
    if isinstance(s, list):
        return s
    if pd.isna(s) or s == "":
        return []
    
    s = str(s).strip()
    try:
        return ast.literal_eval(s)
    except:
        try:
            return [float(x.strip()) for x in s.split(',') if x.strip()]
        except:
            return None

def get_default_data():
    """提供預設資料"""
    return pd.DataFrame([
        {"name": "工作站1", "processTime": 0.001686, "timeLimit": 10, "capacities": "[0, 700, 1400, 2100, 2800, 3500]", "probs": "[0.001, 0.003, 0.005, 0.007, 0.012, 0.972]"},
        {"name": "工作站2", "processTime": 0.010065, "timeLimit": 30, "capacities": "[0, 675, 1350, 2025, 2700, 3375]", "probs": "[0.001, 0.003, 0.005, 0.007, 0.012, 0.972]"},
        {"name": "工作站3", "processTime": 0.032278, "timeLimit": 100, "capacities": "[0, 600, 1200, 1800, 2400, 3000]", "probs": "[0.001, 0.003, 0.005, 0.007, 0.012, 0.972]"},
        {"name": "工作站4", "processTime": 0.008732, "timeLimit": 25, "capacities": "[0, 565, 1130, 1695, 2260, 2825]", "probs": "[0.001, 0.003, 0.005, 0.007, 0.012, 0.972]"},
        {"name": "工作站5", "processTime": 0.025224, "timeLimit": 70, "capacities": "[0, 540, 1080, 1620, 2160, 2700]", "probs": "[0.001, 0.003, 0.005, 0.007, 0.012, 0.972]"}
    ])

@st.cache_data
def calculate_metrics(demand, carbon_factor, working_powers, idle_powers, p, _station_data):
    n = len(_station_data)
    total_input = demand / (p ** n)
    inputs = [total_input * (p ** (i)) for i in range(n)]
    rounded_inputs = [math.ceil(x) for x in inputs]

    process_times = []
    idle_times = []
    energies = []

    for i in range(n):
        p_time = rounded_inputs[i] * _station_data[i]["processTime"]
        i_time = max(0, _station_data[i]["timeLimit"] - p_time)
        w_p = working_powers[i] if i < len(working_powers) else 2.5
        i_p = idle_powers[i] if i < len(idle_powers) else 0.5
        
        energy = (w_p * p_time) + (i_p * i_time)
        process_times.append(p_time)
        idle_times.append(i_time)
        energies.append(energy)

    total_energy = sum(energies)
    carbon_emission = total_energy * carbon_factor

    total_probability = 0
    indices_ranges = [range(len(d["capacities"])) for d in _station_data]
    
    count = 0
    for state_indices in itertools.product(*indices_ranges):
        count += 1
        if count > 50000: break 
        
        current_prob = 1.0
        valid = True
        for i, state_idx in enumerate(state_indices):
            cap = _station_data[i]["capacities"][state_idx]
            prob = _station_data[i]["probs"][state_idx]
            if cap < rounded_inputs[i]:
                valid = False
                break
            current_prob *= prob
        if valid:
            total_probability += current_prob

    return {
        "inputs": inputs,
        "rounded_inputs": rounded_inputs,
        "process_times": process_times,
        "idle_times": idle_times,
        "energies": energies,
        "total_energy": total_energy,
        "carbon_emission": carbon_emission,
        "reliability": total_probability,
        "time_max_limit": sum(d["timeLimit"] for d in _station_data),
        "total_process_time": sum(process_times),
        "total_idle_time": sum(idle_times)
    }

# --- 3. 頂部 Hero Section ---
st.markdown("""
<div style="padding:14px 10px; border-radius:10px; background: linear-gradient(90deg, rgba(6,21,39,0.6), rgba(8,30,46,0.35)); box-shadow:0 6px 18px rgba(2,8,23,0.6); margin-bottom:12px;">
  <h1 style="margin:0;color:#e6f7ff">🏭 製造系統可靠性戰情室</h1>
  <div style="color:#bcd7ea; margin-top:6px;">系統可靠度、能耗與碳排視覺化儀表板 — 含資料編輯器</div>
</div>
""", unsafe_allow_html=True)

# --- 4. 資料載入與 Session State 初始化 ---
if "df_data" not in st.session_state:
    if os.path.exists(DEFAULT_EXCEL_PATH):
        st.session_state.df_data = pd.read_excel(DEFAULT_EXCEL_PATH)
    else:
        st.session_state.df_data = get_default_data()

# --- 分頁順序 (Dashboard 在左) ---
tab_dashboard, tab_editor = st.tabs(["📊 戰情儀表板 (Dashboard)", "📝 資料管理 (Excel 編輯)"])

# --- TAB 1: 戰情儀表板 (Dashboard) ---
with tab_dashboard:
    try:
        source_df = st.session_state.df_data
        STATION_DATA = []
        
        for _, row in source_df.iterrows():
            caps = parse_list_from_string(row['capacities'])
            probs = parse_list_from_string(row['probs'])
            if caps is None: caps = []
            if probs is None: probs = []
            
            STATION_DATA.append({
                "name": str(row['name']),
                "processTime": float(row['processTime']),
                "timeLimit": float(row['timeLimit']),
                "capacities": caps,
                "probs": probs
            })
            
        FIXED_N = len(STATION_DATA)

    except Exception as e:
        st.error(f"資料讀取錯誤: {e}")
        STATION_DATA = []
        FIXED_N = 5

    if not STATION_DATA:
        st.warning("無有效工作站資料，請先至「資料管理」分頁設定。")
    else:
        # --- 側欄控制 ---
        with st.sidebar:
            st.markdown("<div style='padding:8px 6px'><h3 style='margin:0;color:#f3a21a'>系統參數面板</h3><div style='color:#cfeefb'>調整後右側即時更新</div></div>", unsafe_allow_html=True)

            demand = st.number_input("輸出量 (d)", min_value=1, value=2500, step=100)
            carbon_factor = st.number_input("CO₂ 係數 (kg/kWh)", min_value=0.001, value=0.474, step=0.001, format="%.3f")
            p_value = st.number_input("成功率 p", min_value=0.0, max_value=1.0, value=0.96, step=0.01, format="%.2f")

            st.caption("CO₂ 係數用於將能耗轉為碳排放（kg）")
            st.divider()
            with st.expander("⚡ 功率參數設定", expanded=True):
                working_powers = []
                idle_powers = []
                for i in range(FIXED_N):
                    st.write(f"**{STATION_DATA[i]['name']}**")
                    c1, c2 = st.columns([1,1])
                    working_powers.append(c1.number_input(f"加工 (kW)", value=2.89, key=f"w{i}"))
                    idle_powers.append(c2.number_input(f"閒置 (kW)", value=0.4335, key=f"i{i}"))

            st.divider()
            
            res = calculate_metrics(demand, carbon_factor, working_powers, idle_powers, p_value, STATION_DATA)
            
            if res['reliability'] < 0.8:
                st.error(f"可靠度過低：{res['reliability']:.4f}")
            else:
                st.success(f"可靠度正常：{res['reliability']:.4f}")

        # --- KPI ---
        if res['reliability'] >= 0.9:
            rd_border = "kpi-border-green"; rd_alert = "alert-green"; rd_icon = "✅"; rd_msg = "可靠度狀態優秀 (高於 0.9)"
            rd_anim = ""
        elif res['reliability'] >= 0.8:
            rd_border = "kpi-border-yellow"; rd_alert = "alert-yellow"; rd_icon = "⚠️"; rd_msg = "可靠度狀態尚可 (0.8-0.9)"
            rd_anim = "kpi-pulse"
        else:
            rd_border = "kpi-border-red"; rd_alert = "alert-red"; rd_icon = "❗"; rd_msg = "可靠度狀態危險 (低於 0.8)"
            rd_anim = "kpi-shake"

        if res['carbon_emission'] < 250:
            co2_border = "kpi-border-green"; co2_alert = "alert-green"; co2_icon = "✅"; co2_msg = "碳排放狀態正常 (低於 250kg)"
            co2_anim = ""
        elif res['carbon_emission'] <= 300:
            co2_border = "kpi-border-yellow"; co2_alert = "alert-yellow"; co2_icon = "⚠️"; co2_msg = "碳排放偏高 (250-300kg)"
            co2_anim = "kpi-pulse"
        else:
            co2_border = "kpi-border-red"; co2_alert = "alert-red"; co2_icon = "❗"; co2_msg = "碳排放過高！超過 300kg"
            co2_anim = "kpi-shake"

        st.markdown('<div class="kpi-wrapper">', unsafe_allow_html=True)
        k1, k2, k3, k4 = st.columns([1,1,1,1], gap="large")

        with k1:
            st.markdown(f'<div class="kpi-box {rd_border} {rd_anim}"><div class="kpi-label">系統可靠度 (Rd)</div><div class="kpi-value">{res["reliability"]:.4f}</div></div>', unsafe_allow_html=True)
        with k2:
            st.markdown(f'<div class="kpi-box"><div class="kpi-label">輸出量 d</div><div class="kpi-value">{demand}</div></div>', unsafe_allow_html=True)
        with k3:
            st.markdown(f'<div class="kpi-box"><div class="kpi-label">總功率 (kW)</div><div class="kpi-value">{res["total_energy"]:.3f}</div></div>', unsafe_allow_html=True)
        with k4:
            st.markdown(f'<div class="kpi-box {co2_border} {co2_anim}"><div class="kpi-label">碳排放 (kg)</div><div class="kpi-value">{res["carbon_emission"]:.3f}</div></div>', unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

        st.markdown("<div style='height:10px'></div>", unsafe_allow_html=True)
        st.markdown(f'<div class="alert-full {rd_alert}"><div class="icon">{rd_icon}</div><div class="alert-text">{rd_msg}</div></div>', unsafe_allow_html=True)
        st.markdown(f'<div class="alert-full {co2_alert}"><div class="icon">{co2_icon}</div><div class="alert-text">{co2_msg}</div></div>', unsafe_allow_html=True)

        st.divider()

        # --- 圖表 ---
        st.header("📈 數據視覺化分析")

        def layout_common(title):
            return dict(
                title=dict(text=title, x=0.5, xanchor="center", font=dict(size=18, color="#23395B", family="Inter")),
                paper_bgcolor='rgba(0,0,0,0)', plot_bgcolor='rgba(0,0,0,0)',
                margin=dict(l=40, r=20, t=55, b=40), font=dict(color="#23395B"), height=340
            )

        stations = [d["name"] for d in STATION_DATA]
        r1c1, r1c2 = st.columns([1,1], gap="large")
        r2c1, r2c2 = st.columns([1,1], gap="large")

        with r1c1:
            fig1 = go.Figure(go.Bar(x=stations, y=res["inputs"], marker_color='#60d3ff', name="輸入量"))
            fig1.update_layout(**layout_common("各工作站輸入量"))
            st.plotly_chart(fig1, use_container_width=True)

        with r1c2:
            fig2 = go.Figure()
            fig2.add_trace(go.Bar(x=stations, y=res["process_times"], name='平均加工時間 (hr)', marker_color='#35e6b0', hovertemplate='%{y:.3f} hr'))
            fig2.add_trace(go.Bar(x=stations, y=[d["timeLimit"] for d in STATION_DATA], name='時間上限 (hr)', marker_color='#ffa64d', opacity=0.95))
            fig2.update_layout(barmode='group', **layout_common("加工時間 vs 時間上限"))
            st.plotly_chart(fig2, use_container_width=True)

        with r2c1:
            colors = ['#ff6b6b' if e > 4 else '#ffd66b' if e > 2 else '#8ef0c2' for e in res["energies"]]
            fig3 = go.Figure(go.Bar(x=stations, y=res["energies"], marker_color=colors, name="能耗 (kWh)"))
            fig3.update_layout(**layout_common("功率分布"))
            st.plotly_chart(fig3, use_container_width=True)

        with r2c2:
            # 敏感度分析 (含星星標記)
            d_range = [int(x) for x in np.linspace(1000, 5500, 10)]
            d_range.sort()
            
            r_vals = []
            for d_val in d_range:
                tmp = calculate_metrics(d_val, carbon_factor, working_powers, idle_powers, p_value, STATION_DATA)
                r_vals.append(tmp['reliability'])

            fig4 = go.Figure()
            fig4.add_trace(go.Scatter(x=d_range, y=r_vals, mode='lines+markers', name='可靠度曲線', line=dict(color='#00e5ff', width=3), marker=dict(size=8)))
            
            # 臨界點 d=2592
            crit_d = 2592
            crit_res = calculate_metrics(crit_d, carbon_factor, working_powers, idle_powers, p_value, STATION_DATA)
            fig4.add_trace(go.Scatter(
                x=[crit_d], y=[crit_res['reliability']], mode='markers+text', name='臨界點 (d=2592)',
                text=['★ 臨界點'], textposition='top center',
                marker=dict(symbol='star', size=20, color='#ffd700', line=dict(color='#ff0000', width=2))
            ))

            fig4.update_layout(**layout_common("系統可靠度敏感度分析"))
            st.plotly_chart(fig4, use_container_width=True)
            
        st.header("📋 工作站狀態表")
        df_res = pd.DataFrame({
            "工作站": stations, 
            "輸入量": res["inputs"], 
            "取整輸入量": res["rounded_inputs"],
            "加工時間 (hr)": res["process_times"], 
            "閒置時間 (hr)": res["idle_times"], 
            "能耗 (kWh)": res["energies"]
        })
        
        st.dataframe(
            df_res.style.format(
                subset=["輸入量", "取整輸入量", "加工時間 (hr)", "閒置時間 (hr)", "能耗 (kWh)"],
                formatter="{:.3f}"
            ).highlight_max(subset=["能耗 (kWh)"], color='#7f1d1d'),
            use_container_width=True
        )

        # --- 數學公式 ---
        st.divider()
        st.header("🧮 數學模型與公式詳解")
        st.subheader("變數定義")
        st.markdown("""
        <table class="var-table">
          <thead><tr><th>符號</th><th>描述</th><th>單位</th></tr></thead>
          <tbody>
            <tr><td>$d$</td><td>輸出量 (需求)</td><td>單位</td></tr>
            <tr><td>$I$</td><td>系統總輸入量</td><td>單位</td></tr>
            <tr><td>$f_i^{(0)}$</td><td>工作站 $i$ 的輸入量</td><td>單位</td></tr>
          </tbody>
        </table>
        """, unsafe_allow_html=True)
        st.markdown("### 計算公式")
        st.latex(r"I = \frac{d}{p^n}")
        st.latex(r"E_{total} = \sum (P_{work} \times t_{work} + P_{idle} \times t_{idle})")

# --- TAB 2: 資料管理邏輯 ---
with tab_editor:
    st.subheader("Excel 資料編輯器")
    
    col_upload, col_settings = st.columns([2, 1])
    with col_upload:
        # 上傳按鈕的 CSS 已經在最上面設定了，這裡直接使用元件即可
        uploaded_file = st.file_uploader("📂 上傳 Excel 檔案 (若未上傳則嘗試讀取本地預設檔)", type=["xlsx"])
    
    if uploaded_file and uploaded_file.name != st.session_state.get("last_uploaded_name", ""):
        st.session_state.df_data = pd.read_excel(uploaded_file)
        st.session_state.last_uploaded_name = uploaded_file.name
        st.rerun()

    df_edit = st.session_state.df_data.copy()

    c1, c2, c3 = st.columns([1, 1, 2])
    with c1:
        time_unit = st.selectbox("ProcessTime 來源單位", ["Hour (小時)", "Minute (分鐘)"], index=0)
    
    st.markdown("---")
    
    edited_df = st.data_editor(
        df_edit,
        num_rows="dynamic",
        use_container_width=True,
        column_config={
            "name": st.column_config.TextColumn("工作站名稱", required=True),
            "processTime": st.column_config.NumberColumn(f"加工時間 ({'hr' if 'Hour' in time_unit else 'min'})", min_value=0.0, format="%.6f", required=True),
            "timeLimit": st.column_config.NumberColumn("時間上限 (hr)", min_value=0.0, required=True),
            "capacities": st.column_config.TextColumn("產能列表 (List)", help="格式: 1,2,3 或 [1,2,3]"),
            "probs": st.column_config.TextColumn("機率列表 (List)", help="格式: 0.1, 0.2... 加總需為 1")
        }
    )

    if not edited_df.equals(st.session_state.df_data):
        temp_df = edited_df.copy()
        if "Minute" in time_unit:
             temp_df['processTime'] = temp_df['processTime'] / 60.0
        st.session_state.df_data = temp_df 

    col_btn1, col_btn2 = st.columns([1, 1])
    
    with col_btn1:
        if st.button("🔍 驗證資料", type="primary", use_container_width=True):
            errors = []
            try:
                temp_df = edited_df.copy()
                for idx, row in temp_df.iterrows():
                    if row['processTime'] <= 0: errors.append(f"Row {idx+1}: processTime 必須 > 0")
                    if row['timeLimit'] < 0: errors.append(f"Row {idx+1}: timeLimit 必須 >= 0")
                    
                    caps = parse_list_from_string(row['capacities'])
                    probs = parse_list_from_string(row['probs'])
                    
                    if caps is None: errors.append(f"Row {idx+1}: capacities 格式錯誤")
                    elif not all(x < y for x, y in zip(caps, caps[1:])): errors.append(f"Row {idx+1}: capacities 必須為遞增序列")
                    
                    if probs is None: errors.append(f"Row {idx+1}: probs 格式錯誤")
                    elif not math.isclose(sum(probs), 1.0, abs_tol=1e-6): errors.append(f"Row {idx+1}: probs 加總不為 1 (目前: {sum(probs):.4f})")
                        
                    if caps and probs and len(caps) != len(probs): errors.append(f"Row {idx+1}: capacities 與 probs 長度不一致")

                if errors:
                    for err in errors: st.error(err)
                    st.session_state.validation_success = False
                else:
                    st.success("資料驗證通過！所有格式正確。")
                    st.session_state.validation_success = True
                    st.session_state.clean_df = temp_df
            
            except Exception as e:
                st.error(f"驗證過程發生未預期錯誤: {e}")

    with col_btn2:
        if st.button("💾 儲存並更新", disabled=not st.session_state.get("validation_success", False), use_container_width=True):
            try:
                save_df = st.session_state.clean_df.copy()
                if "Minute" in time_unit:
                    save_df['processTime'] = save_df['processTime'] / 60.0
                
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                if os.path.exists(DEFAULT_EXCEL_PATH):
                    backup_name = f"{os.path.splitext(DEFAULT_EXCEL_PATH)[0]}_backup_{timestamp}.xlsx"
                    shutil.copy(DEFAULT_EXCEL_PATH, backup_name)
                    st.write(f"✅ 已建立備份: `{backup_name}`")
                
                save_df.to_excel(DEFAULT_EXCEL_PATH, index=False)
                st.session_state.df_data = save_df
                st.session_state.last_save_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                st.success(f"檔案已成功儲存至 `{DEFAULT_EXCEL_PATH}`")
                st.balloons()
            except Exception as e:
                st.error(f"儲存失敗: {e}")
#在終端機輸入：python -m streamlit run "C:\Users\user\OneDrive\桌面\dashboard.py"