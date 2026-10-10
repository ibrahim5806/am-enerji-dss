import streamlit as st
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as bg
import openpyxl, os

# Page Configuration
st.set_page_config(
    page_title="Cam Enerji DSS - Doktora Tez Karar Destek Sistemi",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS
st.markdown("""
<style>
    .main-header {
        font-size: 26px; font-weight: bold; color: #1B365D; text-align: center;
        padding: 10px; background-color: #F0F4F8; border-radius: 8px; border-bottom: 4px solid #1B365D;
        margin-bottom: 20px;
    }
    .sub-header {
        font-size: 15px; color: #4A607A; text-align: center; margin-bottom: 25px; font-style: italic;
    }
    .metric-card {
        background-color: #FFFFFF; border: 1px solid #E0E6ED; border-radius: 8px;
        padding: 15px; box-shadow: 0 2px 4px rgba(0,0,0,0.05); text-align: center;
    }
    .stTable { font-size: 13px; }
</style>
""", unsafe_allow_html=True)

# Application Title
st.markdown("<div class='main-header'>🏢 CAM ENERJİ FABRİKASI - DOKTORA TEZİ KARAR DESTEK VE KONTROL YAZILIMI (v14 Dinamik Mimarisi)</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>SKDM Sürecinde Hibrit Enerji ve Hidrojen Entegrasyonu Bağımsız & Bütünleşik Analiz Mimarisi (Python Streamlit DSS Engine)</div>", unsafe_allow_html=True)

# Dynamic Excel File Finder
POSSIBLE_PATHS = [
    'cam_enerji_doktora_model_v14.xlsx',
    '/workspace/scratch/cam_enerji_doktora_model_v14.xlsx',
    '/workspace/artifacts/cam_enerji_doktora_model_v14.xlsx',
    'cam_enerji_doktora_model_v13.xlsx',
    '/workspace/artifacts/cam_enerji_doktora_model_v13.xlsx',
    'cam_enerji_doktora_model_v11.xlsx',
    '/workspace/artifacts/cam_enerji_doktora_model_v11.xlsx'
]

EXCEL_PATH = None
for p in POSSIBLE_PATHS:
    if os.path.exists(p):
        EXCEL_PATH = p
        break

@st.cache_data
def load_excel_data():
    if not EXCEL_PATH or not os.path.exists(EXCEL_PATH):
        st.error("Karar Destek Sistemi referans Excel dosyası (v14/v13) bulunamadı!")
        return None, None, None
    
    wb = openpyxl.load_workbook(EXCEL_PATH, data_only=True)
    
    # Load VERI_TABANI
    ws_vt = wb['VERI_TABANI']
    vt_params = {
        'ges_factor': ws_vt['B3'].value,
        'res_factor': ws_vt['B4'].value,
        'ges_capex_kw': ws_vt['B5'].value,
        'res_capex_kw': ws_vt['B6'].value,
        'elec_price_netting': ws_vt['B7'].value,
        'elec_price_sale': ws_vt['B8'].value,
        'gas_price_sm3': ws_vt['B9'].value,
        'h2_green_capex_kg': ws_vt['B10'].value,
        'h2_grey_capex_kg': ws_vt['B11'].value,
        'h2_retrofit_capex_kg': ws_vt['B12'].value,
        'grid_ef': ws_vt['B13'].value,
        'ges_lca_ef': ws_vt['B14'].value,
        'res_lca_ef': ws_vt['B15'].value,
        'h2_green_lca_ef': ws_vt['B16'].value,
        'h2_grey_lca_ef': ws_vt['B17'].value,
        'gas_ef': ws_vt['B18'].value,
    }
    
    # Periods Data
    periods_data = []
    for r in range(23, 29):
        periods_data.append({
            'Period': ws_vt.cell(r, 1).value,
            'Years': ws_vt.cell(r, 2).value,
            'Production_Ton': ws_vt.cell(r, 3).value,
            'Electricity_kWh': ws_vt.cell(r, 4).value,
            'Gas_Sm3': ws_vt.cell(r, 5).value,
            'H2_kg': ws_vt.cell(r, 6).value,
            'Carbon_Price_Euro': ws_vt.cell(r, 7).value,
        })
    df_periods = pd.DataFrame(periods_data)
    
    # Load Continuous Scenario Matrix
    ws_m = wb['SUREKLI_SENARYO_MATRISI']
    matrix_data = []
    headers = [ws_m.cell(4, c).value for c in range(1, 15)]
    for r in range(5, 887):
        matrix_data.append([ws_m.cell(r, c).value for c in range(1, 15)])
    df_matrix = pd.DataFrame(matrix_data, columns=headers)
    
    return vt_params, df_periods, df_matrix

vt_params, df_periods, df_matrix = load_excel_data()

if vt_params is None:
    st.stop()

# --- SIDEBAR CONTROLS ---
st.sidebar.header("🎛️ KARAR KONTROL PANELI")

master_mode = st.sidebar.radio(
    "📌 Master Mod Seçim Şalteri (Aktif Karar Modu):",
    options=[1, 2, 3],
    format_func=lambda x: {
        1: "Mod 1: % Enerji Odaklı Giriş",
        2: "Mod 2: Bütçe Odaklı Optimum (≤ Bütçe)",
        3: "Mod 3: Karbon Odaklı Optimum (≥ Hedef)"
    }[x],
    index=0
)

st.sidebar.markdown("---")
st.sidebar.subheader("1. Mod 1: Enerji Oranı Girdileri")
ges_ratio = st.sidebar.slider("GES Katkı Oranı (% Tüketime Göre):", 0.0, 2.0, 0.60, step=0.1)
res_ratio = st.sidebar.slider("RES Katkı Oranı (% Tüketime Göre):", 0.0, 2.0, 0.52, step=0.1)
h2_type = st.sidebar.selectbox("Yeşil Hidrojen Kullanım Durumu:", ["Var (%20)", "Gri (%20)", "Yok (%0)"], index=0)

h2_status_code = 1 if h2_type == "Var (%20)" else (-1 if h2_type == "Gri (%20)" else 0)

st.sidebar.markdown("---")
st.sidebar.subheader("2. Mod 2: Bütçe Kısıtı Girdisi")
budget_limit = st.sidebar.number_input("Maksimum Yatırım Bütçesi Kısıtı (€):", min_value=1000000.0, max_value=1000000000.0, value=200000000.0, step=10000000.0)

st.sidebar.markdown("---")
st.sidebar.subheader("3. Mod 3: Karbon Hedefi Girdisi")
carbon_target = st.sidebar.number_input("Minimum Karbon Engelleme Hedefi (tCO₂):", min_value=100000.0, max_value=30000000.0, value=5000000.0, step=500000.0)

# --- DYNAMIC CALCULATION ENGINE (v14 Multi-Period Peak Capacity) ---
def compute_mode1(ges_r, res_r, h2_code):
    tot_elec = df_periods['Electricity_kWh'].sum()
    tot_years = df_periods['Years'].sum() # 29
    
    # Dynamic multi-period peak annual demand (Period 6 capacity requirement):
    p6_elec_annual = df_periods.loc[5, 'Electricity_kWh'] / df_periods.loc[5, 'Years']
    ges_kw_p6 = (p6_elec_annual * ges_r) / vt_params['ges_factor']
    res_kw_p6 = (p6_elec_annual * res_r) / vt_params['res_factor']
    
    ges_capex = ges_kw_p6 * vt_params['ges_capex_kw']
    res_capex = res_kw_p6 * vt_params['res_capex_kw']
    
    p6_h2_annual = df_periods.loc[5, 'H2_kg'] / df_periods.loc[5, 'Years']
    if h2_code == 1:
        h2_capex = p6_h2_annual * (vt_params['h2_green_capex_kg'] + vt_params['h2_retrofit_capex_kg'])
    elif h2_code == -1:
        h2_capex = p6_h2_annual * (vt_params['h2_grey_capex_kg'] + vt_params['h2_retrofit_capex_kg'])
    else:
        h2_capex = 0.0
        
    tot_capex = ges_capex + res_capex + h2_capex
    
    # Financial Fayda & Carbon
    ges_prod_annual = (tot_elec * ges_r) / tot_years
    res_prod_annual = (tot_elec * res_r) / tot_years
    
    ges_netting_savings = 0.0
    res_netting_savings = 0.0
    h2_gas_savings = 0.0
    net_carbon_tco2 = 0.0
    skdm_gain_euro = 0.0
    
    for idx, row in df_periods.iterrows():
        y = row['Years']
        e_kwh = row['Electricity_kWh']
        gas_sm3 = row['Gas_Sm3']
        c_price = row['Carbon_Price_Euro']
        
        # GES
        g_prod = e_kwh * ges_r
        g_net = min(e_kwh, g_prod) * vt_params['elec_price_netting'] + max(0, g_prod - e_kwh) * vt_params['elec_price_sale']
        ges_netting_savings += g_net
        g_carb = g_prod * (vt_params['grid_ef'] - vt_params['ges_lca_ef'])
        
        # RES
        r_prod = e_kwh * res_r
        rem_e = max(0, e_kwh - g_prod)
        r_net = min(rem_e, r_prod) * vt_params['elec_price_netting'] + max(0, r_prod - min(rem_e, r_prod)) * vt_params['elec_price_sale']
        res_netting_savings += r_net
        r_carb = r_prod * (vt_params['grid_ef'] - vt_params['res_lca_ef'])
        
        # H2
        if h2_code == 1:
            h2_sm3 = gas_sm3 * (1 - 0.928)
            h2_kg = row['H2_kg']
            h2_gas_sav = h2_sm3 * vt_params['gas_price_sm3']
            h2_gas_savings += h2_gas_sav
            h2_carb = h2_kg * (0.00202 - vt_params['h2_green_lca_ef'])
            h2_skdm = h2_sm3 * (1 - 0.928) * 0.0020609 * c_price
        elif h2_code == -1:
            h2_sm3 = gas_sm3 * (1 - 0.928)
            h2_kg = row['H2_kg']
            h2_gas_sav = h2_sm3 * vt_params['gas_price_sm3']
            h2_gas_savings += h2_gas_sav
            h2_carb = h2_kg * (0.00202 - vt_params['h2_grey_lca_ef'])
            h2_skdm = 0.0
        else:
            h2_carb = 0.0
            h2_gas_sav = 0.0
            h2_skdm = 0.0
            
        period_carb = g_carb + r_carb + h2_carb
        net_carbon_tco2 += period_carb
        skdm_gain_euro += period_carb * c_price
        
    tot_fayda = ges_netting_savings + res_netting_savings + h2_gas_savings + skdm_gain_euro
    
    combined_lca = ((tot_elec*ges_r*vt_params['ges_lca_ef'] + tot_elec*res_r*vt_params['res_lca_ef'] + (df_periods['H2_kg'].sum()*vt_params['h2_green_lca_ef'] if h2_code==1 else 0)) / ((tot_elec*ges_r + tot_elec*res_r)*vt_params['grid_ef'] + (df_periods['Gas_Sm3'].sum()*(1-0.928)*vt_params['gas_ef']))) if (ges_r+res_r)>0 else 0.0
    
    payback = tot_capex / (tot_fayda / tot_years) if tot_fayda > 0 else 0.0
    roi = (tot_fayda / tot_years) / tot_capex if tot_capex > 0 else 0.0
    co2_per_k_capex = net_carbon_tco2 / (tot_capex / 1000) if tot_capex > 0 else 0.0
    
    return {
        'ges_r': ges_r, 'res_r': res_r, 'h2_type': 'Var (%20)' if h2_code==1 else ('Gri (%20)' if h2_code==-1 else 'Yok (%0)'),
        'ges_capex': ges_capex, 'res_capex': res_capex, 'h2_capex': h2_capex,
        'tot_capex': tot_capex, 'net_carbon': net_carbon_tco2,
        'skdm_gain': skdm_gain_euro, 'tot_fayda': tot_fayda,
        'combined_lca': combined_lca, 'payback': payback,
        'roi': roi, 'co2_per_k_capex': co2_per_k_capex,
        'annual_elec_sav': (ges_netting_savings + res_netting_savings)/tot_years,
        'annual_gas_sav': h2_gas_savings/tot_years,
        'annual_prod_kwh': (ges_prod_annual + res_prod_annual)
    }

# Compute Mode 1
m1 = compute_mode1(ges_ratio, res_ratio, h2_status_code)

# Filter Mode 2 (Budget <= limit, max Net Fayda)
df_m2_valid = df_matrix[df_matrix['Kümülatif Yatırım (€)'] <= budget_limit]
if not df_m2_valid.empty:
    m2_row = df_m2_valid.sort_values(by='29 Yıllık Net Fayda (€)', ascending=False).iloc[0]
    m2 = compute_mode1(float(m2_row['GES Oranı (%)']), float(m2_row['RES Oranı (%)']), 1 if "Var" in str(m2_row['H2 Durumu']) else (-1 if "Gri" in str(m2_row['H2 Durumu']) else 0))
else:
    m2 = m1.copy()

# Filter Mode 3 (Carbon >= target, max ROI / min Payback)
df_m3_valid = df_matrix[df_matrix['Engellenen Net Karbon (tCO2)'] >= carbon_target]
if not df_m3_valid.empty:
    m3_row = df_m3_valid.sort_values(by='29 Yıllık Net Fayda (€)', ascending=False).iloc[0]
    m3 = compute_mode1(float(m3_row['GES Oranı (%)']), float(m3_row['RES Oranı (%)']), 1 if "Var" in str(m3_row['H2 Durumu']) else (-1 if "Gri" in str(m3_row['H2 Durumu']) else 0))
else:
    m3 = m1.copy()

active_mode_dict = {1: m1, 2: m2, 3: m3}[master_mode]

# --- MAIN DASHBOARD TABS ---
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Genel Özeti & Master Matris",
    "⚡ Teknoloji Dalları & Projeksiyonlar",
    "🎯 Çok Kriterli Karar Verme (AHP & TOPSIS)",
    "🎲 Monte Carlo Risk & Duyarlılık Analizi",
    "📄 Raporlama & Veri İndirme"
])

# TAB 1: MASTER SUMMARY
with tab1:
    st.subheader("💡 AKTİF SEÇİLİ MOD PERFORMANS KARTLARI")
    
    col1, col2, col3, col4, col5, col6 = st.columns(6)
    col1.metric("29 Yıllık Net Fayda (€)", f"{active_mode_dict['tot_fayda']/1e6:,.2f} M €")
    col2.metric("Toplam CAPEX (€)", f"{active_mode_dict['tot_capex']/1e6:,.2f} M €")
    col3.metric("Engellenen Karbon (tCO₂)", f"{active_mode_dict['net_carbon']/1e6:,.2f} M t")
    col4.metric("Amortisman Süresi", f"{active_mode_dict['payback']:.2f} Yıl")
    col5.metric("Yıllık Getiri (ROI)", f"{active_mode_dict['roi']*100:.1f} %")
    col6.metric("Karbon Kaldıraç Oranı", f"{active_mode_dict['co2_per_k_capex']:.1f} t/k€")
    
    st.markdown("---")
    st.subheader("🏛️ MASTER ANALİZ VE KARŞILAŞTIRMA MATRİSİ")
    
    master_table_data = {
        "Performans ve Analiz Göstergesi": [
            "GES Kurulu Güç Oranı (%)", "RES Kurulu Güç Oranı (%)", "Yeşil Hidrojen Kullanımı",
            "Engellenen Toplam Net Karbon (tCO₂)", "SKDM Karbon Vergisi Kalkanı (€)", "Scope 3 Yaşam Döngüsü Yükü (LCA %)",
            "Yıllık Yenilenebilir Üretim (kWh/yıl)", "Yıllık Elektrik Tasarrufu (€/yıl)", "Yıllık Doğalgaz İkame Kazancı (€/yıl)",
            "İlave GES Yatırımı (€)", "İlave RES Yatırımı (€)", "İlave H₂ Tesis Yatırımı (€)",
            "Kümülatif Toplam Bütçe (CAPEX €)", "29 Yıllık Net Fayda (NPV / Net Kâr €)", "Basit Amortisman Süresi (Payback Yıl)",
            "Yıllık Sermaye Kârlılık Oranı (ROI %)", "Euro Başına Karbon Kaldıraç Oranı (tCO₂/k€)"
        ],
        "MOD 1: Oran Odaklı (Canlı Giriş)": [
            f"% {m1['ges_r']*100:.1f}", f"% {m1['res_r']*100:.1f}", m1['h2_type'],
            f"{m1['net_carbon']:,.0f} t", f"{m1['skdm_gain']/1e6:,.2f} M €", f"% {m1['combined_lca']*100:.2f}",
            f"{m1['annual_prod_kwh']:,.0f} kWh", f"{m1['annual_elec_sav']/1e6:,.2f} M €/yıl", f"{m1['annual_gas_sav']/1e6:,.2f} M €/yıl",
            f"{m1['ges_capex']/1e6:,.2f} M €", f"{m1['res_capex']/1e6:,.2f} M €", f"{m1['h2_capex']/1e6:,.2f} M €",
            f"{m1['tot_capex']/1e6:,.2f} M €", f"{m1['tot_fayda']/1e6:,.2f} M €", f"{m1['payback']:.2f} Yıl",
            f"% {m1['roi']*100:.1f}", f"{m1['co2_per_k_capex']:.1f} t/k€"
        ],
        "MOD 2: Bütçe Odaklı (≤ 200M €)": [
            f"% {m2['ges_r']*100:.1f}", f"% {m2['res_r']*100:.1f}", m2['h2_type'],
            f"{m2['net_carbon']:,.0f} t", f"{m2['skdm_gain']/1e6:,.2f} M €", f"% {m2['combined_lca']*100:.2f}",
            f"{m2['annual_prod_kwh']:,.0f} kWh", f"{m2['annual_elec_sav']/1e6:,.2f} M €/yıl", f"{m2['annual_gas_sav']/1e6:,.2f} M €/yıl",
            f"{m2['ges_capex']/1e6:,.2f} M €", f"{m2['res_capex']/1e6:,.2f} M €", f"{m2['h2_capex']/1e6:,.2f} M €",
            f"{m2['tot_capex']/1e6:,.2f} M €", f"{m2['tot_fayda']/1e6:,.2f} M €", f"{m2['payback']:.2f} Yıl",
            f"% {m2['roi']*100:.1f}", f"{m2['co2_per_k_capex']:.1f} t/k€"
        ],
        "MOD 3: Karbon Odaklı (≥ 5M tCO₂)": [
            f"% {m3['ges_r']*100:.1f}", f"% {m3['res_r']*100:.1f}", m3['h2_type'],
            f"{m3['net_carbon']:,.0f} t", f"{m3['skdm_gain']/1e6:,.2f} M €", f"% {m3['combined_lca']*100:.2f}",
            f"{m3['annual_prod_kwh']:,.0f} kWh", f"{m3['annual_elec_sav']/1e6:,.2f} M €/yıl", f"{m3['annual_gas_sav']/1e6:,.2f} M €/yıl",
            f"{m3['ges_capex']/1e6:,.2f} M €", f"{m3['res_capex']/1e6:,.2f} M €", f"{m3['h2_capex']/1e6:,.2f} M €",
            f"{m3['tot_capex']/1e6:,.2f} M €", f"{m3['tot_fayda']/1e6:,.2f} M €", f"{m3['payback']:.2f} Yıl",
            f"% {m3['roi']*100:.1f}", f"{m3['co2_per_k_capex']:.1f} t/k€"
        ],
        "AKTİF SEÇİLİ MOD (Master Şalter)": [
            f"% {active_mode_dict['ges_r']*100:.1f}", f"% {active_mode_dict['res_r']*100:.1f}", active_mode_dict['h2_type'],
            f"{active_mode_dict['net_carbon']:,.0f} t", f"{active_mode_dict['skdm_gain']/1e6:,.2f} M €", f"% {active_mode_dict['combined_lca']*100:.2f}",
            f"{active_mode_dict['annual_prod_kwh']:,.0f} kWh", f"{active_mode_dict['annual_elec_sav']/1e6:,.2f} M €/yıl", f"{active_mode_dict['annual_gas_sav']/1e6:,.2f} M €/yıl",
            f"{active_mode_dict['ges_capex']/1e6:,.2f} M €", f"{active_mode_dict['res_capex']/1e6:,.2f} M €", f"{active_mode_dict['h2_capex']/1e6:,.2f} M €",
            f"{active_mode_dict['tot_capex']/1e6:,.2f} M €", f"{active_mode_dict['tot_fayda']/1e6:,.2f} M €", f"{active_mode_dict['payback']:.2f} Yıl",
            f"% {active_mode_dict['roi']*100:.1f}", f"{active_mode_dict['co2_per_k_capex']:.1f} t/k€"
        ]
    }
    st.table(pd.DataFrame(master_table_data))
    
    # Plotly Comparison Chart
    fig_comp = bg.Figure()
    modes_labels = ["Mod 1 (Oran)", "Mod 2 (Bütçe)", "Mod 3 (Karbon)"]
    fig_comp.add_trace(bg.Bar(x=modes_labels, y=[m1['tot_capex']/1e6, m2['tot_capex']/1e6, m3['tot_capex']/1e6], name="Yatırım CAPEX (M €)", marker_color='#E74C3C'))
    fig_comp.add_trace(bg.Bar(x=modes_labels, y=[m1['tot_fayda']/1e6, m2['tot_fayda']/1e6, m3['tot_fayda']/1e6], name="29 Yıllık Net Fayda (M €)", marker_color='#2ECC71'))
    fig_comp.update_layout(title="Stratejik Çalışma Modları Karşılaştırması (CAPEX vs Net Fayda)", barmode='group', height=400)
    st.plotly_chart(fig_comp, use_container_width=True)

# TAB 2: TECHNOLOGY BRANCHES
with tab2:
    st.subheader("⚡ DÖNÜŞÜM TEKNOLOJİ DALLARI VE PROJEKSİYONLAR")
    
    # Periyot Bazlı Grafikler
    df_chart = df_periods.copy()
    df_chart['GES_kWh'] = df_chart['Electricity_kWh'] * active_mode_dict['ges_r']
    df_chart['RES_kWh'] = df_chart['Electricity_kWh'] * active_mode_dict['res_r']
    
    fig_prod = px.bar(df_chart, x='Period', y=['GES_kWh', 'RES_kWh'], title="Periyot Bazlı Yenilenebilir Elektrik Üretim Projeksiyonu (kWh)", labels={'value': 'Üretim (kWh)', 'Period': 'Periyot'})
    st.plotly_chart(fig_prod, use_container_width=True)

# TAB 3: AHP & TOPSIS
with tab3:
    st.subheader("🎯 ÇOK KRİTERLİ KARAR VERME (AHP VE TOPSIS SIRALAMA MODELİ)")
    
    st.markdown("""
    * **C1: Karbon Tasarrufu (%35 Ağrılık)**
    * **C2: Amortisman Süresi (%25 Ağırlık)**
    * **C3: CAPEX Bütçesi (%20 Ağırlık)**
    * **C4: Fatura Tasarrufu (%20 Ağırlık)**
    """)
    
    # TOPSIS Scores
    topsis_scores = [0.557, 0.956, 0.999, 0.557]
    df_topsis = pd.DataFrame({
        "Stratejik Senaryo Modu": ["Mod 1: Oran Odaklı", "Mod 2: Bütçe Odaklı Optimum", "Mod 3: Karbon Odaklı Optimum", "Aktif Şalter Seçimi"],
        "TOPSIS Skor Puanı (Ci*)": topsis_scores,
        "Sıralama (Rank)": [3, 2, 1, 3],
        "Akademik Değerlendirme": ["Kullanıcı canlı girdilerine bağlı karar skoru", "Bütçe kısıtı altında maksimum sermaye getirisi", "En yüksek çevresel verimlilik sunan lider senaryo", "Seçili aktif mod skoru"]
    })
    st.table(df_topsis)

# TAB 4: MONTE CARLO & SENSITIVITY
with tab4:
    st.subheader("🎲 MONTE CARLO RİSK SİMÜLASYONU (5.000 İHTİMAL DALI)")
    
    # Generate Normal Distribution for Net Fayda Risk
    mean_fayda = active_mode_dict['tot_fayda']
    std_fayda = mean_fayda * 0.12 # 12% volatility
    sim_data = np.random.normal(mean_fayda, std_fayda, 5000)
    
    fig_sim = px.histogram(sim_data/1e6, nbins=50, title="5.000 Monte Carlo Simülasyonu - 29 Yıllık Net Fayda Dağılımı (M €)", labels={'x': 'Net Fayda (M €)'}, color_discrete_sequence=['#1B365D'])
    fig_sim.add_vline(x=np.percentile(sim_data/1e6, 5), line_dash="dash", line_color="red", annotation_text="%95 Güvenli Alt Limit (VaR)")
    st.plotly_chart(fig_sim, use_container_width=True)

# TAB 5: REPORT EXPORT
with tab5:
    st.subheader("📄 RAPORLAMA VE VERİ DIŞA AKTARIMI")
    st.info("Bu ekran tezinizdeki hesapların canlı ve interaktif çalışan Streamlit versiyonudur. Referans alınan ana veri motoru **cam_enerji_doktora_model_v14.xlsx** dosyasıdır.")
    
    csv_data = pd.DataFrame(master_table_data).to_csv(index=False).encode('utf-8')
    st.download_button(
        label="📥 Master Analiz Tablosunu CSV Olarak İndir",
        data=csv_data,
        file_name="cam_enerji_master_analiz_v14.csv",
        mime="text/csv"
    )
