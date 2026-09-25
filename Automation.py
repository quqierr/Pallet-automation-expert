import streamlit as st
import pandas as pd
from collections import Counter
import base64

# === 1. Konfiguration & Styling ===
st.set_page_config(page_title="Paletten-Assistent PRO", page_icon="📦", layout="wide")

st.markdown("""
<style>
.stApp { background-color: #FFFFFF; }
div.stButton > button {
    background-color: #FFFFFF !important;
    color: #333333 !important;
    border: 1px solid #CCCCCC !important;
    border-radius: 4px !important;
    height: 3em;
    width: 100% !important;
    transition: all 0.2s;
    white-space: nowrap;
    overflow: hidden;
    text-overflow: ellipsis;
    padding-left: 10px;
    padding-right: 10px;
}
div.stButton > button:hover {
    border-color: #000000 !important;
    background-color: #F9F9F9 !important;
}
div[data-testid="stBaseButton-primary"] {
    font-weight: bold !important;
    border: 2px solid #333333 !important;
}
</style>
""", unsafe_allow_html=True)

# === 2. Daten laden ===
EXCEL_PFAD_MAIN = "Expert Automation Final v02.xlsx"
EXCEL_PFAD_STOCK = "Stock & Supply Report 20260925.xlsx"

@st.cache_data
def lade_daten(datei_pfad_main, datei_pfad_stock):
    try:
        # --- 1. 加载主表 (Models) ---
        df_main = pd.read_excel(datei_pfad_main, sheet_name="Models")
        df_main.columns = df_main.columns.str.strip() # 清理表头空格
        
        # 统一 Product code 格式
        df_main["Product code"] = df_main["Product code"].astype(str).str.strip().str.upper()

        # --- 2. 加载 SAP 库存表 ---
        # 假设数据从第 6 行开始（header=5）
        df_stock_raw = pd.read_excel(datei_pfad_stock, sheet_name="Lagerabgleich", header=5)
        df_stock_raw.columns = df_stock_raw.columns.str.strip()

        # 提取 Material (第1列) 和 Stock (第8列)
        df_stock_subset = df_stock_raw.iloc[:, [0, 7]].copy()
        df_stock_subset.columns = ["Material", "Stock_Value"]
        
        # 清理 Material 格式
        df_stock_subset["Material_Key"] = df_stock_subset["Material"].apply(
            lambda x: str(x).strip().upper().split('.')[0]
        )
        
        # 转换库存为数字
        df_stock_subset["Stock_Value"] = pd.to_numeric(df_stock_subset["Stock_Value"], errors="coerce").fillna(0)

        # --- 3. 合并数据 (Merge) ---
        df_final = pd.merge(
            df_main, 
            df_stock_subset[["Material_Key", "Stock_Value"]], 
            left_on="Product code", 
            right_on="Material_Key", 
            how="left"
        )

        df_final["Stock_Value"] = df_final["Stock_Value"].fillna(0)
        df_final["Verfügbarkeit"] = df_final["Stock_Value"].apply(lambda x: "Verfügbar" if x > 10 else "Nicht verfügbar")

        # --- 4. 创建 Mappings ---
        return {
            "p_zu_kat": dict(zip(df_final["Product code"], df_final["Category code"])),
            "p_zu_sub": dict(zip(df_final["Product code"], df_final["Sub-Categories"])),
            "p_zu_name": dict(zip(df_final["Product code"], df_final["Product name"])),
            "p_zu_preis": dict(zip(df_final["Product code"], df_final["Product price"])),
            "p_zu_stock": dict(zip(df_final["Product code"], df_final["Stock_Value"])),
            "p_zu_status": dict(zip(df_final["Product code"], df_final["Verfügbarkeit"])),
            "alle_skus": df_final["Product code"].unique().tolist()
        }

    except Exception as e:
        st.error(f"❌ 数据加载失败: {str(e)}")
        return None

# 初始化数据
data = lade_daten(EXCEL_PFAD_MAIN, EXCEL_PFAD_STOCK)

if data is not None:
    p_zu_kat = data["p_zu_kat"]
    p_zu_sub = data["p_zu_sub"]
    p_zu_name = data["p_zu_name"]
    p_zu_preis = data["p_zu_preis"]
    p_zu_status = data["p_zu_status"]
    ALLE_PRODUKTE = data["alle_skus"]
else:
    st.warning("⚠️ 无法解析数据文件，请检查 Excel 文件路径和列名是否正确。")
    st.stop()

# === 3. Palettenregeln ===
PALETTEN_REGELN = [
    {"K1": 1}, {"K2": 2}, {"KB": 2}, {"B": 6}, {"K6": 24}, {"K8": 6}, {"S": 4}, {"A": 4},
    {"T4": 4}, {"T2": 2}, {"8888": 2}, {"A": 3, "S": 1}, {"A": 2, "S": 2}, {"A": 1, "S": 3},
    {"A": 2, "B": 2, "K6": 2}, {"A": 1, "S": 1, "B": 2, "K6": 2}, {"S": 2, "B": 2, "K6": 2},
    {"A": 3, "B": 1}, {"A": 2, "S": 1, "B": 1}, {"A": 1, "S": 2, "B": 1}, {"S": 3, "B": 1},
    {"KB": 1, "B": 1, "A": 1, "K6": 1}, {"KB": 1, "B": 1, "S": 1, "K6": 1}, {"A": 2, "K8": 2},
]

def check_palette_valid(test_counts):
    if not test_counts:
        return True
    for regel in PALETTEN_REGELN:
        ist_regel_erfuellt = True
        for kat, menge in test_counts.items():
            if menge > regel.get(kat, 0):
                ist_regel_erfuellt = False
                break
        if ist_regel_erfuellt:
            return True
    return False

# === 4. Session State ===
if "palette_nr" not in st.session_state:
    st.session_state["palette_nr"] = 1
if "waren_auf_palette" not in st.session_state:
    st.session_state["waren_auf_palette"] = {}
if "verlauf" not in st.session_state:
    st.session_state["verlauf"] = []

# === 5. Header ===
def get_base64(bin_file):
    try:
        with open(bin_file, 'rb') as f:
            return base64.b64encode(f.read()).decode()
    except:
        return None

logo_base64 = get_base64("Logo Final.png")
logo_html = f'<img src="data:image/png;base64,{logo_base64}" width="270">' if logo_base64 else '<div style="font-size:24px; font-weight:bold;">[LOGO]</div>'

st.markdown(f"""
<div style="display: flex; align-items: center; gap: 30px; margin-bottom: 20px;">
    <div>{logo_html}</div>
    <div>
        <h1 style="margin: 0; font-size: 40px;">Paletten-Management</h1>
        <p style="margin: 0; font-size: 18px; color: #666666;">Play with the number ones</p>
    </div>
</div>
""", unsafe_allow_html=True)
st.divider()

# === 6. Haupt-Layout ===
aktuelle_counts = Counter()
for p, q in st.session_state["waren_auf_palette"].items():
    aktuelle_counts[p_zu_kat[p]] += q  # 修改

col1, col2 = st.columns([2, 3], gap="large")

with col1:
    st.subheader(f"Aktuelle Palette #{st.session_state['palette_nr']}")

    moegliche_produkte = [
        p for p in ALLE_PRODUKTE
        if check_palette_valid({**aktuelle_counts, p_zu_kat[p]: aktuelle_counts[p_zu_kat[p]] + 1})  # 修改
    ]

    if not moegliche_produkte:
        st.success("✅ **Palette ist optimal ausgelastet!**")
        gewaehlte_sku = None
        menge_erlaubt = False
    else:
        gewaehlte_sku = st.selectbox(
            "Produkt wählen",
            options=moegliche_produkte,
            format_func=lambda x: f"[{p_zu_sub.get(x, '')}] {x} – {p_zu_name.get(x, '')}"
        )
        menge = st.number_input("Menge", min_value=1, max_value=50, value=1)
        menge_erlaubt = False
        if gewaehlte_sku:
            if p_zu_status.get(gewaehlte_sku) == "Nicht verfügbar":
                st.error("❌ Produkt aktuell nicht verfügbar!")
            else:
                test_counts = aktuelle_counts.copy()
                test_counts[p_zu_kat[gewaehlte_sku]] += menge  # 修改
                menge_erlaubt = check_palette_valid(test_counts)
                if not menge_erlaubt:
                    st.error("❌ Kombination nicht erlaubt oder Limit überschritten!")

    b_col1, b_col2, b_col3 = st.columns([1,1,1], gap="small")
    with b_col1:
        if st.button("➕ Hinzufügen", type="primary", use_container_width=True, disabled=not (gewaehlte_sku and menge_erlaubt)):
            st.session_state["waren_auf_palette"][gewaehlte_sku] = st.session_state["waren_auf_palette"].get(gewaehlte_sku, 0) + menge
            st.rerun()
    with b_col2:
        if st.button("🗑️ Leeren", use_container_width=True):
            st.session_state["waren_auf_palette"] = {}
            st.rerun()
    with b_col3:
        if st.button("💾 Speichern", use_container_width=True, disabled=not st.session_state["waren_auf_palette"]):
            preis = sum(p_zu_preis.get(p,0)*q for p,q in st.session_state["waren_auf_palette"].items())
            total_anzahl = sum(st.session_state["waren_auf_palette"].values())
            if total_anzahl==1:
                sku=list(st.session_state["waren_auf_palette"].keys())[0]
                if p_zu_kat.get(sku)!="K1":  # 修改
                    preis += 81
            st.session_state["verlauf"].append({
                "id": st.session_state["palette_nr"],
                "items": st.session_state["waren_auf_palette"].copy(),
                "total": preis
            })
            st.session_state["waren_auf_palette"] = {}
            st.session_state["palette_nr"] += 1
            st.rerun()

with col2:
    st.subheader("Ladungsübersicht")
    if st.session_state["waren_auf_palette"]:
        df_list = []
        for p,q in st.session_state["waren_auf_palette"].items():
            df_list.append({
                "SKU": p,
                "Sub-Kategorie": p_zu_sub.get(p,""),
                "Name": p_zu_name.get(p,""),
                "Menge": f"{q} Stk.",
                "Gesamtpreis": f"{p_zu_preis.get(p,0)*q:,.2f} €"
            })
        st.dataframe(pd.DataFrame(df_list), use_container_width=True, hide_index=True)

        total_summe = sum(p_zu_preis.get(p,0)*q for p,q in st.session_state["waren_auf_palette"].items())
        if sum(st.session_state["waren_auf_palette"].values())==1:
            sku=list(st.session_state["waren_auf_palette"].keys())[0]
            if p_zu_kat.get(sku)!="K1":
                total_summe+=81
                st.info("Versandkosten für Einzelstück-Lieferung: 81,00 €")

        st.markdown(f"""
        <div style="text-align: right; padding: 10px; border-top: 2px solid #EEEEEE;">
        <span style="font-size: 16px; color: #666666;">Gesamtwert der aktuellen Palette:</span><br>
        <span style="font-size: 24px; font-weight: bold; color: #0C5CA8;">{total_summe:,.2f} €</span>
        </div>
        """, unsafe_allow_html=True)
    else:
        st.info("Die Palette ist leer.")

# === 7. Historie ===
st.divider()
st.subheader("Palettenübersicht")
for e in reversed(st.session_state["verlauf"]):
    with st.container():
        c1,c2 = st.columns(2)
        c1.write(f"**Palette #{e['id']}**")
        c2.markdown(f"<p style='text-align:right;'><b>{e['total']:,.2f} €</b></p>", unsafe_allow_html=True)
        h_df=[{
            "Sub-Kategorie":p_zu_sub.get(s),
            "Produkt":p_zu_name.get(s),
            "Menge":q
        } for s,q in e['items'].items()]
        st.dataframe(pd.DataFrame(h_df), use_container_width=True, hide_index=True)

st.write("")
gesamt_aller_paletten=sum(e["total"] for e in st.session_state["verlauf"])
st.markdown(f"""
<div style="text-align: right; padding: 20px; border-top: 3px double #EEEEEE; background-color: #F9F9F9; border-radius: 8px;">
<span style="font-size: 18px; color: #333333;">GESAMTSUMME ALLER PALETTEN:</span><br>
<span style="font-size: 32px; font-weight: bold; color: #0C5CA8;">{gesamt_aller_paletten:,.2f} €</span>
</div>
<div style="text-align: center; font-size: 13px; color: #666666; margin-top: 30px;">
Preise Stand 16.02.2026. Alle Preise sind freibleibend und unverbindlich. Änderungen und Irrtümer vorbehalten.
</div>
""", unsafe_allow_html=True)

