
import streamlit as st
import pandas as pd
import io
import matplotlib.pyplot as plt
from collections import Counter

PRIMARY_COLOR = "#0A1D3F"
ACCENT_COLOR = "#F77F00"
BG_COLOR = "#F5F5F5"

CATEGORY_PRIORITY = [
    "Hijyen & Güven", "Eksik Ürün", "Yanlış Ürün", "Teslimat",
    "Paketleme", "Lezzet", "Kalite", "Porsiyon", "Ürün Notu",
]

def load_keyword_file(xls_file):
    xls = pd.ExcelFile(xls_file)
    cat_sheet = pd.read_excel(xls, "Yorum Kategorileri Kewords")
    kategori_keywords = {col: cat_sheet[col].dropna().astype(str).str.strip().str.lower().tolist() for col in cat_sheet.columns}
    prod_sheet = pd.read_excel(xls, "Ürünler Kewords")
    product_keywords = []
    for _, row in prod_sheet.iterrows():
        vals = row.dropna().astype(str).str.strip().tolist()
        if vals:
            product_keywords.append({"name": vals[0], "variants": [v.lower() for v in vals]})
    kurye = pd.read_excel(xls, "Kurye Kewords").iloc[:,0].dropna().astype(str).str.lower().tolist()
    gorsel = pd.read_excel(xls, "Görsel Kewords").iloc[:,0].dropna().astype(str).str.lower().tolist()
    ikram = pd.read_excel(xls, "İkram Ürün Kewords").iloc[:,0].dropna().astype(str).str.lower().tolist()
    return kategori_keywords, product_keywords, {"kurye":kurye, "görsel":gorsel, "ikram":ikram}

def order_and_join_categories(found):
    ordered = [cat for cat in CATEGORY_PRIORITY if cat in found]
    return " + ".join(ordered) if ordered else ""

def find_categories(text, keywords):
    t = text.lower()
    found = [k for k, ws in keywords.items() if any(w in t for w in ws)]
    return order_and_join_categories(found)

def find_products(text, products):
    t = text.lower()
    matched = [p["name"] for p in products if any(v in t for v in p["variants"])]
    matched = list(dict.fromkeys(matched))
    return ", ".join(matched) if matched else "Genel ürün"

def find_extra(text, extras):
    t = text.lower()
    labels = [label for label, ws in extras.items() if any(w in t for w in ws)]
    labels = list(dict.fromkeys(labels))
    return ", ".join(labels)

def build_report(df, cats, prods, extras):
    rows=[]
    for _,r in df.iterrows():
        rows.append({
            "ID":r["ID"],
            "Yorum":r["Yorum"],
            "Kategori":find_categories(str(r["Yorum"]), cats),
            "Ürün":find_products(str(r["Yorum"]), prods),
            "Ekstra Not":find_extra(str(r["Yorum"]), extras)
        })
    return pd.DataFrame(rows)

def summarize_kpis(df):
    total = len(df)
    all_cats = [c.strip() for x in df["Kategori"] for c in str(x).split("+") if c.strip()]
    all_prods = [p.strip() for x in df["Ürün"] for p in str(x).split(",") if p.strip()!="Genel ürün"]
    all_notes = [n.strip() for x in df["Ekstra Not"] for n in str(x).split(",") if n.strip()]
    most_cat = Counter(all_cats).most_common(1)[0][0] if all_cats else "-"
    most_prod = Counter(all_prods).most_common(1)[0][0] if all_prods else "-"
    notesum = " | ".join([f"{k} {v}" for k,v in Counter(all_notes).most_common()]) if all_notes else "-"
    return {"total":total,"most_cat":most_cat,"most_prod":most_prod,"notesum":notesum,"unique_cat":len(set(all_cats))}

def make_excel_download(df):
    all_prods = sorted({p.strip() for x in df["Ürün"] for p in str(x).split(",") if p and p!="Genel ürün"})
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="Rapor")
        pd.DataFrame({"Ürünler":all_prods}).to_excel(w,index=False,sheet_name="Ürünler")
    out.seek(0)
    return out

def plot_bar(data_dict, title, color=ACCENT_COLOR):
    fig, ax = plt.subplots(figsize=(6,4))
    keys = list(data_dict.keys())
    values = list(data_dict.values())
    ax.barh(keys, values, color=color)
    ax.set_xlabel("Adet")
    ax.set_title(title)
    plt.tight_layout()
    return fig

def plot_pie(data_dict, title):
    fig, ax = plt.subplots(figsize=(4,4))
    if not data_dict:
        ax.text(0.5,0.5,"Veri Yok",ha="center",va="center")
    else:
        ax.pie(data_dict.values(), labels=data_dict.keys(), autopct="%1.0f%%", colors=["#F77F00","#FFD580","#79C99E"])
    ax.set_title(title)
    plt.tight_layout()
    return fig

def main():
    st.set_page_config(
        page_title="YorumPro",
        page_icon="YorumPro favicon.png",
        layout="wide"
    )

    st.image("YorumPro logo.png", width=220)
    st.markdown(
        f"""
        <div style="background-color:{PRIMARY_COLOR};padding:16px 30px;border-radius:10px;margin-bottom:20px;display:flex;justify-content:space-between;align-items:center;">
            <h1 style="color:white;margin:0;">💬 YorumPro</h1>
            <span style="color:{ACCENT_COLOR};font-weight:600;">Yorumları duyar, veriye dönüştürür</span>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("### 📂 Dosyaları Yükle")
    col1, col2 = st.columns(2)
    with col1:
        yorum_file = st.file_uploader("Yorumlar (input.xlsx: ID | Yorum)", type=["xlsx"])
    with col2:
        keyword_file = st.file_uploader("Kurallar (Yorum Kategori Kewords.xlsx)", type=["xlsx"])

    if yorum_file and keyword_file:
        if st.button("🚀 Analiz Et"):
            df_comments = pd.read_excel(yorum_file)
            cats, prods, extras = load_keyword_file(keyword_file)
            report = build_report(df_comments, cats, prods, extras)
            kpi = summarize_kpis(report)

            st.success("Analiz tamamlandı ✅")

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("💬 Yorum Sayısı", kpi["total"])
            k2.metric("🥇 En Çok Kategori", kpi["most_cat"])
            k3.metric("🍔 En Çok Ürün", kpi["most_prod"])
            k4.metric("📦 Ekstra Notlar", kpi["notesum"])
            k5.metric("🧩 Kategori Çeşitliliği", kpi["unique_cat"])

            st.divider()
            st.write("### 📊 Görsel Dashboard")

            from collections import Counter
            cat_counts = Counter([c.strip() for x in report["Kategori"] for c in str(x).split("+") if c.strip()])
            prod_counts = Counter([p.strip() for x in report["Ürün"] for p in str(x).split(",") if p.strip()!="Genel ürün"])
            note_counts = Counter([n.strip() for x in report["Ekstra Not"] for n in str(x).split(",") if n.strip()])

            colg1, colg2, colg3 = st.columns(3)
            with colg1:
                st.pyplot(plot_bar(dict(cat_counts.most_common(5)), "En Çok Geçen 5 Kategori"))
            with colg2:
                st.pyplot(plot_bar(dict(prod_counts.most_common(5)), "En Çok Geçen 5 Ürün"))
            with colg3:
                st.pyplot(plot_pie(dict(note_counts), "Ekstra Not Dağılımı"))

            st.divider()
            st.write("### 🗂️ Detaylı Tablo")
            st.dataframe(report, use_container_width=True)

            excel_data = make_excel_download(report)
            st.download_button(
                "📥 rapor_cikis.xlsx indir",
                data=excel_data,
                file_name="rapor_cikis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
