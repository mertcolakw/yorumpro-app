
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
    kategori_keywords = {
        col: cat_sheet[col]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lower()
        .tolist()
        for col in cat_sheet.columns
    }

    product_keywords = []
    prod_sheet = pd.read_excel(xls, "Ürünler Kewords")
    for _, row in prod_sheet.iterrows():
        vals = (
            row.dropna()
            .astype(str)
            .str.strip()
            .tolist()
        )
        if vals:
            product_keywords.append({
                "name": vals[0],
                "variants": [v.lower() for v in vals]
            })

    kurye_sheet = pd.read_excel(xls, "Kurye Kewords")
    gorsel_sheet = pd.read_excel(xls, "Görsel Kewords")
    ikram_sheet = pd.read_excel(xls, "İkram Ürün Kewords")

    kurye_words = (
        kurye_sheet.iloc[:, 0]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lower()
        .tolist()
    )
    gorsel_words = (
        gorsel_sheet.iloc[:, 0]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lower()
        .tolist()
    )
    ikram_words = (
        ikram_sheet.iloc[:, 0]
        .dropna()
        .astype(str)
        .str.strip()
        .str.lower()
        .tolist()
    )

    extra_keywords = {
        "kurye": kurye_words,
        "görsel": gorsel_words,
        "ikram": ikram_words,
    }

    return kategori_keywords, product_keywords, extra_keywords


def order_and_join_categories(found_categories):
    ordered = [cat for cat in CATEGORY_PRIORITY if cat in found_categories]
    return " + ".join(ordered) if ordered else ""


def find_categories(text, kategori_keywords):
    t = text.lower()
    found = []
    for kategori_adi, kelimeler in kategori_keywords.items():
        for kw in kelimeler:
            if kw and kw in t:
                found.append(kategori_adi)
                break
    return order_and_join_categories(found)


def find_products(text, product_keywords):
    t = text.lower()
    matched = []

    for product in product_keywords:
        name = product["name"]
        variants = product["variants"]
        for v in variants:
            if v and v in t:
                matched.append(name)
                break

    matched = list(dict.fromkeys(matched))

    if not matched:
        return "Genel ürün"
    return ", ".join(matched)


def find_extra(text, extra_keywords):
    t = text.lower()
    labels = []

    for label, wordlist in extra_keywords.items():
        for w in wordlist:
            if w and w in t:
                labels.append(label)
                break

    labels = list(dict.fromkeys(labels))
    return ", ".join(labels)


def build_report(df_comments, kategori_keywords, product_keywords, extra_keywords):
    if "ID" not in df_comments.columns:
        raise ValueError("Yüklediğin yorum dosyasında 'ID' sütunu yok.")
    if "Yorum" not in df_comments.columns:
        raise ValueError("Yüklediğin yorum dosyasında 'Yorum' sütunu yok.")

    rows = []
    for _, row in df_comments.iterrows():
        yorum_id = row["ID"]
        yorum_text = str(row["Yorum"])

        kategori_val = find_categories(yorum_text, kategori_keywords)
        urun_val = find_products(yorum_text, product_keywords)
        ekstra_val = find_extra(yorum_text, extra_keywords)

        rows.append({
            "ID": yorum_id,
            "Yorum": yorum_text,
            "Kategori": kategori_val,
            "Ürün": urun_val,
            "Ekstra Not": ekstra_val,
        })

    return pd.DataFrame(rows)


def summarize_kpis(df):
    total = len(df)

    all_cats = [
        c.strip()
        for x in df["Kategori"]
        for c in str(x).split("+")
        if c.strip()
    ]

    all_prods = [
        p.strip()
        for x in df["Ürün"]
        for p in str(x).split(",")
        if p.strip() and p.strip() != "Genel ürün"
    ]

    all_notes = [
        n.strip()
        for x in df["Ekstra Not"]
        for n in str(x).split(",")
        if n.strip()
    ]

    most_cat = Counter(all_cats).most_common(1)[0][0] if all_cats else "-"
    most_prod = Counter(all_prods).most_common(1)[0][0] if all_prods else "-"
    note_summary = (
        " | ".join([f"{k} {v}" for k, v in Counter(all_notes).most_common()])
        if all_notes else "-"
    )

    kpis = {
        "total": total,
        "most_cat": most_cat,
        "most_prod": most_prod,
        "note_summary": note_summary,
        "unique_cat": len(set(all_cats)),
    }

    return kpis


def make_excel_download(report_df):
    all_products = sorted({
        p.strip()
        for x in report_df["Ürün"]
        for p in str(x).split(",")
        if p and p.strip() != "Genel ürün"
    })

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        report_df.to_excel(writer, index=False, sheet_name="Rapor")
        pd.DataFrame({"Ürünler": all_products}).to_excel(
            writer,
            index=False,
            sheet_name="Ürünler"
        )
    output.seek(0)
    return output


def plot_bar(data_dict, title, color=ACCENT_COLOR):
    fig, ax = plt.subplots(figsize=(6, 4))
    keys = list(data_dict.keys())
    values = list(data_dict.values())

    ax.barh(keys, values, color=color)
    ax.set_xlabel("Adet")
    ax.set_title(title)
    plt.tight_layout()
    return fig


def plot_pie(data_dict, title):
    fig, ax = plt.subplots(figsize=(4, 4))

    if not data_dict:
        ax.text(0.5, 0.5, "Veri Yok", ha="center", va="center")
    else:
        ax.pie(
            data_dict.values(),
            labels=data_dict.keys(),
            autopct="%1.0f%%",
            colors=["#F77F00", "#FFD580", "#79C99E"]
        )

    ax.set_title(title)
    plt.tight_layout()
    return fig


def main():
    st.set_page_config(
        page_title="YorumPro",
        page_icon="YorumPro favicon.png",
        layout="wide"
    )

    try:
        st.image("YorumPro logo.png", width=220)
    except Exception:
        st.warning("Logo yüklenemedi.")

    st.markdown(
        f"""
        <div style="
            background-color:{PRIMARY_COLOR};
            padding:16px 30px;
            border-radius:10px;
            margin-bottom:20px;
            display:flex;
            justify-content:space-between;
            align-items:center;
        ">
            <h1 style="color:white;margin:0;">💬 YorumPro</h1>
            <span style="color:{ACCENT_COLOR};font-weight:600;">
                Yorumları duyar, veriye dönüştürür
            </span>
        </div>
        """,
        unsafe_allow_html=True
    )

    st.write("### 📂 Dosyaları Yükle")
    col1, col2 = st.columns(2)
    with col1:
        yorum_file = st.file_uploader(
            "Yorumlar (input.xlsx: ID | Yorum)",
            type=["xlsx"]
        )
    with col2:
        keyword_file = st.file_uploader(
            "Kurallar (Yorum Kategori Kewords.xlsx)",
            type=["xlsx"]
        )

    if yorum_file and keyword_file:
        if st.button("🚀 Analiz Et"):
            df_comments = pd.read_excel(yorum_file)
            kategori_keywords, product_keywords, extra_keywords = load_keyword_file(keyword_file)

            report_df = build_report(
                df_comments,
                kategori_keywords,
                product_keywords,
                extra_keywords,
            )

            kpi = summarize_kpis(report_df)

            st.success("Analiz tamamlandı ✅")

            k1, k2, k3, k4, k5 = st.columns(5)
            k1.metric("💬 Yorum Sayısı", kpi["total"])
            k2.metric("🥇 En Çok Kategori", kpi["most_cat"])
            k3.metric("🍔 En Çok Ürün", kpi["most_prod"])
            k4.metric("📦 Ekstra Notlar", kpi["note_summary"])
            k5.metric("🧩 Kategori Çeşitliliği", kpi["unique_cat"])

            st.divider()
            st.write("### 📊 Görsel Dashboard")

            cat_counts = Counter([
                c.strip()
                for x in report_df["Kategori"]
                for c in str(x).split("+")
                if c.strip()
            ])

            prod_counts = Counter([
                p.strip()
                for x in report_df["Ürün"]
                for p in str(x).split(",")
                if p.strip() and p.strip() != "Genel ürün"
            ])

            note_counts = Counter([
                n.strip()
                for x in report_df["Ekstra Not"]
                for n in str(x).split(",")
                if n.strip()
            ])

            colg1, colg2, colg3 = st.columns(3)

            with colg1:
                st.pyplot(
                    plot_bar(
                        dict(cat_counts.most_common(5)),
                        "En Çok Geçen 5 Kategori"
                    )
                )

            with colg2:
                st.pyplot(
                    plot_bar(
                        dict(prod_counts.most_common(5)),
                        "En Çok Geçen 5 Ürün"
                    )
                )

            with colg3:
                st.pyplot(
                    plot_pie(
                        dict(note_counts),
                        "Ekstra Not Dağılımı"
                    )
                )

            st.divider()
            st.write("### 🗂️ Detaylı Tablo")
            st.dataframe(report_df, use_container_width=True)

            excel_data = make_excel_download(report_df)
            st.download_button(
                "📥 rapor_cikis.xlsx indir",
                data=excel_data,
                file_name="rapor_cikis.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

if __name__ == "__main__":
    main()
