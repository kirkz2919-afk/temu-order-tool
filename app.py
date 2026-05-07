import streamlit as st
import pandas as pd
from io import BytesIO

st.set_page_config(page_title="TEMU自动备货工具", layout="wide")

st.title("TEMU订单自动备货工具")
st.caption("上传订单Excel → 自动生成备货单")

# 品牌规则
BRAND_RULES = {
    "Samsung": "Galaxy",
    "nubia": "ZTE NUBIA",
    "SHARP": "SHARP",
    "Xiaomi": "Xiaomi",
    "OPPO": "OPPO",
}

# 备注规则
REMARK_RULES = {
    "Gua Sheng": "防盗刷",
    "Li Ti Wen": "立体纹",
    "Li Zhi Wen": "荔枝纹",
    "Duo Gong Neng GS+JD": "多功能 挂绳+肩带",
    "Duo Gong Neng GS": "多功能 挂绳",
}

def detect_brand(product_name):
    if pd.isna(product_name):
        return "未知品牌"

    for keyword, output in BRAND_RULES.items():
        if keyword.lower() in str(product_name).lower():
            return output

    return "未知品牌"

def extract_color_model(sku_attr):
    if pd.isna(sku_attr):
        return "", ""

    text = str(sku_attr)

    if "-" not in text:
        return "", text

    parts = text.split("-", 1)

    color = parts[0].strip()
    model = parts[1].strip()

    return color, model

def build_final_model(brand, model):
    final_model = f"{brand} {model}".strip()

    words = final_model.split()

    cleaned = []

    for word in words:
        if not cleaned or cleaned[-1].lower() != word.lower():
            cleaned.append(word)

    return " ".join(cleaned)

def detect_remark(skc_code):
    if pd.isna(skc_code) or str(skc_code).strip() == "":
        return "防盗刷"

    skc = str(skc_code)

    for keyword, remark in REMARK_RULES.items():
        if keyword.lower() in skc.lower():
            return remark

    return ""

def detect_error(row):
    errors = []

    if row["品牌"] == "未知品牌":
        errors.append("未识别品牌")

    if row["颜色"] == "":
        errors.append("颜色异常")

    if row["机型"] == "":
        errors.append("机型异常")

    return " | ".join(errors)

def export_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="备货单")

    output.seek(0)

    return output

uploaded_files = st.file_uploader(
    "上传订单Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    all_data = []

    for file in uploaded_files:
        df = pd.read_excel(file)
        all_data.append(df)

    raw_df = pd.concat(all_data, ignore_index=True)

    result_df = pd.DataFrame()

    result_df["订单号"] = raw_df.iloc[:, 0]

    product_name = raw_df.iloc[:, 1]

    sku_attr = raw_df.iloc[:, 5]

    skc_code = raw_df.iloc[:, 3]

    result_df["数量"] = raw_df.iloc[:, 7]

    result_df["店铺"] = raw_df.iloc[:, 8]

    result_df["品牌"] = product_name.apply(detect_brand)

    extracted = sku_attr.apply(extract_color_model)

    result_df["颜色"] = extracted.apply(lambda x: x[0])

    result_df["机型"] = extracted.apply(lambda x: x[1])

    result_df["型号"] = result_df.apply(
        lambda row: build_final_model(row["品牌"], row["机型"]),
        axis=1
    )

    result_df["备注"] = skc_code.apply(detect_remark)

    result_df["异常"] = result_df.apply(detect_error, axis=1)

    final_df = result_df[
        [
            "型号",
            "颜色",
            "备注",
            "数量",
            "订单号",
            "店铺",
            "异常"
        ]
    ]

    summary_df = final_df.groupby(
        ["型号", "颜色", "备注"],
        as_index=False
    ).agg({
        "数量": "sum"
    })

    final_df = final_df.sort_values(by="型号")

    st.subheader("备货单")

    st.dataframe(final_df, use_container_width=True)

    st.subheader("SKU汇总")

    st.dataframe(summary_df, use_container_width=True)

    excel_file = export_excel(final_df)

    st.download_button(
        label="下载备货单Excel",
        data=excel_file,
        file_name="备货单.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )