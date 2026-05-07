# TEMU订单自动备货工具｜MVP代码 V1

import streamlit as st
import pandas as pd
from io import BytesIO
import re

st.set_page_config(page_title="TEMU自动备货工具", layout="wide")

st.title("TEMU订单自动备货工具")
st.caption("上传订单Excel → 自动生成备货单 / 拣货单")

# =========================
# 模式选择
# =========================
mode = st.radio(
    "选择系统模式",
    ["备货单模式", "拣货单模式"],
    horizontal=True
)

# =========================
# 品牌规则
# =========================
brand_df = pd.read_excel("brand_rules.xlsx")

BRAND_RULES = dict(
    zip(
        brand_df["关键词"],
        brand_df["标准品牌"]
    )
)

# =========================
# 备注规则
# =========================
REMARK_RULES = {
    "Gua Sheng": "防盗刷",
    "Li Ti Wen": "立体纹",
    "Li Zhi Wen": "荔枝纹",
    "Duo Gong Neng GS+JD": "多功能 挂绳+肩带",
    "Duo Gong Neng GS": "多功能 挂绳",
}

# =========================
# 品牌识别
# =========================
def detect_brand(product_name):
    if pd.isna(product_name):
        return "未知品牌"

    for keyword, output in BRAND_RULES.items():
        if keyword.lower() in str(product_name).lower():
            return output

    return "未知品牌"

# =========================
# 提取颜色与机型
# =========================
def extract_color_model(sku_attr):

    if pd.isna(sku_attr):
        return "", ""

    text = str(sku_attr).strip()

    # 统一分隔符
    text = text.replace(" / ", "-")
    text = text.replace("/", "-")
    text = text.replace(" - ", "-")

    # 分割
    parts = text.split("-", 1)

    if len(parts) < 2:
        return "", text

    color = parts[0].strip()
    model = parts[1].strip()

    return color, model

# =========================
# 最终型号生成
# =========================
def build_final_model(brand, model):
    final_model = f"{brand} {model}".strip()

    # 去重 Galaxy Galaxy S24
    words = final_model.split()

    cleaned = []
    for word in words:
        if not cleaned or cleaned[-1].lower() != word.lower():
            cleaned.append(word)

    return " ".join(cleaned)

# =========================
# 备注识别
# =========================
def detect_remark(skc_code):
    if pd.isna(skc_code) or str(skc_code).strip() == "":
        return "防盗刷"

    skc = str(skc_code)

    for keyword, remark in REMARK_RULES.items():
        if keyword.lower() in skc.lower():
            return remark

    return ""

# =========================
# 异常检测
# =========================
def detect_error(row):
    errors = []

    if row["品牌"] == "未知品牌":
        errors.append("未识别品牌")

    if row["颜色"] == "":
        errors.append("颜色异常")

    if row["机型"] == "":
        errors.append("机型异常")

    return " | ".join(errors)

# =========================
# Excel导出
# =========================
def export_excel(df):
    output = BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="备货单")

    output.seek(0)
    return output

# =========================
# 文件上传
# =========================
uploaded_files = st.file_uploader(
    "上传订单Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:

    all_data = []

    for file in uploaded_files:
        try:
            df = pd.read_excel(file)
            all_data.append(df)
        except Exception as e:
            st.error(f"读取失败: {file.name}")

    if all_data:

        raw_df = pd.concat(all_data, ignore_index=True)

        # 字段映射
        result_df = pd.DataFrame()

        # 订单号
        result_df["订单号"] = raw_df["订单号"]

        # 产品名称
        product_name = raw_df["产品名称"]

        # SKU属性
        sku_attr = raw_df["SKU属性"]

        # SKC货号
        skc_code = raw_df["SKC货号"]

        # 数量
        result_df["数量"] = raw_df["发货数"]

        # 店铺
        result_df["店铺"] = raw_df["店铺"]

        # 品牌
        result_df["品牌"] = product_name.apply(detect_brand)

        # 颜色 + 机型
        extracted = sku_attr.apply(extract_color_model)

        result_df["颜色"] = extracted.apply(lambda x: x[0])
        result_df["机型"] = extracted.apply(lambda x: x[1])

        # 最终型号
        result_df["型号"] = result_df.apply(
            lambda row: build_final_model(row["品牌"], row["机型"]),
            axis=1
        )

        # 备注
        result_df["备注"] = skc_code.apply(detect_remark)

        # 异常
        result_df["异常"] = result_df.apply(detect_error, axis=1)

        # 最终列
        final_df = result_df[[
            "型号",
            "颜色",
            "备注",
            "数量",
            "订单号",
            "店铺",
            "异常"
        ]]

        # 汇总
        summary_df = final_df.groupby(
            ["型号", "颜色", "备注"],
            as_index=False
        ).agg({
            "数量": "sum"
        })

        # 排序
        final_df = final_df.sort_values(by="型号")
        summary_df = summary_df.sort_values(by="型号")

        # =========================
        # 备货单模式
        # =========================
        if mode == "备货单模式":

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

        # =========================
        # 拣货单模式
        # =========================
        if mode == "拣货单模式":

            pick_df = pd.DataFrame()

            # 原始字段
            pick_df["批次号"] = raw_df["批次号"]
            pick_df["物流信息"] = raw_df["物流信息"]
            pick_df["发货单号"] = raw_df["发货单号"]
            pick_df["订单号"] = raw_df["订单号"]
            pick_df["产品名称"] = raw_df["产品名称"]
            pick_df["SKC"] = raw_df["SKC"]
            pick_df["SKU ID"] = raw_df["SKU ID"]

            # 自动识别字段
            pick_df["型号"] = result_df["型号"]
            pick_df["备注"] = result_df["备注"]
            pick_df["颜色"] = result_df["颜色"]

            # 其它字段
            pick_df["发货数"] = raw_df["发货数"]
            pick_df["收货仓库"] = raw_df["收货仓库"]
            pick_df["店铺"] = raw_df["店铺"]

            # 排序逻辑
            pick_df = pick_df.sort_values(
                by=[
                    "店铺",
                    "物流信息",
                    "发货单号",
                    "订单号"
                ],
                ascending=[True, True, True, True]
            )

            st.subheader("仓库拣货单")
            st.dataframe(pick_df, use_container_width=True)

            pick_excel = export_excel(pick_df)

            st.download_button(
                label="下载拣货单Excel",
                data=pick_excel,
                file_name="仓库拣货单.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )