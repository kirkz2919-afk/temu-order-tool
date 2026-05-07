# TEMU订单自动备货工具｜双模式稳定终极版 V6

import streamlit as st
import pandas as pd
from io import BytesIO
import re

# ==================================================
# 页面配置
# ==================================================
st.set_page_config(
    page_title="TEMU自动备货工具",
    layout="wide"
)

st.title("TEMU订单自动备货工具")
st.caption("上传订单Excel → 自动生成备货单 / 拣货单 / 销量统计")

# ==================================================
# 模式选择
# ==================================================
mode = st.radio(
    "选择系统模式",
    ["备货单模式", "拣货单模式"],
    horizontal=True
)

# ==================================================
# 品牌规则（外置品牌库）
# ==================================================
try:

    brand_df = pd.read_excel("brand_rules.xlsx")

    brand_df.columns = brand_df.columns.str.strip()

    BRAND_RULES = dict(
        zip(
            brand_df.iloc[:, 0],
            brand_df.iloc[:, 1]
        )
    )

except Exception as e:

    st.error(f"品牌库读取失败: {e}")

    BRAND_RULES = {}

# ==================================================
# 备注规则（仅识别SKC货号）
# ==================================================
REMARK_RULES = {
    "Gua Sheng": "防盗刷",
    "Li Ti Wen": "立体纹",
    "Li Zhi Wen": "荔枝纹",
    "Duo Gong Neng GS+JD": "多功能 挂绳+肩带",
    "Duo Gong Neng GS": "多功能 挂绳",
}

# ==================================================
# 安全取字段
# ==================================================
def safe_col(df, col_name):

    if col_name in df.columns:
        return df[col_name]

    return pd.Series(
        [""] * len(df),
        index=df.index
    )

# ==================================================
# 品牌识别（机型最后品牌优先版）
# ==================================================
def detect_brand(product_name, model_text=""):

    product_text = str(product_name).lower().strip()
    model_text = str(model_text).lower().strip()

    if len(BRAND_RULES) == 0:
        return "未知品牌"

    # ==================================================
    # 所有品牌关键词
    # ==================================================
    brand_matches = []

    for keyword, output in BRAND_RULES.items():

        keyword_clean = str(keyword).lower().strip()

        if keyword_clean == "":
            continue

        if keyword_clean in model_text:

            brand_matches.append(
                (
                    model_text.rfind(keyword_clean),
                    str(output).strip()
                )
            )

    # ==================================================
    # 机型里识别到品牌
    # 取最后出现的品牌
    # ==================================================
    if len(brand_matches) > 0:

        brand_matches.sort(
            key=lambda x: x[0]
        )

        return brand_matches[-1][1]

    # ==================================================
    # 机型没识别到
    # 才识别产品名称
    # ==================================================
    sorted_rules = sorted(
        BRAND_RULES.items(),
        key=lambda x: len(str(x[0])),
        reverse=True
    )

    for keyword, output in sorted_rules:

        keyword = str(keyword).lower().strip()

        if keyword == "":
            continue

        if keyword in product_text:
            return str(output).strip()

    return "未知品牌"

# ==================================================
# 提取颜色 + 型号
# ==================================================
def extract_color_model(text):

    if pd.isna(text):
        return "", ""

    text = str(text).strip()

    text = text.replace(" / ", "-")
    text = text.replace("/", "-")
    text = text.replace(" - ", "-")

    parts = re.split(r"-", text, maxsplit=1)

    if len(parts) < 2:
        return "", text

    color = parts[0].strip()
    model = parts[1].strip()

    return color, model

# ==================================================
# 最终型号生成（终极去重版）
# ==================================================
def build_final_model(brand, model):

    brand = str(brand).strip()
    model = str(model).strip()

    if model == "":
        return brand

    # 清理空格
    model = re.sub(
        r"\s+",
        " ",
        model
    ).strip()

    # ==================================================
    # Xiaomi Redmi → Redmi
    # ==================================================
    model = re.sub(
        r"^xiaomi\s+redmi\s+",
        "Redmi ",
        model,
        flags=re.IGNORECASE
    )

    # ==================================================
    # Samsung Galaxy → Galaxy
    # ==================================================
    model = re.sub(
        r"^samsung\s+galaxy\s+",
        "Galaxy ",
        model,
        flags=re.IGNORECASE
    )

    model = model.strip()

    brand_lower = brand.lower()
    model_lower = model.lower()

    # ==================================================
    # 型号已经自带品牌
    # 支持：
    # iPhone12
    # iPhone 12
    # iPhone-12
    # GalaxyS23
    # RedmiNote15
    # ==================================================
    pattern = rf"^{re.escape(brand_lower)}[\s\-_]*"

    if re.match(pattern, model_lower):
        return model

    # ==================================================
    # 否则拼接品牌
    # ==================================================
    final_model = f"{brand} {model}"

    final_model = re.sub(
        r"\s+",
        " ",
        final_model
    ).strip()

    return final_model

# ==================================================
# 备注识别（仅识别SKC）
# ==================================================
def detect_remark(skc_code):

    if pd.isna(skc_code):
        return "防盗刷"

    skc = str(skc_code).strip()

    if skc == "":
        return "防盗刷"

    if skc.lower() == "nan":
        return "防盗刷"

    for keyword, remark in REMARK_RULES.items():

        if keyword.lower() in skc.lower():
            return remark

    return "防盗刷"

# ==================================================
# 异常检测
# ==================================================
def detect_error(row):

    errors = []

    if row["品牌"] == "未知品牌":
        errors.append("未识别品牌")

    if row["颜色"] == "":
        errors.append("颜色异常")

    if row["机型"] == "":
        errors.append("机型异常")

    return " | ".join(errors)

# ==================================================
# Excel导出
# ==================================================
def export_excel(df_dict):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

        for sheet_name, df in df_dict.items():

            df.to_excel(
                writer,
                index=False,
                sheet_name=sheet_name
            )

    output.seek(0)

    return output

# ==================================================
# 文件上传
# ==================================================
uploaded_files = st.file_uploader(
    "上传订单Excel",
    type=["xlsx"],
    accept_multiple_files=True
)

# ==================================================
# 开始处理
# ==================================================
if uploaded_files:

    all_data = []

    for file in uploaded_files:

        try:

            df = pd.read_excel(file)

            df.columns = df.columns.str.strip()

            all_data.append(df)

        except Exception as e:

            st.error(f"读取失败: {file.name} | {e}")

    if all_data:

        raw_df = pd.concat(
            all_data,
            ignore_index=True
        )

        result_df = pd.DataFrame()

        # ==================================================
        # 产品名称
        # ==================================================
        product_name = safe_col(
            raw_df,
            "产品名称"
        )

        # ==================================================
        # SKU属性兼容
        # ==================================================
        if "SKU属性" in raw_df.columns:

            sku_attr = raw_df["SKU属性"]

        elif "SKU 属性" in raw_df.columns:

            sku_attr = raw_df["SKU 属性"]

        elif "SKU Attr" in raw_df.columns:

            sku_attr = raw_df["SKU Attr"]

        else:

            sku_attr = pd.Series(
                [""] * len(raw_df),
                index=raw_df.index
            )

        # ==================================================
        # SKC货号兼容
        # ==================================================
        if "SKC货号" in raw_df.columns:

            skc_code = raw_df["SKC货号"]

        elif "SKC" in raw_df.columns:

            skc_code = raw_df["SKC"]

        else:

            skc_code = pd.Series(
                [""] * len(raw_df),
                index=raw_df.index
            )

        # ==================================================
        # 基础字段
        # ==================================================
        result_df["订单号"] = safe_col(raw_df, "订单号")
        result_df["店铺"] = safe_col(raw_df, "店铺")

        # ==================================================
        # 数量兼容
        # ==================================================
        if "备货数量" in raw_df.columns:

            result_df["数量"] = raw_df["备货数量"]

        elif "发货数" in raw_df.columns:

            result_df["数量"] = raw_df["发货数"]

        elif "数量" in raw_df.columns:

            result_df["数量"] = raw_df["数量"]

        else:

            result_df["数量"] = 1

        result_df["数量"] = pd.to_numeric(
            result_df["数量"],
            errors="coerce"
        ).fillna(1).astype(int)

        # ==================================================
        # 品牌识别
        # ==================================================
        result_df["品牌"] = raw_df.apply(
            lambda row: detect_brand(
                row.get("产品名称", ""),
                sku_attr.loc[row.name]
            ),
            axis=1
        )

        # ==================================================
        # 提取颜色 + 型号
        # ==================================================
        extracted = sku_attr.apply(
            extract_color_model
        )

        result_df["颜色"] = extracted.apply(
            lambda x: x[0]
        )

        result_df["机型"] = extracted.apply(
            lambda x: x[1]
        )

        # ==================================================
        # 最终型号
        # ==================================================
        result_df["型号"] = result_df.apply(
            lambda row: build_final_model(
                row["品牌"],
                row["机型"]
            ),
            axis=1
        )

        # ==================================================
        # 备注
        # ==================================================
        result_df["备注"] = skc_code.apply(
            detect_remark
        )

        # ==================================================
        # 异常
        # ==================================================
        result_df["异常"] = result_df.apply(
            detect_error,
            axis=1
        )

        # ==================================================
        # 最终备货单
        # ==================================================
        final_df = result_df[[
            "品牌",
            "型号",
            "颜色",
            "备注",
            "数量",
            "订单号",
            "店铺",
            "异常"
        ]]

        # ==================================================
        # SKU汇总
        # ==================================================
        summary_df = final_df.fillna("").groupby(
            ["型号", "颜色", "备注"],
            as_index=False
        ).agg({
            "数量": "sum"
        })

        summary_df = summary_df.sort_values(
            by=["型号", "颜色"]
        )

        # ==================================================
        # 销量统计
        # ==================================================
        st.divider()
        st.subheader("销量统计")

        stat_options = st.multiselect(
            "选择统计维度",
            ["品牌", "型号", "颜色", "备注"],
            default=["型号"]
        )

        stat_df = pd.DataFrame()

        if len(stat_options) > 0:

            sales_df = result_df.copy()

            sales_df = sales_df.fillna("")

            stat_df = sales_df.groupby(
                stat_options,
                as_index=False
            ).agg({
                "数量": "sum"
            })

            stat_df = stat_df.sort_values(
                by="数量",
                ascending=False
            )

            st.subheader("销量汇总")

            st.dataframe(
                stat_df,
                use_container_width=True
            )

            st.subheader("销量可视化")

            stat_df["统计标签"] = stat_df[
                stat_options
            ].astype(str).agg(
                " | ".join,
                axis=1
            )

            chart_df = stat_df.set_index(
                "统计标签"
            )["数量"]

            st.bar_chart(chart_df)

        # ==================================================
        # 备货单模式
        # ==================================================
        if mode == "备货单模式":

            st.subheader("备货单")

            st.dataframe(
                final_df,
                use_container_width=True
            )

            st.subheader("SKU汇总")

            st.dataframe(
                summary_df,
                use_container_width=True
            )

            excel_file = export_excel({
                "备货单": final_df,
                "SKU汇总": summary_df,
                "销量统计": stat_df
            })

            st.download_button(
                label="下载备货单Excel",
                data=excel_file,
                file_name="备货单.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )

        # ==================================================
        # 拣货单模式
        # ==================================================
        if mode == "拣货单模式":

            pick_df = pd.DataFrame()

            pick_df["批次号"] = safe_col(raw_df, "批次号")
            pick_df["物流信息"] = safe_col(raw_df, "物流信息")
            pick_df["发货单号"] = safe_col(raw_df, "发货单号")
            pick_df["订单号"] = safe_col(raw_df, "订单号")
            pick_df["产品名称"] = safe_col(raw_df, "产品名称")
            pick_df["SKC"] = safe_col(raw_df, "SKC")
            pick_df["SKU ID"] = safe_col(raw_df, "SKU ID")

            pick_df["品牌"] = result_df["品牌"]
            pick_df["型号"] = result_df["型号"]
            pick_df["备注"] = result_df["备注"]
            pick_df["颜色"] = result_df["颜色"]

            if "发货数" in raw_df.columns:

                pick_df["发货数"] = raw_df["发货数"]

            elif "备货数量" in raw_df.columns:

                pick_df["发货数"] = raw_df["备货数量"]

            else:

                pick_df["发货数"] = 1

            pick_df["发货数"] = pd.to_numeric(
                pick_df["发货数"],
                errors="coerce"
            ).fillna(1).astype(int)

            pick_df["收货仓库"] = safe_col(raw_df, "收货仓库")
            pick_df["店铺"] = safe_col(raw_df, "店铺")

            sort_cols = []

            for col in [
                "店铺",
                "物流信息",
                "发货单号",
                "订单号"
            ]:

                if col in pick_df.columns:
                    sort_cols.append(col)

            if len(sort_cols) > 0:

                pick_df = pick_df.sort_values(
                    by=sort_cols,
                    ascending=True
                )

            st.subheader("仓库拣货单")

            st.dataframe(
                pick_df,
                use_container_width=True
            )

            pick_excel = export_excel({
                "仓库拣货单": pick_df,
                "销量统计": stat_df
            })

            st.download_button(
                label="下载拣货单Excel",
                data=pick_excel,
                file_name="仓库拣货单.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )