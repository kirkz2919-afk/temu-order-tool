# TEMU订单自动备货工具｜双模式稳定版 V4

import streamlit as st
import pandas as pd
from io import BytesIO
import re

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

    # 去除列名空格
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
# 备注规则
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
# 品牌识别
# ==================================================
def detect_brand(product_name):

    if pd.isna(product_name):
        return "未知品牌"

    text = str(product_name).lower()

    # 长关键词优先
    sorted_rules = sorted(
        BRAND_RULES.items(),
        key=lambda x: len(str(x[0])),
        reverse=True
    )

    for keyword, output in sorted_rules:

        keyword = str(keyword).lower().strip()

        if keyword in text:
            return str(output).strip()

    return "未知品牌"

# ==================================================
# 提取颜色 + 型号
# ==================================================
def extract_color_model(text):

    if pd.isna(text):
        return "", ""

    text = str(text).strip()

    # 统一分隔符
    text = re.sub(
        r"\s*[/\-]\s*",
        "-",
        text
    )

    # 分割
    parts = re.split(
        r"-",
        text,
        maxsplit=1
    )

    # 无分隔符
    if len(parts) < 2:
        return "", text

    color = parts[0].strip()
    model = parts[1].strip()

    return color, model

# ==================================================
# 最终型号生成
# ==================================================
def build_final_model(brand, model):

    brand = str(brand).strip()
    model = str(model).strip()

    if brand == "":
        return model

    brand_lower = brand.lower()
    model_lower = model.lower()

    # 去除型号前面的重复品牌
    pattern = rf"^{re.escape(brand_lower)}[\s\-_\/]+"

    if re.match(pattern, model_lower):

        model = re.sub(
            pattern,
            "",
            model,
            flags=re.IGNORECASE
        ).strip()

    # 最终型号
    final_model = f"{brand} {model}"

    # 清理多余空格
    final_model = re.sub(
        r"\s+",
        " ",
        final_model
    )

    return final_model.strip()

# ==================================================
# 备注识别
# ==================================================
def detect_remark(skc_code):

    # 空值默认防盗刷
    if pd.isna(skc_code):
        return "防盗刷"

    skc = str(skc_code).strip()

    # 空字符串
    if skc == "" or skc.lower() == "nan":
        return "防盗刷"

    # 规则识别
    for keyword, remark in REMARK_RULES.items():

        if keyword.lower() in skc.lower():
            return remark

    # 默认防盗刷
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
def export_excel(df, sheet_name="Sheet1"):

    output = BytesIO()

    with pd.ExcelWriter(
        output,
        engine="openpyxl"
    ) as writer:

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

            # 去除列名空格
            df.columns = df.columns.str.strip()

            all_data.append(df)

        except Exception as e:

            st.error(f"读取失败: {file.name}")

    # ==================================================
    # 合并数据
    # ==================================================
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
        result_df["订单号"] = safe_col(
            raw_df,
            "订单号"
        )

        result_df["店铺"] = safe_col(
            raw_df,
            "店铺"
        )

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

        # 转整数
        result_df["数量"] = pd.to_numeric(
            result_df["数量"],
            errors="coerce"
        ).fillna(1).astype(int)

        # ==================================================
        # 品牌识别
        # ==================================================
        result_df["品牌"] = product_name.apply(
            detect_brand
        )

        # ==================================================
        # 颜色 + 型号
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

        # ==================================================
        # 排序
        # ==================================================
        final_df = final_df.sort_values(
            by=["型号", "颜色"]
        )

        summary_df = summary_df.sort_values(
            by=["型号", "颜色"]
        )

        # ==================================================
        # 销量统计
        # ==================================================
        st.subheader("销量统计")

        stat_mode = st.selectbox(
            "统计维度",
            [
                "品牌",
                "型号",
                "型号+颜色",
                "型号+颜色+备注"
            ]
        )

        if stat_mode == "品牌":

            stat_cols = ["品牌"]

        elif stat_mode == "型号":

            stat_cols = ["型号"]

        elif stat_mode == "型号+颜色":

            stat_cols = ["型号", "颜色"]

        else:

            stat_cols = ["型号", "颜色", "备注"]

        sales_df = result_df.fillna("").groupby(
            stat_cols,
            as_index=False
        ).agg({
            "数量": "sum"
        })

        sales_df = sales_df.sort_values(
            by="数量",
            ascending=False
        )

        st.dataframe(
            sales_df,
            use_container_width=True
        )

        # ==================================================
        # 下载销量统计
        # ==================================================
        sales_excel = export_excel(
            sales_df,
            "销量统计"
        )

        st.download_button(
            label="下载销量统计Excel",
            data=sales_excel,
            file_name="销量统计.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

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

            excel_file = export_excel(
                final_df,
                "备货单"
            )

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

            # 原始字段
            pick_df["批次号"] = safe_col(raw_df, "批次号")
            pick_df["物流信息"] = safe_col(raw_df, "物流信息")
            pick_df["发货单号"] = safe_col(raw_df, "发货单号")
            pick_df["订单号"] = safe_col(raw_df, "订单号")
            pick_df["产品名称"] = safe_col(raw_df, "产品名称")
            pick_df["SKC"] = safe_col(raw_df, "SKC")
            pick_df["SKU ID"] = safe_col(raw_df, "SKU ID")

            # 自动识别字段
            pick_df["型号"] = result_df["型号"]
            pick_df["备注"] = result_df["备注"]
            pick_df["颜色"] = result_df["颜色"]

            # 发货数兼容
            if "发货数" in raw_df.columns:

                pick_df["发货数"] = raw_df["发货数"]

            elif "备货数量" in raw_df.columns:

                pick_df["发货数"] = raw_df["备货数量"]

            else:

                pick_df["发货数"] = 1

            # 转整数
            pick_df["发货数"] = pd.to_numeric(
                pick_df["发货数"],
                errors="coerce"
            ).fillna(1).astype(int)

            # 其它字段
            pick_df["收货仓库"] = safe_col(raw_df, "收货仓库")
            pick_df["店铺"] = safe_col(raw_df, "店铺")

            # ==================================================
            # 排序
            # ==================================================
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

            # ==================================================
            # 展示
            # ==================================================
            st.subheader("仓库拣货单")

            st.dataframe(
                pick_df,
                use_container_width=True
            )

            # ==================================================
            # 下载
            # ==================================================
            pick_excel = export_excel(
                pick_df,
                "仓库拣货单"
            )

            st.download_button(
                label="下载拣货单Excel",
                data=pick_excel,
                file_name="仓库拣货单.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )