import pandas as pd
import re
import os
import openpyxl
from datetime import datetime

# --- 设定你的基础路径 ---
DATA_DIR = "./data"  # 请确保这个目录下有 default_price.xlsx 和 default_template.xlsx
DEFAULT_FILE = os.path.join(DATA_DIR, "default_price.xlsx")
DEFAULT_TEMPLATE = os.path.join(DATA_DIR, "default_template.xlsx")
TEMP_OUTPUT_DIR = "./output"  # AI 生成的报价单临时存放处

if not os.path.exists(TEMP_OUTPUT_DIR):
    os.makedirs(TEMP_OUTPUT_DIR)


# --- 1. 粘贴你原本的核心数据处理函数 ---
# 在这里保留你原本 main.py 里的 load_and_parse_data, get_item_type, clean_price, fill_excel_template 等函数
# ... [此处省略这几个基础函数，直接从你的 main.py 复制过来即可] ...


def generate_quote_excel(customer: str, project: str, sku_list: list, discount: float = 35.0) -> dict:
    """
    OpenClaw 调用的主入口函数
    """
    try:
        # 1. 加载数据库
        full_data, _ = load_and_parse_data(DEFAULT_FILE)
        if full_data.empty:
            return {"status": "error", "message": "未能加载价格数据库，请检查 data 目录。"}

        new_rows = []
        not_found = []

        # 2. 遍历 AI 提取出来的 SKU 列表，执行智能匹配逻辑
        for sku in sku_list:
            sku = sku.strip().upper()
            match = full_data[full_data['SKU'] == sku]
            duration_val = 'Fixed'

            # 智能 -DD 服务周期转化逻辑
            if match.empty and re.search(r'-(12|36|60)$', sku):
                base_sku = re.sub(r'-(12|36|60)$', '-DD', sku)
                match = full_data[full_data['SKU'] == base_sku]
                if sku.endswith('-12'):
                    duration_val = '1 Yr'
                elif sku.endswith('-36'):
                    duration_val = '3 Yr'
                elif sku.endswith('-60'):
                    duration_val = '5 Yr'

            if not match.empty:
                item = match.iloc[0].to_dict()
                if item.get('Type') == 'Service' and duration_val == 'Fixed':
                    duration_val = '3 Yr'

                # 计算折后单价 (假设默认汇率 8.3，过单点 10)
                base_price = clean_price(item.get('BASE', 0))  # 简化演示，实际请用你原本的精细逻辑
                sell_price = base_price * (1 - discount / 100) * 8.3 / (1 - 10.0 / 100)

                item.update({
                    'Duration': duration_val,
                    'Discount': discount,
                    'Qty': 1,  # 默认数量1，AI 暂不处理复杂数量，如有需要可在 Schema 中扩展
                    'Unit(¥)': sell_price,
                    'Total(¥)': sell_price * 1
                })
                new_rows.append(item)
            else:
                not_found.append(sku)

        if not new_rows:
            return {"status": "error", "message": f"抱歉，您提供的型号 {', '.join(not_found)} 在库中均未找到。"}

        # 3. 组装数据并生成 Excel
        df_calc = pd.DataFrame(new_rows)
        meta = {
            'customer': customer, 'project': project, 'agent': '直客',
            's_name': "AI 报价助手", 's_phone': "", 's_email': "",
            'total_str': f"¥ {df_calc['Total(¥)'].sum():,.2f}"
        }

        excel_data = fill_excel_template(DEFAULT_TEMPLATE, df_calc, meta)

        # 4. 保存为实体文件供 OpenClaw 读取并发送
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe_cust = re.sub(r'[\\/*?:"<>|]', "", customer)
        file_name = f"Quote_{safe_cust}_{timestamp}.xlsx"
        file_path = os.path.join(TEMP_OUTPUT_DIR, file_name)

        with open(file_path, 'wb') as f:
            f.write(excel_data)

        # 5. 返回结构化结果给 AI
        response_msg = f"✅ 报价单生成成功！已匹配 {len(new_rows)} 个产品。"
        if not_found:
            response_msg += f"\n⚠️ 注意：以下型号未找到，未加入报价单：{', '.join(not_found)}"

        return {
            "status": "success",
            "message": response_msg,
            "total_amount": df_calc['Total(¥)'].sum(),
            "file_path": file_path  # OpenClaw 可以通过这个路径把文件发给用户
        }

    except Exception as e:
        return {"status": "error", "message": f"生成报价单时发生内部错误: {str(e)}"}