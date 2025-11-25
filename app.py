import os
import json
from datetime import datetime
from openpyxl import Workbook, load_workbook

# ====== 配置加载 ======
CONFIG_FILE = "config.json"

def load_config():
    """加载或创建配置文件"""
    if not os.path.exists(CONFIG_FILE):
        default_config = {
            "excel_file": "卖货登记.xlsx",
            "sheet_name": "销售记录",
            "data_start_row": 2,
            "data_end_row": 999,
            "summary_row": 1000
        }
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(default_config, f, ensure_ascii=False, indent=2)
        print(f"✅ 默认配置已生成: {CONFIG_FILE}")
        return default_config
    
    with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
        return json.load(f)

# 全局配置（程序启动时加载一次）
CONFIG = load_config()
EXCEL_FILE = CONFIG["excel_file"]
SHEET_NAME = CONFIG["sheet_name"]
DATA_START_ROW = CONFIG["data_start_row"]
DATA_END_ROW = CONFIG["data_end_row"]
SUMMARY_ROW = CONFIG["summary_row"]

# ====== 工具函数 ======
def get_today():
    return datetime.now().strftime("%Y年%m月%d日")

def _init_sheet_structure(ws):
    """初始化工作表结构（使用配置）"""
    ws.delete_rows(1, ws.max_row)
    
    headers = ["日期", "货名", "克重", "成本单价", "成本总价",
               "平台", "货源", "卖价", "退款前利润", "退款金额", "退款后利润"]
    ws.append(headers)
    
    # 预留数据区空行
    for _ in range(DATA_END_ROW - DATA_START_ROW + 1):
        ws.append([""] * 11)
    
    # 统计行
    ws.cell(row=SUMMARY_ROW, column=1, value="总计")
    ws.cell(row=SUMMARY_ROW, column=5, value=f"=SUM(E{DATA_START_ROW}:E{DATA_END_ROW})")
    ws.cell(row=SUMMARY_ROW, column=9, value=f"=SUM(I{DATA_START_ROW}:I{DATA_END_ROW})")
    ws.cell(row=SUMMARY_ROW, column=11, value=f"=SUM(K{DATA_START_ROW}:K{DATA_END_ROW})")

def safe_load_workbook(filename):
    """安全加载工作簿（自动修复缺失Sheet）"""
    if not os.path.exists(filename):
        init_template(filename, SHEET_NAME)
    
    wb = load_workbook(filename)
    if SHEET_NAME not in wb.sheetnames:
        print(f"⚠️ 工作表 '{SHEET_NAME}' 不存在，正在创建...")
        ws = wb.create_sheet(SHEET_NAME)
        _init_sheet_structure(ws)
        wb.save(filename)
        print(f"✅ 工作表 '{SHEET_NAME}' 已创建")
    return wb

def init_template(filename, sheet_name):
    """初始化模板（使用配置）"""
    print("ℹ️ 创建Excel模板...")
    wb = Workbook()
    wb.remove(wb.active)
    ws = wb.create_sheet(sheet_name)
    _init_sheet_structure(ws)
    wb.save(filename)
    print(f"✅ 模板已创建: {filename}")

def find_insert_row(ws):
    """在配置的数据区内找第一个空行"""
    for row in range(DATA_START_ROW, DATA_END_ROW + 1):
        if ws.cell(row=row, column=1).value is None:
            return row
    return None

# ====== 核心功能（add_record / process_refund 略，与之前一致，仅需替换行号逻辑）=====
# 注意：所有涉及行号的地方都使用 DATA_START_ROW, DATA_END_ROW, SUMMARY_ROW

def search_by_weight(target_weight, excel_file, sheet_name):
    wb = safe_load_workbook(excel_file)
    ws = wb[sheet_name]
    matches = []
    for row in range(DATA_START_ROW, DATA_END_ROW + 1):  # 使用配置
        cell_value = ws.cell(row=row, column=3).value
        if cell_value is not None and abs(cell_value - target_weight) < 1e-5:
            data = [ws.cell(row=row, column=i).value for i in range(1, 12)]
            matches.append((row, data))
    return matches

# add_record 和 process_refund 函数保持不变（它们只依赖全局 CONFIG 变量）
# （此处省略重复代码，实际使用时保留之前的完整实现）
