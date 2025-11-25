# -*- coding: utf-8 -*-
import os
from datetime import datetime
from openpyxl import Workbook, load_workbook

EXCEL_FILE = "卖货登记.xlsx"
SHEET_NAME = "销售记录"

def get_today():
    return datetime.now().strftime("%Y年%m月%d日")

def safe_load_workbook(filename):
    """安全加载工作簿（处理文件不存在或被占用）"""
    if not os.path.exists(filename):
        init_template(filename, SHEET_NAME)
    return load_workbook(filename)

def init_template(filename, sheet_name):
    """初始化Excel模板（含表头和统计行）"""
    print("ℹ️ 首次运行，正在创建Excel模板...")
    wb = Workbook()
    ws = wb.active
    ws.title = sheet_name
    
    # 表头
    headers = ["日期", "货名", "克重", "成本单价", "成本总价",
               "平台", "货源", "卖价", "退款前利润", "退款金额", "退款后利润"]
    ws.append(headers)
    
    # 预留998行数据区（第2~999行），第1000行为统计行
    for _ in range(998):
        ws.append([""] * 11)
    
    # 第1000行：统计公式
    ws.cell(row=1000, column=1, value="总计")
    ws.cell(row=1000, column=5, value="=SUM(E2:E999)")   # 总成本
    ws.cell(row=1000, column=9, value="=SUM(I2:I999)")   # 总退款前利润
    ws.cell(row=1000, column=11, value="=SUM(K2:K999)")  # 总退款后利润
    
    wb.save(filename)
    print(f"✅ 模板已创建: {filename}")

def find_insert_row(ws):
    """找到第一个A列为空的行（从第2行开始）"""
    for row in range(2, 1000):  # 限制在数据区（2~999行）
        if ws.cell(row=row, column=1).value is None:
            return row
    return None  # 数据区已满

def add_record(excel_file, sheet_name):
    """新增销售记录（顺序追加 + 智能公式 + 横向回显）"""
    print("\n【新增销售记录】")
    try:
        goods = input("货名: ").strip()
        weight = float(input("克重 (纯数字): "))
        cost = float(input("成本单价 (纯数字): "))
        platform = input("平台: ").strip()
        source = input("货源: ").strip()
        sell_price = float(input("卖价 (纯数字): "))
    except ValueError:
        print("❌ 输入错误！请确保克重、成本单价、卖价为数字")
        return

    total_cost = weight * cost
    profit_before = sell_price - total_cost

    wb = safe_load_workbook(excel_file)
    ws = wb[sheet_name]
    
    insert_row = find_insert_row(ws)
    if insert_row is None:
        print("❌ 数据区已满（最多998条记录）！")
        return

    print(f"ℹ️ 新记录将添加在第{insert_row}行")
    
    # 写入带公式的完整数据
    data = [
        get_today(), goods, weight, cost,
        f"=C{insert_row}*D{insert_row}",  # E: 成本总价
        platform, source, sell_price,
        f"=H{insert_row}-E{insert_row}",  # I: 退款前利润
        "",  # J: 退款金额（初始空）
        f"=IF(J{insert_row}=\"\", MAX(0,H{insert_row}-E{insert_row}), MAX(0,H{insert_row}-E{insert_row}-J{insert_row}))"  # K: 智能公式
    ]
    
    for col_idx, value in enumerate(data, start=1):
        ws.cell(row=insert_row, column=col_idx, value=value)
    
    wb.save(excel_file)
    
    # ====== 横向回显 ======
    headers = ["日期", "货名", "克重", "成本单价", "成本总价",
               "平台", "货源", "卖价", "退款前利润", "退款金额", "退款后利润"]
    display_values = [
        get_today(), goods, f"{weight:.2f}", f"{cost:.2f}", f"{total_cost:.2f}",
        platform, source, f"{sell_price:.2f}", f"{profit_before:.2f}", "", f"{max(0, profit_before):.2f}"
    ]
    
    print("\n✅ 记录已成功添加！完整数据如下：")
    print("=" * 120)
    print("".join([f"{h:>10}" for h in headers]))
    print("".join([f"{str(v):>10}" for v in display_values]))
    print("=" * 120)

def search_by_weight(target_weight, excel_file, sheet_name):
    """按克重搜索记录（返回 [(行号, 数据), ...]）"""
    wb = safe_load_workbook(excel_file)
    ws = wb[sheet_name]
    matches = []
    for row in range(2, 1000):
        cell_value = ws.cell(row=row, column=3).value  # C列：克重
        if cell_value is not None and abs(cell_value - target_weight) < 1e-5:
            data = [ws.cell(row=row, column=i).value for i in range(1, 12)]
            matches.append((row, data))
    return matches

def process_refund(excel_file, sheet_name):
    """处理退款（仅更新J列，K列由公式自动计算）"""
    print("\n【处理退款】")
    print("🔍 请输入克重（必须输入，纯数字，如：17.68）")
    
    while True:
        weight_input = input("克重: ").strip()
        if not weight_input:
            print("❌ 克重不能为空！请重新输入")
            continue
        try:
            weight_val = float(weight_input)
            break
        except ValueError:
            print("❌ 克重必须是数字！请重新输入")
    
    matches = search_by_weight(weight_val, excel_file, sheet_name)
    
    if not matches:
        print(f"❌ 未找到克重 {weight_val} 的记录")
        return
    
    print(f"\n🔍 找到 {len(matches)} 条克重 {weight_val} 的记录，请选择：")
    for i, (row_idx, data) in enumerate(matches):
        profit_before = data[8] if data[8] is not None else "N/A"
        print(f"  {i+1}. 行{row_idx} | 平台:{data[5]} | 卖价:{data[7]} | 退款前利润:{profit_before}")
    
    try:
        choice = int(input("选择序号: ")) - 1
        if not (0 <= choice < len(matches)):
            print("❌ 无效序号")
            return
        row_num = matches[choice][0]
    except ValueError:
        print("❌ 请输入有效数字")
        return
    
    try:
        refund = float(input("\n退款金额 (纯数字): "))
    except ValueError:
        print("❌ 退款金额必须为数字")
        return

    # 仅更新J列（退款金额）
    wb = safe_load_workbook(excel_file)
    ws = wb[sheet_name]
    ws.cell(row=row_num, column=10, value=refund)  # J列
    wb.save(excel_file)
    
    print("✅ 退款金额已更新！")
    print(f"ℹ️ K{row_num}（退款后利润）将由公式自动计算")

def main():
    while True:
        print("\n" + "="*50)
        print("       卖货登记助手")
        print("="*50)
        print("1. 新增销售记录")
        print("2. 处理退款")
        print("3. 退出")
        choice = input("请选择操作: ").strip()
        
        if choice == "1":
            add_record(EXCEL_FILE, SHEET_NAME)
        elif choice == "2":
            process_refund(EXCEL_FILE, SHEET_NAME)
        elif choice == "3":
            print("👋 再见！")
            break
        else:
            print("❌ 无效选项，请重新选择")

if __name__ == "__main__":
    main()

