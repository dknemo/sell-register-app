# -*- coding: utf-8 -*-
import os
from datetime import datetime
from openpyxl import load_workbook, Workbook

EXCEL_FILE = "卖货登记.xlsx"
SHEET_NAME = "Sheet1"

def init_excel():
    """初始化Excel表格（自动创建表头）"""
    if not os.path.exists(EXCEL_FILE):
        wb = Workbook()
        ws = wb.active
        ws.title = SHEET_NAME
        headers = [
            "日期", "货名", "克重", "成本", "成本总价",
            "平台", "货源", "卖价", "退款前利润", "退款金额", "退款后利润"
        ]
        ws.append(headers)
        wb.save(EXCEL_FILE)

def get_today():
    """获取当前日期（格式：2025年11月25日）"""
    return datetime.now().strftime("%Y年%m月%d日")

def calculate_profit(sell_price, cost):
    """计算退款前利润（卖价 - 成本）"""
    return sell_price - cost

def add_record():
    """新增销售记录（平台可自定义）"""
    print("\n【新增销售记录】")
    try:
        goods = input("货名: ").strip()
        weight = float(input("克重 (纯数字): "))
        cost = float(input("成本 (纯数字): "))
        platform = input("平台: ").strip()  # ✅ 新增：平台可自定义
        source = input("货源: ").strip()
        sell_price = float(input("卖价 (纯数字): "))
    except ValueError:
        print("❌ 输入错误！请确保克重、成本、卖价为数字")
        return

    # 自动计算
    total_cost = weight * cost  # 成本总价
    profit_before = calculate_profit(sell_price, cost)  # 退款前利润

    # 保存到Excel
    wb = load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME]
    ws.append([
        get_today(), goods, weight, cost, total_cost,
        platform, source, sell_price, profit_before,
        "", profit_before  # 退款后利润（默认等于退款前利润）
    ])
    wb.save(EXCEL_FILE)
    print(f"✅ 记录已添加！\n平台: {platform} | 成本总价: {total_cost} | 退款前利润: {profit_before}")

def search_records(criteria):
    """根据条件查找记录（包含平台匹配）"""
    wb = load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME]
    matches = []
    for row_idx in range(2, ws.max_row + 1):
        data = [ws.cell(row=row_idx, column=col).value for col in range(1, 12)]
        # 检查所有关键字段是否匹配（新增平台匹配）
        if (data[1] == criteria["货名"] and 
            data[2] == criteria["克重"] and 
            data[3] == criteria["成本"] and 
            data[5] == criteria["平台"] and  # ✅ 新增：平台匹配
            data[6] == criteria["货源"] and 
            data[7] == criteria["卖价"]):
            matches.append((row_idx, data))
    return matches

def process_refund():
    """处理退款（平台需匹配）"""
    print("\n【处理退款】")
    try:
        goods = input("货名: ").strip()
        weight = float(input("克重: "))
        cost = float(input("成本: "))
        platform = input("平台: ").strip()  # ✅ 新增：退款时需输入平台
        source = input("货源: ").strip()
        sell_price = float(input("卖价: "))
    except ValueError:
        print("❌ 数字格式错误！请确保输入为数字")
        return

    criteria = {
        "货名": goods, 
        "克重": weight, 
        "成本": cost, 
        "平台": platform,  # ✅ 新增：平台字段
        "货源": source, 
        "卖价": sell_price
    }
    matches = search_records(criteria)

    if not matches:
        print("❌ 未找到匹配记录（平台不匹配）")
        return

    # 多条记录处理
    if len(matches) > 1:
        print(f"🔍 找到 {len(matches)} 条匹配记录，请选择：")
        for i, (r, d) in enumerate(matches):
            print(f"  {i+1}. 行{r} | {d[1]} | 克重:{d[2]} | 成本:{d[3]} | 平台:{d[5]} | 卖价:{d[7]}")
        try:
            choice = int(input("选择序号: ")) - 1
            if 0 <= choice < len(matches):
                row_num = matches[choice][0]
            else:
                print("❌ 无效序号")
                return
        except:
            print("❌ 请输入数字")
            return
    else:
        row_num = matches[0][0]

    # 输入退款金额
    try:
        refund = float(input("退款金额: "))
    except:
        print("❌ 退款金额必须为数字")
        return

    # 更新Excel
    wb = load_workbook(EXCEL_FILE)
    ws = wb[SHEET_NAME]
    
    # 获取当前卖价和成本
    sell_val = ws.cell(row_num, 8).value
    cost_val = ws.cell(row_num, 4).value
    
    # 更新退款金额和退款后利润
    ws.cell(row_num, 10, refund)  # 第10列：退款金额
    
    # 退款后利润逻辑：退款≥卖价→0，否则=退款前利润
    if refund >= sell_val:
        ws.cell(row_num, 11, 0)
        print("✅ 退款后利润已更新为 0（退款金额 ≥ 卖价）")
    else:
        ws.cell(row_num, 11, calculate_profit(sell_val, cost_val))
        print(f"✅ 退款后利润已更新为 {calculate_profit(sell_val, cost_val)}")

    wb.save(EXCEL_FILE)
    print("✅ 退款记录更新成功！")

def main():
    """主程序入口"""
    init_excel()
    while True:
        print("\n📦「卖货登记助手」")
        print("1️⃣ 新增销售记录  2️⃣ 处理退款  3️⃣ 退出")
        choice = input("请选择: ").strip()
        
        if choice == "1":
            add_record()
        elif choice == "2":
            process_refund()
        elif choice == "3":
            print("👋 谢谢使用，再见！")
            break
        else:
            print("⚠️ 请输入 1/2/3")

if __name__ == "__main__":
    main()
