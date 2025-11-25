# -*- coding: utf-8 -*-
import os
import configparser
from datetime import datetime
from openpyxl import load_workbook, Workbook

# 配置文件路径
CONFIG_FILE = "config.ini"

def load_config():
    """加载或初始化配置文件"""
    config = configparser.ConfigParser()
    
    # 检查配置文件是否存在
    if not os.path.exists(CONFIG_FILE):
        # 创建默认配置
        config['DEFAULT'] = {
            'excel_file': '卖货登记.xlsx',
            'sheet_name': 'Sheet1'
        }
        with open(CONFIG_FILE, 'w') as f:
            config.write(f)
        print("ℹ️ 配置文件已创建，使用默认设置：")
        print(f"   Excel文件: {config['DEFAULT']['excel_file']}")
        print(f"   工作表: {config['DEFAULT']['sheet_name']}")
        return config
    
    # 读取现有配置
    config.read(CONFIG_FILE)
    return config

def get_config():
    """获取当前配置"""
    config = load_config()
    return config['DEFAULT']['excel_file'], config['DEFAULT']['sheet_name']

def init_excel(excel_file, sheet_name):
    """初始化Excel表格（自动创建表头）"""
    if not os.path.exists(excel_file):
        wb = Workbook()
        ws = wb.active
        ws.title = sheet_name
        headers = [
            "日期", "货名", "克重", "成本", "成本总价",
            "平台", "货源", "卖价", "退款前利润", "退款金额", "退款后利润"
        ]
        ws.append(headers)
        wb.save(excel_file)

def get_today():
    """获取当前日期（格式：2025年11月25日）"""
    return datetime.now().strftime("%Y年%m月%d日")

def calculate_profit(sell_price, cost):
    """计算退款前利润（卖价 - 成本）"""
    return sell_price - cost

def add_record(excel_file, sheet_name):
    """新增销售记录"""
    print("\n【新增销售记录】")
    try:
        goods = input("货名: ").strip()
        weight = float(input("克重 (纯数字): "))
        cost = float(input("成本 (纯数字): "))
        platform = input("平台: ").strip()
        source = input("货源: ").strip()
        sell_price = float(input("卖价 (纯数字): "))
    except ValueError:
        print("❌ 输入错误！请确保克重、成本、卖价为数字")
        return

    # 自动计算
    total_cost = weight * cost  # 成本总价
    profit_before = calculate_profit(sell_price, cost)  # 退款前利润

    # 保存到Excel
    wb = load_workbook(excel_file)
    ws = wb[sheet_name]
    ws.append([
        get_today(), goods, weight, cost, total_cost,
        platform, source, sell_price, profit_before,
        "", profit_before
    ])
    wb.save(excel_file)
    print(f"✅ 记录已添加！\n平台: {platform} | 成本总价: {total_cost} | 退款前利润: {profit_before}")

def search_records(criteria, excel_file, sheet_name):
    """根据条件查找记录"""
    wb = load_workbook(excel_file)
    ws = wb[sheet_name]
    matches = []
    for row_idx in range(2, ws.max_row + 1):
        data = [ws.cell(row=row_idx, column=col).value for col in range(1, 12)]
        if (data[1] == criteria["货名"] and 
            data[2] == criteria["克重"] and 
            data[3] == criteria["成本"] and 
            data[5] == criteria["平台"] and 
            data[6] == criteria["货源"] and 
            data[7] == criteria["卖价"]):
            matches.append((row_idx, data))
    return matches

def process_refund(excel_file, sheet_name):
    """处理退款"""
    print("\n【处理退款】")
    try:
        goods = input("货名: ").strip()
        weight = float(input("克重: "))
        cost = float(input("成本: "))
        platform = input("平台: ").strip()
        source = input("货源: ").strip()
        sell_price = float(input("卖价: "))
    except ValueError:
        print("❌ 数字格式错误！请确保输入为数字")
        return

    criteria = {
        "货名": goods, 
        "克重": weight, 
        "成本": cost, 
        "平台": platform, 
        "货源": source, 
        "卖价": sell_price
    }
    matches = search_records(criteria, excel_file, sheet_name)

    if not matches:
        print("❌ 未找到匹配记录（平台不匹配）")
        return

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

    try:
        refund = float(input("退款金额: "))
    except:
        print("❌ 退款金额必须为数字")
        return

    wb = load_workbook(excel_file)
    ws = wb[sheet_name]
    
    sell_val = ws.cell(row_num, 8).value
    cost_val = ws.cell(row_num, 4).value
    
    ws.cell(row_num, 10, refund)
    
    if refund >= sell_val:
        ws.cell(row_num, 11, 0)
        print("✅ 退款后利润已更新为 0（退款金额 ≥ 卖价）")
    else:
        ws.cell(row_num, 11, calculate_profit(sell_val, cost_val))
        print(f"✅ 退款后利润已更新为 {calculate_profit(sell_val, cost_val)}")

    wb.save(excel_file)
    print("✅ 退款记录更新成功！")

def main():
    """主程序入口"""
    # 获取配置
    excel_file, sheet_name = get_config()
    
    # 初始化Excel（如果不存在）
    init_excel(excel_file, sheet_name)
    
    while True:
        print("\n📦「卖货登记助手」")
        print("1️⃣ 新增销售记录  2️⃣ 处理退款  3️⃣ 配置文件  4️⃣ 退出")
        choice = input("请选择: ").strip()
        
        if choice == "1":
            add_record(excel_file, sheet_name)
        elif choice == "2":
            process_refund(excel_file, sheet_name)
        elif choice == "3":
            # 配置菜单
            print("\n🔧 配置管理")
            print("1. 修改Excel文件名")
            print("2. 修改工作表名称")
            print("3. 返回")
            config_choice = input("请选择: ").strip()
            
            if config_choice == "1":
                new_file = input("请输入新的Excel文件名（含扩展名）: ").strip()
                if new_file:
                    config = configparser.ConfigParser()
                    config.read(CONFIG_FILE)
                    config['DEFAULT']['excel_file'] = new_file
                    with open(CONFIG_FILE, 'w') as f:
                        config.write(f)
                    print(f"✅ Excel文件已更新为: {new_file}")
                    # 重新加载配置
                    excel_file, sheet_name = get_config()
            elif config_choice == "2":
                new_sheet = input("请输入新的工作表名称: ").strip()
                if new_sheet:
                    config = configparser.ConfigParser()
                    config.read(CONFIG_FILE)
                    config['DEFAULT']['sheet_name'] = new_sheet
                    with open(CONFIG_FILE, 'w') as f:
                        config.write(f)
                    print(f"✅ 工作表已更新为: {new_sheet}")
                    # 重新加载配置
                    excel_file, sheet_name = get_config()
            elif config_choice == "3":
                continue
            else:
                print("⚠️ 无效选项")
        elif choice == "4":
            print("👋 谢谢使用，再见！")
            break
        else:
            print("⚠️ 请输入 1/2/3/4")

if __name__ == "__main__":
    main()
