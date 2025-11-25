# -*- coding: utf-8 -*-
import os
import configparser
import sys
from datetime import datetime
from openpyxl import load_workbook, Workbook
from openpyxl.utils import get_column_letter

# 配置文件路径
CONFIG_FILE = "config.ini"

def safe_load_workbook(file_path):
    """安全加载Excel文件（处理被占用的情况）"""
    try:
        return load_workbook(file_path)
    except PermissionError:
        print(f"❌ 无法打开Excel文件: {file_path}")
        print("⚠️ 请关闭所有打开的Excel文件（包括Excel的后台进程）")
        print("👉 解决方法：在任务管理器中结束Excel进程")
        sys.exit(1)
    except Exception as e:
        print(f"❌ Excel加载错误: {str(e)}")
        print("👉 请检查文件路径或Excel文件是否损坏")
        sys.exit(1)

def load_config():
    """加载或初始化配置文件（带安全检查）"""
    config = configparser.ConfigParser()
    
    # 检查配置文件是否存在
    if not os.path.exists(CONFIG_FILE):
        # 创建默认配置
        config['DEFAULT'] = {
            'excel_file': '卖货登记.xlsx',
            'sheet_name': 'Sheet1'
        }
        try:
            with open(CONFIG_FILE, 'w') as f:
                config.write(f)
            print("ℹ️ 配置文件已创建，使用默认设置：")
            print(f"   Excel文件: {config['DEFAULT']['excel_file']}")
            print(f"   工作表: {config['DEFAULT']['sheet_name']}")
        except PermissionError:
            print(f"❌ 无法创建配置文件: {CONFIG_FILE}")
            print("👉 请确保程序有权限写入当前目录")
            sys.exit(1)
        return config
    
    # 读取现有配置
    try:
        config.read(CONFIG_FILE)
        return config
    except Exception as e:
        print(f"❌ 配置文件读取错误: {str(e)}")
        print("👉 请检查配置文件权限或内容")
        sys.exit(1)

def get_config():
    """获取当前配置（安全处理）"""
    config = load_config()
    return config['DEFAULT']['excel_file'], config['DEFAULT']['sheet_name']

def init_excel(excel_file, sheet_name):
    """初始化Excel表格（安全创建）"""
    # 检查文件是否被占用
    if os.path.exists(excel_file):
        try:
            wb = safe_load_workbook(excel_file)
            wb.close()
        except:
            pass  # 如果被占用，尝试关闭后再创建

    if not os.path.exists(excel_file):
        try:
            wb = Workbook()
            ws = wb.active
            ws.title = sheet_name
            headers = [
                "日期", "货名", "克重", "成本单价", "成本总价",
                "平台", "货源", "卖价", "退款前利润", "退款金额", "退款后利润"
            ]
            ws.append(headers)
            wb.save(excel_file)
            print(f"✅ Excel文件已创建: {excel_file}")
        except Exception as e:
            print(f"❌ 创建Excel文件失败: {str(e)}")
            print("👉 请检查文件路径或权限")
            sys.exit(1)

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
        cost = float(input("成本单价 (纯数字): "))
        platform = input("平台: ").strip()
        source = input("货源: ").strip()
        sell_price = float(input("卖价 (纯数字): "))
    except ValueError:
        print("❌ 输入错误！请确保克重、成本单价、卖价为数字")
        return

    total_cost = weight * cost
    profit_before = calculate_profit(sell_price, cost)

    wb = safe_load_workbook(excel_file)
    ws = wb[sheet_name]
    
    ws.append([
        get_today(), goods, weight, cost, total_cost,
        platform, source, sell_price, profit_before,
        "", profit_before
    ])
    wb.save(excel_file)
    print(f"✅ 记录已添加！\n平台: {platform} | 成本总价: {total_cost} | 退款前利润: {profit_before}")

def search_records(criteria, excel_file, sheet_name):
    """智能匹配：支持任意字段匹配（安全处理）"""
    wb = safe_load_workbook(excel_file)
    ws = wb[sheet_name]
    matches = []
    
    for row_idx in range(2, ws.max_row + 1):
        data = [ws.cell(row=row_idx, column=col).value for col in range(1, 12)]
        matches_all = True
        
        for key, value in criteria.items():
            if value:  # 只检查非空条件
                col_idx = {
                    "货名": 2,
                    "平台": 6,
                    "卖价": 8,
                    "货源": 7
                }[key]
                
                # 安全处理空值
                cell_value = data[col_idx-1] if data[col_idx-1] is not None else ""
                if str(cell_value) != str(value):
                    matches_all = False
                    break
        
        if matches_all:
            matches.append((row_idx, data))
    
    return matches

def process_refund(excel_file, sheet_name):
    """优化版退款流程：智能匹配+用户选择+自动计算（安全处理）"""
    print("\n【处理退款】")
    print("🔍 请输入任意匹配条件（留空跳过），系统将自动查找匹配记录")
    
    criteria = {
        "货名": input("货名 (可留空): ").strip(),
        "平台": input("平台 (可留空): ").strip(),
        "卖价": input("卖价 (可留空): ").strip(),
        "货源": input("货源 (可留空): ").strip()
    }
    
    matches = search_records(criteria, excel_file, sheet_name)
    
    if not matches:
        print("❌ 未找到匹配记录，请检查输入条件")
        return
    
    print(f"\n🔍 找到 {len(matches)} 条匹配记录，请选择：")
    for i, (row_idx, data) in enumerate(matches):
        # 安全处理空值
        profit_before = data[8] if data[8] is not None else 0
        print(f"  {i+1}. 行{row_idx} | 货名:{data[1] or 'N/A'} | 平台:{data[5] or 'N/A'} | 卖价:{data[7] or 'N/A'} | 退款前利润:{profit_before}")
    
    try:
        choice = int(input("选择序号: ")) - 1
        if 0 <= choice < len(matches):
            row_num = matches[choice][0]
        else:
            print("❌ 无效序号")
            return
    except ValueError:
        print("❌ 请输入有效数字")
        return
    
    try:
        refund = float(input("\n退款金额 (纯数字): "))
    except ValueError:
        print("❌ 退款金额必须为数字")
        return
    
    wb = safe_load_workbook(excel_file)
    ws = wb[sheet_name]
    
    # 安全获取卖价和成本
    sell_val = ws.cell(row=row_num, column=8).value
    cost_val = ws.cell(row=row_num, column=4).value
    
    if sell_val is None or cost_val is None:
        print("❌ 记录数据不完整（卖价/成本缺失）")
        return
    
    # 更新退款金额 (第10列)
    ws.cell(row=row_num, column=10, value=refund)
    
    # 自动计算退款后利润 (第11列)
    if refund >= sell_val:
        new_profit = 0
        print("✅ 退款后利润已更新为 0（退款金额 ≥ 卖价）")
    else:
        new_profit = calculate_profit(sell_val, cost_val)
        print(f"✅ 退款后利润已更新为 {new_profit}（退款金额 < 卖价）")
    
    # 更新退款后利润
    ws.cell(row=row_num, column=11, value=new_profit)
    
    wb.save(excel_file)
    print("✅ 退款记录更新成功！")

def main():
    """主程序入口（安全启动）"""
    try:
        excel_file, sheet_name = get_config()
        init_excel(excel_file, sheet_name)
    except Exception as e:
        print(f"❌ 初始化失败: {str(e)}")
        print("👉 请检查配置文件或Excel文件权限")
        sys.exit(1)
    
    while True:
        print("\n📦「卖货登记助手」")
        print("1️⃣ 新增销售记录  2️⃣ 处理退款  3️⃣ 配置文件  4️⃣ 退出")
        choice = input("请选择: ").strip()
        
        if choice == "1":
            add_record(excel_file, sheet_name)
        elif choice == "2":
            process_refund(excel_file, sheet_name)
        elif choice == "3":
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
                    try:
                        with open(CONFIG_FILE, 'w') as f:
                            config.write(f)
                        print(f"✅ Excel文件已更新为: {new_file}")
                        excel_file, sheet_name = get_config()
                    except Exception as e:
                        print(f"❌ 配置保存失败: {str(e)}")
            elif config_choice == "2":
                new_sheet = input("请输入新的工作表名称: ").strip()
                if new_sheet:
                    config = configparser.ConfigParser()
                    config.read(CONFIG_FILE)
                    config['DEFAULT']['sheet_name'] = new_sheet
                    try:
                        with open(CONFIG_FILE, 'w') as f:
                            config.write(f)
                        print(f"✅ 工作表已更新为: {new_sheet}")
                        excel_file, sheet_name = get_config()
                    except Exception as e:
                        print(f"❌ 配置保存失败: {str(e)}")
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
    try:
        main()
    except Exception as e:
        print(f"❌ 程序运行时发生严重错误: {str(e)}")
        print("👉 请截图此错误信息并联系开发者")
        input("按回车键退出...")
