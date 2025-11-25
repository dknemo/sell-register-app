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
    """新增销售记录（K列使用智能公式 + 横向显示）"""
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

    # Python 实时计算用于回显（不影响Excel公式）
    total_cost = weight * cost
    profit_before = max(0, sell_price - total_cost)  # 回显也按>=0处理

    wb = safe_load_workbook(excel_file)
    ws = wb[sheet_name]
    
    max_row = ws.max_row
    if max_row < 2:
        new_row = 2
    else:
        new_row = max_row - 1
    
    print(f"ℹ️ 新记录将添加在第{new_row}行（倒数第二行）")
    
    # ====== 关键：K列使用您指定的智能公式 ======
    data = [
        get_today(), goods, weight, cost,
        f"=C{new_row}*D{new_row}",          # E: 成本总价
        platform, source, sell_price,
        f"=H{new_row}-E{new_row}",          # I: 退款前利润（可为负）
        "",                                  # J: 退款金额（初始空）
        f"=IF(J{new_row}=\"\", MAX(0, H{new_row}-E{new_row}), MAX(0, H{new_row}-E{new_row}-J{new_row}))"  # K: 智能公式
    ]
    
    for col_idx, value in enumerate(data, start=1):
        ws.cell(row=new_row, column=col_idx, value=value)
    
    wb.save(excel_file)
    
    # ====== 横向回显（使用Python计算值，模拟Excel效果） ======
    headers = ["日期", "货名", "克重", "成本单价", "成本总价",
               "平台", "货源", "卖价", "退款前利润",
               "退款金额", "退款后利润"]
    
    # 回显时：退款后利润 = MAX(0, 卖价 - 成本总价) （因J为空）
    refund_after_display = max(0, sell_price - total_cost)
    
    display_values = [
        get_today(), goods,
        f"{weight:.2f}",
        f"{cost:.2f}",
        f"{total_cost:.2f}",
        platform, source,
        f"{sell_price:.2f}",
        f"{sell_price - total_cost:.2f}",   # I列可能为负（如实显示）
        "",
        f"{refund_after_display:.2f}"       # K列 >=0
    ]
    
    print("\n✅ 记录已成功添加！完整数据如下：")
    print("=" * 120)
    print("".join([f"{h:>10}" for h in headers]))
    print("".join([f"{str(v):>10}" for v in display_values]))
    print("=" * 120)
    
    print("\nℹ️ 公式说明：")
    print(f"  • K{new_row} = IF(J{new_row}=\"\", MAX(0,H{new_row}-E{new_row}), MAX(0,H{new_row}-E{new_row}-J{new_row}))")
    print("  • 当您填写退款金额（J列）后，K列将自动更新！")
    
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
    """处理退款（仅更新退款金额J列，K列由公式自动计算）"""
    print("\n【处理退款】")
    print("🔍 请输入克重（必须输入，纯数字，如：17.68）")
    
    # 安全输入克重
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

    # ====== 关键修复：只更新 J 列（退款金额），不碰 K 列！ ======
    wb = safe_load_workbook(excel_file)
    ws = wb[sheet_name]
    
    # 写入退款金额到 J 列（第10列）
    ws.cell(row=row_num, column=10, value=refund)
    
    # ⚠️ 不再操作 K 列（第11列）！公式会自动更新
    
    wb.save(excel_file)
    
    print("✅ 退款金额已更新！")
    print(f"ℹ️ K{row_num}（退款后利润）将由公式自动计算：")
    print(f"   =IF(J{row_num}=\"\", MAX(0,H{row_num}-E{row_num}), MAX(0,H{row_num}-E{row_num}-J{row_num}))")
    
def search_by_weight(weight, excel_file, sheet_name):
    """仅按克重匹配记录（支持浮点数）"""
    wb = safe_load_workbook(excel_file)
    ws = wb[sheet_name]
    matches = []
    
    for row_idx in range(2, ws.max_row + 1):
        weight_cell = ws.cell(row=row_idx, column=3).value
        if weight_cell is None:
            continue
            
        try:
            weight_val = float(weight_cell)
        except:
            continue
            
        if abs(weight_val - weight) < 1e-5:
            data = [ws.cell(row=row_idx, column=col).value for col in range(1, 12)]
            matches.append((row_idx, data))
    
    return matches
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









