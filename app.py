def process_refund(excel_file, sheet_name):
    """优化版退款流程：智能匹配+用户选择+自动计算"""
    print("\n【处理退款】")
    print("🔍 请输入任意匹配条件（留空跳过），系统将自动查找匹配记录")
    
    # 1. 获取用户输入（支持部分匹配）
    criteria = {
        "货名": input("货名 (可留空): ").strip(),
        "平台": input("平台 (可留空): ").strip(),
        "卖价": input("卖价 (可留空): ").strip(),
        "货源": input("货源 (可留空): ").strip()
    }
    
    # 2. 查找所有匹配记录
    matches = search_records(criteria, excel_file, sheet_name)
    
    if not matches:
        print("❌ 未找到匹配记录，请检查输入条件")
        return
    
    # 3. 显示匹配记录供用户选择
    print(f"\n🔍 找到 {len(matches)} 条匹配记录，请选择：")
    for i, (row_idx, data) in enumerate(matches):
        # 格式化显示关键信息
        print(f"  {i+1}. 行{row_idx} | 货名:{data[1]} | 平台:{data[5]} | 卖价:{data[7]} | 退款前利润:{data[8]}")
    
    # 4. 用户选择记录
    try:
        choice = int(input("选择序号: ")) - 1
        if 0 <= choice < len(matches):
            row_num = matches[choice][0]
        else:
            print("❌ 无效序号")
            return
    except:
        print("❌ 请输入有效数字")
        return
    
    # 5. 用户输入退款金额
    try:
        refund = float(input("\n退款金额 (纯数字): "))
    except:
        print("❌ 退款金额必须为数字")
        return
    
    # 6. 自动计算并更新
    wb = load_workbook(excel_file)
    ws = wb[sheet_name]
    
    # 获取当前卖价和成本
    sell_val = ws.cell(row_num, 8).value  # 第8列: 卖价
    cost_val = ws.cell(row_num, 4).value  # 第4列: 成本单价
    
    # 更新退款金额 (第10列)
    ws.cell(row_num, 10, refund)
    
    # 自动计算退款后利润 (第11列)
    if refund >= sell_val:
        new_profit = 0
        print("✅ 退款后利润已更新为 0（退款金额 ≥ 卖价）")
    else:
        new_profit = calculate_profit(sell_val, cost_val)  # 保持原退款前利润
        print(f"✅ 退款后利润已更新为 {new_profit}（退款金额 < 卖价）")
    
    # 更新退款后利润
    ws.cell(row_num, 11, new_profit)
    
    wb.save(excel_file)
    print("✅ 退款记录更新成功！")
