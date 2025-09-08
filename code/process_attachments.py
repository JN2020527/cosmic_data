import os
import re
from datetime import datetime
from typing import Optional, Tuple, List
import json

from openpyxl import load_workbook
import requests


DATA_DIR = "/Users/hua/Desktop/my-code/cosmic_data/data_file"
DEEPSEEK_API_KEY = "sk-e59e19d1305c4fbf8e4833e0f26b2ceb"
DEEPSEEK_API_URL = "https://api.deepseek.com/v1/chat/completions"


def print_step(title: str) -> None:
    print(f"==== {title} ====")


def parse_attachment_filename(filename: str) -> Optional[Tuple[str, str, str, str]]:
    """
    Parse filenames like: 附件x-需求名@属性.扩展名
    Returns tuple: (prefix_with_number, requirement_name, attribute, extension)
    Example: ("附件3", "关于xxx的需求", "COSMIC工作量评估基础表", ".xlsx")
    """
    name, ext = os.path.splitext(filename)
    # Work with just the basename (no directories)
    base = os.path.basename(name)
    match = re.match(r"^附件(\d+)-(.+?)@(.+)$", base)
    if not match:
        return None
    number, req_name, attribute = match.groups()
    return (f"附件{number}", req_name, attribute, ext)


def build_attachment_filename(prefix_with_number: str, requirement_name: str, attribute: str, ext: str) -> str:
    return f"{prefix_with_number}-{requirement_name}@{attribute}{ext}"


def batch_rename(requirement_name: str) -> None:
    print_step("第一步：批量修改文件名字")
    if not os.path.isdir(DATA_DIR):
        print(f"目录不存在：{DATA_DIR}")
        return

    files = os.listdir(DATA_DIR)
    rename_count = 0
    for fname in files:
        parsed = parse_attachment_filename(fname)
        if not parsed:
            continue
        prefix, _old_req, attribute, ext = parsed
        new_name = build_attachment_filename(prefix, requirement_name, attribute, ext)
        if new_name == fname:
            print(f"跳过（已是目标名）：{fname}")
            continue
        src = os.path.join(DATA_DIR, fname)
        dst = os.path.join(DATA_DIR, new_name)
        os.rename(src, dst)
        rename_count += 1
        print(f"重命名：{fname} -> {new_name}")

    print(f"重命名完成，共处理 {rename_count} 个文件。")


def find_attachment_by_number(number: int) -> Optional[str]:
    """Find file path for '附件{number}-...@....xlsx' in DATA_DIR."""
    if not os.path.isdir(DATA_DIR):
        return None
    for fname in os.listdir(DATA_DIR):
        parsed = parse_attachment_filename(fname)
        if not parsed:
            continue
        prefix, _req, _attr, _ext = parsed
        if prefix == f"附件{number}":
            return os.path.join(DATA_DIR, fname)
    return None


def write_attachment3_sheet2_cells(requirement_name: str) -> None:
    print_step("第二步：修改附件3 sheet2 的 A3、B3")
    path = find_attachment_by_number(3)
    if not path:
        print("未找到附件3 文件。")
        return
    wb = load_workbook(path)
    sheet_name = None
    # Prefer explicit 'sheet2' by index (second worksheet)
    if len(wb.sheetnames) >= 2:
        sheet_name = wb.sheetnames[1]
    else:
        sheet_name = wb.active.title
    ws = wb[sheet_name]
    ws["A3"] = requirement_name
    ws["B3"] = requirement_name
    wb.save(path)
    print(f"已更新 {os.path.basename(path)} -> {sheet_name} 的 A3, B3 为：{requirement_name}")


def write_attachment4_cells(requirement_name: str) -> None:
    print_step("第三步：修改附件4 B2")
    path = find_attachment_by_number(4)
    if not path:
        print("未找到附件4 文件。")
        return
    wb = load_workbook(path)
    ws = wb.active
    ws["B2"] = requirement_name
    wb.save(path)
    print(f"已更新 {os.path.basename(path)} -> B2 为：{requirement_name}")


def sum_attachment5_col_L_from_L2() -> float:
    print_step("第四步：计算附件5 L列(L2开始)总和")
    path = find_attachment_by_number(5)
    if not path:
        print("未找到附件5 文件。返回 0。")
        return 0.0

    ext = os.path.splitext(path)[1].lower()

    # Handle legacy .xls via xlrd
    if ext == ".xls":
        try:
            import xlrd  # type: ignore
        except Exception:
            print("未安装 xlrd，无法读取 .xls 文件，请先安装 xlrd。返回 0。")
            return 0.0
        try:
            wb = xlrd.open_workbook(path)
            sheet = wb.sheet_by_index(0)
            total: float = 0.0
            col_index_for_L = 11  # 0-based index for column 'L'
            for row_idx in range(1, sheet.nrows):  # start from row 2 -> index 1
                value = sheet.cell_value(row_idx, col_index_for_L)
                if value in (None, ""):
                    continue
                try:
                    total += float(value)
                except Exception:
                    # skip non-numeric cells
                    continue
            print(f"L列合计：{total}")
            return total
        except Exception as e:
            print(f"读取 .xls 失败：{e}")
            return 0.0

    # Default: .xlsx via openpyxl
    try:
        wb = load_workbook(path, data_only=True)
        ws = wb.active
        total = 0.0
        max_row = ws.max_row if ws.max_row and ws.max_row > 1 else 1
        for row in range(2, max_row + 1):
            cell = ws[f"L{row}"]
            value = cell.value
            if value is None or (isinstance(value, str) and value.strip() == ""):
                continue
            try:
                total += float(value)
            except Exception:
                continue
        print(f"L列合计：{total}")
        return total
    except Exception as e:
        print(f"读取 .xlsx 失败：{e}")
        return 0.0


def write_attachment4_with_sum(total: float) -> None:
    print_step("第五步：将总和写入附件4 D7")
    path = find_attachment_by_number(4)
    if not path:
        print("未找到附件4 文件。")
        return
    wb = load_workbook(path)
    ws = wb.active
    ws["D7"] = total
    wb.save(path)
    print(f"已更新 {os.path.basename(path)} -> D7 为：{total}")


def write_attachment3_sheet2_F3_with_sum(total: float) -> None:
    print_step("第六步：将总和写入附件3 sheet2 的 F3")
    path = find_attachment_by_number(3)
    if not path:
        print("未找到附件3 文件。")
        return
    wb = load_workbook(path)
    # second sheet or active
    if len(wb.sheetnames) >= 2:
        ws = wb[wb.sheetnames[1]]
    else:
        ws = wb.active
    ws["F3"] = total
    wb.save(path)
    print(f"已更新 {os.path.basename(path)} -> {ws.title} 的 F3 为：{total}")


def write_attachment4_C6_with_today() -> None:
    print_step("第七步：填充当前时间到附件4 C6（xxxx年x月x日）")
    path = find_attachment_by_number(4)
    if not path:
        print("未找到附件4 文件。")
        return
    today = datetime.now()
    # Remove leading zeros in month/day
    date_str = f"{today.year}年{today.month}月{today.day}日"
    wb = load_workbook(path)
    ws = wb.active
    ws["C6"] = date_str
    wb.save(path)
    print(f"已更新 {os.path.basename(path)} -> C6 为：{date_str}")


def calculate_attachment3_e3_formula() -> int:
    """手动计算附件3 sheet2 E3单元格的公式: =COUNTA(COSMIC功能点拆分表!K:K)-1
    注意：实际数据从K4开始，K1-K3是标题行"""
    path3 = find_attachment_by_number(3)
    if not path3:
        return 0
    
    wb3 = load_workbook(path3, data_only=True)
    
    # 找到COSMIC功能点拆分表工作表
    cosmic_sheet = None
    for sheet_name in wb3.sheetnames:
        if 'COSMIC功能点拆分表' in sheet_name:
            cosmic_sheet = wb3[sheet_name]
            break
    
    if not cosmic_sheet:
        return 0
    
    # 从K4开始计算非空单元格数量（K1-K3是标题行）
    k_count = 0
    for row in range(4, cosmic_sheet.max_row + 1):
        k_value = cosmic_sheet[f'K{row}'].value
        if k_value is not None and str(k_value).strip():
            k_count += 1
    
    print(f"COSMIC功能点拆分表K列数据行数（从K4开始）: {k_count}")
    # 返回实际的功能点数量
    return k_count


def write_attachment4_B7_from_attachment3_E3() -> None:
    print_step("第八步：将附件3 sheet2 E3 写入附件4 B7")
    path3 = find_attachment_by_number(3)
    path4 = find_attachment_by_number(4)
    if not path3:
        print("未找到附件3 文件。")
        return
    if not path4:
        print("未找到附件4 文件。")
        return
    
    # 首先尝试直接读取E3的值
    wb3 = load_workbook(path3, data_only=True)
    if len(wb3.sheetnames) >= 2:
        ws3 = wb3[wb3.sheetnames[1]]
    else:
        ws3 = wb3.active
    e3_value = ws3["E3"].value

    # 如果E3的值是None，说明它包含公式且无法直接计算，手动计算
    if e3_value is None:
        print("E3单元格包含公式，手动计算结果...")
        e3_value = calculate_attachment3_e3_formula()
        print(f"计算得到E3公式结果: {e3_value}")

    wb4 = load_workbook(path4)
    ws4 = wb4.active
    ws4["B7"] = e3_value
    wb4.save(path4)
    print(f"已更新 {os.path.basename(path4)} -> B7 为：{e3_value}")


def extract_attachment5_h_i_content() -> Tuple[List[str], List[str]]:
    """提取附件5中H列和I列的内容"""
    path = find_attachment_by_number(5)
    if not path:
        return [], []
    
    try:
        import xlrd
    except Exception:
        print("未安装 xlrd，无法读取 .xls 文件")
        return [], []
    
    try:
        wb = xlrd.open_workbook(path)
        sheet = wb.sheet_by_index(0)
        
        h_contents = []
        i_contents = []
        h_col_index = 7  # H列索引
        i_col_index = 8  # I列索引
        
        # 从第2行开始读取（跳过标题行）
        for row_idx in range(1, sheet.nrows):
            h_value = sheet.cell_value(row_idx, h_col_index) if h_col_index < sheet.ncols else ''
            i_value = sheet.cell_value(row_idx, i_col_index) if i_col_index < sheet.ncols else ''
            
            if h_value and str(h_value).strip():
                h_contents.append(str(h_value).strip())
            if i_value and str(i_value).strip():
                i_contents.append(str(i_value).strip())
        
        return h_contents, i_contents
    except Exception as e:
        print(f"读取附件5失败：{e}")
        return [], []


def call_deepseek_api(content: str) -> str:
    """调用DeepSeek API生成内容概述"""
    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}"
    }
    
    prompt = f"""基于以下工作项内容，请生成一个精简的需求内容概述。要求：
1. 以"内容概述："开头
2. 总结的内容为有序列表
3. 每个列表项应该简洁明了，概括主要功能点
4. 合并相似的功能点
5. 按照逻辑顺序排列

工作项内容：
{content}

请生成概述："""

    data = {
        "model": "deepseek-chat",
        "messages": [
            {
                "role": "user",
                "content": prompt
            }
        ],
        "max_tokens": 1000,
        "temperature": 0.7
    }
    
    try:
        print(f"正在调用DeepSeek API生成概述...")
        
        response = requests.post(DEEPSEEK_API_URL, headers=headers, json=data, timeout=30)
        
        if response.status_code != 200:
            error_msg = f"API调用失败，状态码: {response.status_code}"
            print(error_msg)
            raise Exception(error_msg)
        
        result = response.json()
        
        if 'choices' in result and len(result['choices']) > 0:
            api_summary = result['choices'][0]['message']['content'].strip()
            print("✅ API调用成功，已生成概述")
            return api_summary
        else:
            error_msg = "API响应格式错误"
            print(error_msg)
            raise Exception(error_msg)
            
    except Exception as e:
        error_msg = f"调用DeepSeek API失败：{e}"
        print(error_msg)
        raise Exception(error_msg)


def summarize_requirement_content_and_update_h4() -> None:
    print_step("第九步：基于附件5 H和I列内容生成需求概述并更新附件4 A4")
    
    try:
        # 提取附件5的H和I列内容
        h_contents, i_contents = extract_attachment5_h_i_content()
        
        if not h_contents and not i_contents:
            print("未找到附件5的H和I列内容")
            return
        
        print(f"提取到H列内容 {len(h_contents)} 项，I列内容 {len(i_contents)} 项")
        
        # 合并H和I列内容（去重）
        all_contents = []
        all_contents.extend(h_contents)
        for content in i_contents:
            if content not in all_contents:
                all_contents.append(content)
        
        # 将内容合并为一个字符串
        content_text = "\n".join([f"{i+1}. {content}" for i, content in enumerate(all_contents)])
        
        print("正在调用DeepSeek API生成需求概述...")
        summary = call_deepseek_api(content_text)
        
        print(f"生成的概述：\n{summary}")
        
        # 更新附件4的A4单元格
        path4 = find_attachment_by_number(4)
        if not path4:
            print("未找到附件4文件")
            return
        
        wb4 = load_workbook(path4)
        ws4 = wb4.active
        ws4["A4"] = summary
        wb4.save(path4)
        
        print(f"已更新 {os.path.basename(path4)} -> A4 为概述内容")
        
    except Exception as e:
        print(f"第九步执行失败：{e}")
        print("程序终止，请检查API配置或网络连接")
        raise


def load_function_codes() -> List[Tuple[str, str, str]]:
    """加载一二三级功能点码值文件"""
    codes_path = os.path.join(os.path.dirname(__file__), "一二三级功能点.xlsx")
    if not os.path.exists(codes_path):
        print(f"未找到功能点码值文件：{codes_path}")
        return []
    
    try:
        wb = load_workbook(codes_path, data_only=True)
        ws = wb.active
        
        codes = []
        current_level1 = ""
        current_level2 = ""
        
        for row in range(2, ws.max_row + 1):
            level1 = ws.cell(row, 1).value
            level2 = ws.cell(row, 2).value
            level3 = ws.cell(row, 3).value
            
            # 更新当前的一级、二级功能点
            if level1 and str(level1).strip():
                current_level1 = str(level1).strip()
            if level2 and str(level2).strip():
                current_level2 = str(level2).strip()
            
            if level3 and str(level3).strip():
                codes.append((current_level1, current_level2, str(level3).strip()))
        
        print(f"加载了 {len(codes)} 个功能点码值")
        return codes
    except Exception as e:
        print(f"加载功能点码值失败：{e}")
        return []


def parse_ai_function_matches(ai_response: str, function_codes: List[Tuple[str, str, str]]) -> List[Tuple[str, str, str, str, float]]:
    """解析AI返回的功能点匹配结果"""
    matches = []
    lines = ai_response.split('\n')
    
    for line in lines:
        line = line.strip()
        if '|' in line and not line.startswith('功能点编号'):
            try:
                parts = line.split('|')
                if len(parts) >= 6:
                    number = parts[0].strip()
                    level1 = parts[1].strip()
                    level2 = parts[2].strip()
                    level3 = parts[3].strip()
                    description = parts[4].strip()
                    workload = float(parts[5].strip())
                    
                    # 验证功能点是否在码值中存在
                    if (level1, level2, level3) in function_codes:
                        matches.append((level1, level2, level3, description, workload))
            except Exception as e:
                print(f"解析行失败：{line} - {e}")
                continue
    
    return matches


def match_functions_with_ai(requirement_content: str, function_codes: List[Tuple[str, str, str]], total_workload: float) -> List[Tuple[str, str, str, str, float]]:
    """使用AI匹配需求内容与功能点码值"""
    codes_text = "\n".join([f"{i+1}. {l1} -> {l2} -> {l3}" for i, (l1, l2, l3) in enumerate(function_codes)])
    
    prompt = f"""基于以下需求内容，从功能点码值中选择最恰当的功能点进行匹配。

需求内容：
{requirement_content}

可用功能点码值：
{codes_text}

请按照以下格式返回匹配结果，每行一个匹配项：
功能点编号|一级功能点|二级功能点|三级功能点|功能描述|工作量估计

要求：
1. 根据需求内容的复杂度，选择2-3个最相关的功能点，最多选择5个
2. 为每个功能点提供简洁的功能描述（从需求内容中提取相关部分）
3. 根据功能复杂度估计工作量（人天），总和应接近{total_workload}人天
4. 功能点编号从1开始递增

示例格式：
1|市场洞察|建筑视角|建筑查询|实现建筑信息查询功能|3.0
2|...|...|...|...|..."""

    headers = {
        "Content-Type": "application/json",
        "Authorization": f"Bearer {DEEPSEEK_API_KEY}"
    }
    
    data = {
        "model": "deepseek-chat",
        "messages": [
            {
                "role": "user",
                "content": prompt
            }
        ],
        "max_tokens": 1500,
        "temperature": 0.7
    }
    
    try:
        print("正在调用DeepSeek API进行功能点匹配...")
        response = requests.post(DEEPSEEK_API_URL, headers=headers, json=data, timeout=30)
        
        if response.status_code != 200:
            error_msg = f"API调用失败，状态码: {response.status_code}"
            print(error_msg)
            raise Exception(error_msg)
        
        result = response.json()
        
        if 'choices' in result and len(result['choices']) > 0:
            api_response = result['choices'][0]['message']['content'].strip()
            print("✅ AI匹配成功")
            return parse_ai_function_matches(api_response, function_codes)
        else:
            error_msg = "API响应格式错误"
            print(error_msg)
            raise Exception(error_msg)
            
    except Exception as e:
        error_msg = f"调用AI匹配失败：{e}"
        print(error_msg)
        raise Exception(error_msg)


def update_wbs_document() -> None:
    print_step("第十步：基于A4数据和功能点码值更新WBS工作量文档")
    
    try:
        # 获取附件4的A4内容
        path4 = find_attachment_by_number(4)
        if not path4:
            print("未找到附件4文件")
            return
        
        wb4 = load_workbook(path4, data_only=True)
        ws4 = wb4.active
        a4_content = ws4['A4'].value
        d7_workload = ws4['D7'].value or 19.0
        
        if not a4_content:
            print("附件4的A4单元格为空")
            return
        
        print(f"提取到需求内容：{len(str(a4_content))} 字符")
        print(f"工作量总和：{d7_workload} 人天")
        
        # 加载功能点码值
        function_codes = load_function_codes()
        if not function_codes:
            print("无法加载功能点码值，程序终止")
            return
        
        # AI匹配功能点
        matches = match_functions_with_ai(str(a4_content), function_codes, d7_workload)
        
        if not matches:
            print("AI匹配失败，程序终止")
            return
        
        print(f"匹配到 {len(matches)} 个功能点")
        
        # 更新WBS文档
        path2 = find_attachment_by_number(2)
        if not path2:
            print("未找到附件2 WBS文件")
            return
        
        wb2 = load_workbook(path2)
        ws2 = wb2.active
        
        # 取消所有合并的单元格
        merged_ranges = list(ws2.merged_cells.ranges)
        for merged_range in merged_ranges:
            ws2.unmerge_cells(str(merged_range))
        
        # 清空现有数据（保留标题行）
        for row in range(2, ws2.max_row + 10):  # 多清理几行确保完全清空
            for col in range(1, 7):
                ws2.cell(row, col).value = None
        
        # 填入匹配的功能点数据
        current_row = 2
        for i, (level1, level2, level3, description, workload) in enumerate(matches):
            ws2.cell(current_row, 1).value = f"=ROW()-1"  # 编号使用公式
            ws2.cell(current_row, 2).value = level1  # 一级功能点
            ws2.cell(current_row, 3).value = level2  # 二级功能点
            ws2.cell(current_row, 4).value = level3  # 三级功能点
            ws2.cell(current_row, 5).value = description  # 功能描述
            ws2.cell(current_row, 6).value = workload  # 工作量
            current_row += 1
        
        # 添加合计行
        total_row = current_row
        ws2.cell(total_row, 1).value = f"=ROW()-1"  # 编号
        
        # 合并B列和E列单元格并填入"合计"
        ws2.merge_cells(f'B{total_row}:E{total_row}')
        ws2.cell(total_row, 2).value = "合计"
        
        # F列工作量总和
        ws2.cell(total_row, 6).value = d7_workload
        
        wb2.save(path2)
        
        print(f"已更新 {os.path.basename(path2)}:")
        print(f"  - 填入 {len(matches)} 个功能点")
        print(f"  - 工作量总和: {d7_workload} 人天")
        print(f"  - 添加合计行")
        
    except Exception as e:
        print(f"第十步执行失败：{e}")
        print("程序终止，请检查API配置或网络连接")
        raise


def main() -> None:
    print_step("输入变量：统一替换的需求名")
    requirement_name = input("请输入需求名字（用于重命名与单元格填充）：").strip()
    if not requirement_name:
        print("未输入需求名字，程序结束。")
        return

    # 1) 批量重命名
    batch_rename(requirement_name)

    # 2) 附件3 sheet2 A3/B3
    write_attachment3_sheet2_cells(requirement_name)

    # 3) 附件4 B2
    write_attachment4_cells(requirement_name)

    # 4) 计算附件5 L列总和
    total = sum_attachment5_col_L_from_L2()

    # 5) 附件4 D7 = total
    write_attachment4_with_sum(total)

    # 6) 附件3 sheet2 F3 = total
    write_attachment3_sheet2_F3_with_sum(total)

    # 7) 附件4 C6 = 今天日期
    write_attachment4_C6_with_today()

    # 8) 附件4 B7 = 附件3 sheet2 E3
    write_attachment4_B7_from_attachment3_E3()

    # 9) 附件4 A4 = 附件5 H和I列内容概述
    summarize_requirement_content_and_update_h4()

    # 10) 更新WBS文档
    update_wbs_document()

    print_step("全部步骤完成")


if __name__ == "__main__":
    main() 