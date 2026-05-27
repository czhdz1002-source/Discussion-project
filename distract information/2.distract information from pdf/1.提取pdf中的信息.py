import os
import re
import pdfplumber
import csv
from collections import defaultdict

def extract_pdf_text(pdf_path):
    """提取PDF文本内容"""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            full_text = ""
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    full_text += text + "\n"
            return full_text
    except Exception as e:
        print(f"   [!] 读取PDF出错: {e}")
        return ""

def extract_by_keywords(text, patterns):
    """通过关键词模式提取信息"""
    for pattern in patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            # 返回匹配的第一个分组
            if match.lastindex:
                return match.group(1).strip(' :：,，')
            else:
                return match.group().strip(' :：,，')
    return ""

def extract_security_info(text):
    """提取证券相关信息"""
    result = {}
    
    # 证券代码：6位数字
    code_patterns = [
        r'证券代码[：:]\s*(\d{6})',
        r'股票代码[：:]\s*(\d{6})',
        r'代码[：:]\s*(\d{6})',
        r'(\d{6})\s*证券代码',
        r'(\d{6})\s*股票代码'
    ]
    result["证券代码"] = extract_by_keywords(text, code_patterns)
    
    # 证券简称：可能包含ST等前缀
    short_patterns = [
        r'证券简称[：:]\s*([\u4e00-\u9fa5A-Z]+)',
        r'股票简称[：:]\s*([\u4e00-\u9fa5A-Z]+)',
        r'简称[：:]\s*([\u4e00-\u9fa5A-Z]+)',
        r'([\u4e00-\u9fa5A-Z]+)\s*证券简称',
        r'([\u4e00-\u9fa5A-Z]+)\s*股票简称'
    ]
    result["证券简称"] = extract_by_keywords(text, short_patterns)
    
    # 公告编号：格式如2018-01
    notice_patterns = [
        r'公告编号[：:]\s*(\d{4}-\d+)',
        r'公告号[：:]\s*(\d{4}-\d+)',
        r'编号[：:]\s*(\d{4}-\d+)',
        r'(\d{4}-\d+)\s*公告编号',
        r'(\d{4}-\d+)\s*公告号'
    ]
    result["公告编号"] = extract_by_keywords(text, notice_patterns)
    
    # 公告时间：多种日期格式
    date_patterns = [
        r'公告(?:日|期|时间|日期)[：:]\s*(\d{4})[年\-/\.](\d{1,2})[月\-/\.](\d{1,2})日?',
        r'公告(?:日|期|时间|日期)[：:]\s*(\d{4})[年\-/\.](\d{1,2})[月\-/\.](\d{1,2})',
        r'日期[：:]\s*(\d{4})[年\-/\.](\d{1,2})[月\-/\.](\d{1,2})日?',
        r'时间[：:]\s*(\d{4})[年\-/\.](\d{1,2})[月\-/\.](\d{1,2})日?',
        r'(\d{4})年(\d{1,2})月(\d{1,2})日\s*公告',
        r'(\d{4})-(\d{1,2})-(\d{1,2})\s*公告'
    ]
    
    for pattern in date_patterns:
        match = re.search(pattern, text, re.IGNORECASE)
        if match:
            groups = match.groups()
            if len(groups) >= 3:
                year, month, day = groups[0], groups[1], groups[2]
                result["公告时间"] = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                break
    
    return result

def extract_case_numbers(text):
    """提取所有案号"""
    case_patterns = [
        r'案号[：:]\s*([^\s，,。；;\n]{5,40})',
        r'案件编号[：:]\s*([^\s，,。；;\n]{5,40})',
        r'编号[：:]\s*([^\s，,。；;\n]{5,40})',
        r'案号[：:]\s*([(（][^）)]+[）)][^，,。；;\n]{3,20})',
        r'案件编号[：:]\s*([(（][^）)]+[）)][^，,。；;\n]{3,20})'
    ]
    
    case_numbers = []
    for pattern in case_patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            if isinstance(match, str):
                case_no = match.strip(' :：,，')
                # 清理多余空格
                case_no = re.sub(r'\s+', ' ', case_no)
                if case_no and case_no not in case_numbers:
                    case_numbers.append(case_no)
    
    # 去重并返回
    unique_cases = []
    for case in case_numbers:
        if len(case) > 4 and case not in unique_cases:
            unique_cases.append(case)
    
    return unique_cases

def extract_debtor_info_for_case(text, case_number):
    """针对特定案号提取失信被执行人信息"""
    result = {
        "失信被执行人姓名/名称": "",
        "立案时间": ""
    }
    
    # 在案号附近查找相关上下文（前后200字符）
    case_pos = text.find(case_number)
    if case_pos == -1:
        return result
    
    # 提取案号附近的上下文
    start = max(0, case_pos - 200)
    end = min(len(text), case_pos + len(case_number) + 200)
    context = text[start:end]
    
    # 在上下文中查找失信被执行人信息
    debtor_patterns = [
        r'失信被执行人(?:姓名|名称)?[：:]\s*([^，,。；;\n]{2,30})',
        r'被执行人(?:姓名|名称)?[：:]\s*([^，,。；;\n]{2,30})',
        r'姓名[：:]\s*([^，,。；;\n]{2,10})',
        r'名称[：:]\s*([^，,。；;\n]{2,30})',
        r'当事人[：:]\s*([^，,。；;\n]{2,30})'
    ]
    
    # 查找所有可能的匹配
    for pattern in debtor_patterns:
        matches = re.findall(pattern, context)
        for match in matches:
            if isinstance(match, str) and match:
                # 检查是否为公司
                if "公司" in match or "有限" in match or "集团" in match:
                    result["失信被执行人姓名/名称"] = match.strip(' :：,，')
                    break
                # 检查是否为个人（2-4个汉字）
                elif re.match(r'^[\u4e00-\u9fa5]{2,4}$', match.strip()):
                    # 在全文范围内检查该姓名是否出现多次
                    name = match.strip()
                    count = text.count(name)
                    if count >= 2:
                        result["失信被执行人姓名/名称"] = name
                        break
        if result["失信被执行人姓名/名称"]:
            break
    
    # 如果没有找到，尝试其他方法
    if not result["失信被执行人姓名/名称"]:
        # 查找在案号前50个字符内的人名
        before_case = text[max(0, case_pos-50):case_pos]
        name_match = re.search(r'([\u4e00-\u9fa5]{2,4})[^，,。；;\n]{0,10}案号', before_case)
        if name_match:
            result["失信被执行人姓名/名称"] = name_match.group(1)
    
    # 提取立案时间
    case_date_patterns = [
        r'立案时间[：:]\s*(\d{4})[年\-/\.](\d{1,2})[月\-/\.](\d{1,2})日?',
        r'立案[：:]\s*(\d{4})[年\-/\.](\d{1,2})[月\-/\.](\d{1,2})日?',
        r'立案日期[：:]\s*(\d{4})[年\-/\.](\d{1,2})[月\-/\.](\d{1,2})日?',
        r'于\s*(\d{4})[年\-/\.](\d{1,2})[月\-/\.](\d{1,2})日?\s*立案'
    ]
    
    # 先在上下文中查找
    for pattern in case_date_patterns:
        match = re.search(pattern, context)
        if match:
            year, month, day = match.groups()
            result["立案时间"] = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
            break
    
    # 如果上下文中没找到，在整个文本中查找
    if not result["立案时间"]:
        for pattern in case_date_patterns:
            match = re.search(pattern, text)
            if match:
                year, month, day = match.groups()
                result["立案时间"] = f"{year}-{month.zfill(2)}-{day.zfill(2)}"
                break
    
    return result

def extract_director_info(text):
    """提取董事/高管信息（用于第二类PDF）"""
    # 查找董事模式
    director_patterns = [
        r'([\u4e00-\u9fa5]{2,4})董事',
        r'董事([\u4e00-\u9fa5]{2,4})',
        r'([\u4e00-\u9fa5]{2,4})先生',
        r'([\u4e00-\u9fa5]{2,4})女士',
        r'高级管理人员([\u4e00-\u9fa5]{2,4})',
        r'监事([\u4e00-\u9fa5]{2,4})'
    ]
    
    directors = []
    for pattern in director_patterns:
        matches = re.findall(pattern, text)
        for match in matches:
            if isinstance(match, str) and 2 <= len(match) <= 4:
                # 检查该姓名在全文出现的频率
                count = text.count(match)
                if count >= 2 and match not in directors:
                    directors.append(match)
    
    return directors

def process_pdf_file(pdf_path, filename):
    """处理单个PDF文件"""
    print(f"处理: {filename}")
    
    # 提取文本
    text = extract_pdf_text(pdf_path)
    if not text:
        print("   [!] 无法提取文本内容")
        return []
    
    # 提取公共信息
    security_info = extract_security_info(text)
    print(f"   证券代码: {security_info.get('证券代码', '未提取')}")
    print(f"   证券简称: {security_info.get('证券简称', '未提取')}")
    print(f"   公告编号: {security_info.get('公告编号', '未提取')}")
    print(f"   公告时间: {security_info.get('公告时间', '未提取')}")
    
    # 提取案号
    case_numbers = extract_case_numbers(text)
    
    results = []
    
    if case_numbers:
        print(f"   发现 {len(case_numbers)} 个案号")
        
        # 为每个案号创建一条记录
        for i, case_no in enumerate(case_numbers):
            print(f"   案号 {i+1}: {case_no}")
            
            # 提取该案号对应的失信被执行人信息
            case_info = extract_debtor_info_for_case(text, case_no)
            
            result = {
                "pdf文件名": filename,
                "证券代码": security_info.get("证券代码", ""),
                "证券简称": security_info.get("证券简称", ""),
                "公告编号": security_info.get("公告编号", ""),
                "公告时间": security_info.get("公告时间", ""),
                "案号": case_no,
                "失信被执行人姓名/名称": case_info.get("失信被执行人姓名/名称", ""),
                "立案时间": case_info.get("立案时间", ""),
                "备注": ""
            }
            
            results.append(result)
        
        # 如果没有提取到失信被执行人姓名，尝试从文件名提取
        if all(not r["失信被执行人姓名/名称"] for r in results):
            name_from_filename = re.search(r'关于(.+?)(?:被纳入|的公告)', filename)
            if name_from_filename:
                name = name_from_filename.group(1).strip()
                for result in results:
                    if not result["失信被执行人姓名/名称"]:
                        result["失信被执行人姓名/名称"] = name
        
        # 标记多个失信人的情况
        if len(case_numbers) > 1:
            for result in results:
                result["备注"] = "存在多个失信人需人工检验"
    
    else:
        print("   未发现案号")
        
        # 提取董事信息（第二类PDF）
        directors = extract_director_info(text)
        
        if directors:
            print(f"   发现 {len(directors)} 个董事/高管姓名")
            # 为每个董事创建一条记录
            for director in directors:
                result = {
                    "pdf文件名": filename,
                    "证券代码": security_info.get("证券代码", ""),
                    "证券简称": security_info.get("证券简称", ""),
                    "公告编号": security_info.get("公告编号", ""),
                    "公告时间": security_info.get("公告时间", ""),
                    "案号": "",
                    "失信被执行人姓名/名称": director,
                    "立案时间": "",
                    "备注": "董事/高管信息（非失信人）"
                }
                results.append(result)
        else:
            # 没有案号和董事信息，只保存公共信息
            result = {
                "pdf文件名": filename,
                "证券代码": security_info.get("证券代码", ""),
                "证券简称": security_info.get("证券简称", ""),
                "公告编号": security_info.get("公告编号", ""),
                "公告时间": security_info.get("公告时间", ""),
                "案号": "",
                "失信被执行人姓名/名称": "",
                "立案时间": "",
                "备注": "无案号信息"
            }
            results.append(result)
    
    print()
    return results

def batch_process_pdfs(pdf_folder, output_csv, limit=50):
    """批量处理PDF文件"""
    all_results = []
    
    # 获取PDF文件列表
    pdf_files = []
    for f in os.listdir(pdf_folder):
        if f.lower().endswith('.pdf'):
            pdf_files.append(f)
    
    # 按文件名排序
    pdf_files.sort()
    
    total_files = len(pdf_files)
    process_count = min(limit, total_files) if limit > 0 else total_files
    
    print(f"找到 {total_files} 个PDF文件，开始处理前 {process_count} 个...")
    print("=" * 80)
    
    for i, filename in enumerate(pdf_files[:process_count]):
        pdf_path = os.path.join(pdf_folder, filename)
        print(f"[{i+1}/{process_count}] ", end="")
        
        # 处理单个PDF
        file_results = process_pdf_file(pdf_path, filename)
        all_results.extend(file_results)
    
    # 保存结果到CSV
    if all_results:
        # 确保输出目录存在
        output_dir = os.path.dirname(output_csv)
        if output_dir and not os.path.exists(output_dir):
            os.makedirs(output_dir)
        
        with open(output_csv, 'w', newline='', encoding='utf-8-sig') as f:
            fieldnames = [
                'pdf文件名', '证券代码', '证券简称', '公告编号', 
                '公告时间', '案号', '失信被执行人姓名/名称', '立案时间', '备注'
            ]
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(all_results)
        
        # 生成统计报告
        total_rows = len(all_results)
        files_with_cases = sum(1 for r in all_results if r["案号"]) // max(1, len(set(r["pdf文件名"] for r in all_results)))
        
        # 统计提取成功率
        extracted_counts = defaultdict(int)
        total_records = len(all_results)
        
        for result in all_results:
            for field in ['证券代码', '证券简称', '公告编号', '公告时间', '案号', '失信被执行人姓名/名称', '立案时间']:
                if result[field]:
                    extracted_counts[field] += 1
        
        print("=" * 80)
        print("📊 处理完成! 统计报告:")
        print(f"   处理文件总数: {len(pdf_files[:process_count])}")
        print(f"   生成记录总数: {total_rows}")
        print(f"   平均每文件记录数: {total_rows/len(pdf_files[:process_count]):.1f}")
        print("\n   字段提取成功率:")
        for field in ['证券代码', '证券简称', '公告编号', '公告时间', '案号', '失信被执行人姓名/名称', '立案时间']:
            success_rate = extracted_counts[field] / total_records * 100
            print(f"     {field}: {extracted_counts[field]}/{total_records} ({success_rate:.1f}%)")
        
        print(f"\n   结果文件: {os.path.abspath(output_csv)}")
        print("=" * 80)
        
        # 打印前10条记录示例
        print("\n📋 提取结果示例 (前10条记录):")
        print("-" * 120)
        header = f"{'文件名':<25} {'代码':<8} {'简称':<8} {'公告编号':<12} {'公告时间':<12} {'案号':<20} {'姓名/名称':<15} {'立案时间':<12}"
        print(header)
        print("-" * 120)
        
        for i, result in enumerate(all_results[:10]):
            line = f"{result['pdf文件名'][:22]:<25} "
            line += f"{result['证券代码'] or '无':<8} "
            line += f"{result['证券简称'] or '无':<8} "
            line += f"{result['公告编号'] or '无':<12} "
            line += f"{result['公告时间'] or '无':<12} "
            line += f"{result['案号'][:18] if result['案号'] else '无':<20} "
            line += f"{result['失信被执行人姓名/名称'][:12] if result['失信被执行人姓名/名称'] else '无':<15} "
            line += f"{result['立案时间'] or '无':<12}"
            print(line)
        
        print("-" * 120)

# ====== 主程序配置 ======
if __name__ == "__main__":
    # 配置参数
    PDF_FOLDER = r"D:\This computer\desktop\失信人文件（提取信息用）"
    OUTPUT_CSV = r"D:\This computer\desktop\失信人信息提取结果_案号为核心.csv"
    PROCESS_LIMIT = 30  # 先处理30个进行测试
    
    # 检查文件夹是否存在
    if not os.path.exists(PDF_FOLDER):
        print(f"❌ 错误: 文件夹不存在 - {PDF_FOLDER}")
        print("请检查路径是否正确")
    else:
        # 运行批处理
        batch_process_pdfs(PDF_FOLDER, OUTPUT_CSV, PROCESS_LIMIT)
