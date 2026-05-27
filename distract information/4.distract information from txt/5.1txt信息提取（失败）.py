import os
import re
import docx
import csv

def load_surnames(surname_file_path):
    """
    从姓氏表docx文件中加载姓氏
    """
    surnames = set()
    
    try:
        # 读取docx文件
        doc = docx.Document(surname_file_path)
        
        # 遍历文档中的所有段落
        for para in doc.paragraphs:
            text = para.text.strip()
            if text:
                # 按"、"分隔姓氏
                parts = text.split('、')
                for part in parts:
                    surname = part.strip()
                    # 只保留单个字的姓氏（中文姓氏通常为单字）
                    if surname and len(surname) == 1:
                        surnames.add(surname)
        
        # 如果从文件中提取失败，使用常见的中国姓氏作为备选
        if not surnames:
            common_surnames = [
                '王', '李', '张', '刘', '陈', '杨', '黄', '吴', '赵', '周',
                '徐', '朱', '孙', '马', '胡', '林', '郭', '何', '郑', '高',
                '罗', '梁', '谢', '宋', '许', '唐', '冯', '邓', '韩', '曹',
                '叶', '曾', '彭', '潘', '蔡', '董', '蒋', '肖', '袁', '苏',
                '余', '程', '田', '丁', '杜', '沈', '魏', '卢', '吕', '于',
                '邬', '倪', '王', '朱', '林', '吴', '郑', '刘'
            ]
            surnames.update(common_surnames)
            
    except Exception as e:
        print(f"加载姓氏文件时出错: {e}")
        # 使用常见姓氏作为备选
        common_surnames = [
            '王', '李', '张', '刘', '陈', '杨', '黄', '吴', '赵', '周',
            '徐', '朱', '孙', '马', '胡', '林', '郭', '何', '郑', '高',
            '罗', '梁', '谢', '宋', '许', '唐', '冯', '邓', '韩', '曹',
            '叶', '曾', '彭', '潘', '蔡', '董', '蒋', '肖', '袁', '苏',
            '余', '程', '田', '丁', '杜', '沈', '魏', '卢', '吕', '于',
            '邬', '倪', '王', '朱', '林', '吴', '郑', '刘'
        ]
        surnames.update(common_surnames)
    
    return list(surnames)

def normalize_content(content):
    """
    标准化内容：处理换行符和空白字符
    将换行符替换为空格，并合并多个空白字符
    """
    # 将各种换行符替换为空格
    content = content.replace('\r\n', ' ')
    content = content.replace('\n', ' ')
    content = content.replace('\r', ' ')
    
    # 将多个连续空格合并为一个空格
    content = re.sub(r'\s+', ' ', content)
    
    return content

def find_most_common_substring(strings):
    """
    找出字符串列表中出现次数最多的最长公共子串
    不要求所有字符串都包含，只要求是多个字符串的公共子串
    """
    if not strings:
        return ""
    
    all_strings = strings
    substring_counts = {}
    
    # 分析每个字符串的所有可能2-4字符子串
    for s in all_strings:
        length = len(s)
        if length < 2:
            continue
            
        # 为该字符串生成所有可能的2-4字符子串
        substrings = set()
        for sub_len in range(2, min(5, length + 1)):
            for i in range(0, length - sub_len + 1):
                substring = s[i:i + sub_len]
                if re.match(r'^[\u4e00-\u9fa5]+$', substring):
                    substrings.add(substring)
        
        # 统计这些子串
        for substring in substrings:
            substring_counts[substring] = substring_counts.get(substring, 0) + 1
    
    if not substring_counts:
        return ""
    
    # 找出出现次数最多的子串
    max_count = max(substring_counts.values())
    
    # 找出所有出现次数最多的子串
    most_common_substrings = [sub for sub, count in substring_counts.items() 
                             if count == max_count]
    
    if not most_common_substrings:
        return ""
    
    # 如果有多个出现次数相同的子串，选择最长的那个
    most_common_substrings.sort(key=lambda x: len(x), reverse=True)
    
    return most_common_substrings[0]

def find_common_name_from_extracted_results(extracted_names):
    """
    从提取到的姓名结果中找出最长公共子串作为姓名
    规则：找出所有提取结果中最长的公共子串
    然后检查这个公共子串在多少个提取结果中出现
    """
    if not extracted_names:
        return []
    
    # 去除重复项
    unique_names = list(set(extracted_names))
    
    if len(unique_names) == 1:
        # 只有一个结果，直接返回
        return unique_names
    
    # 找出出现次数最多的公共子串
    common_substring = find_most_common_substring(unique_names)
    
    if not common_substring or len(common_substring) < 2:
        # 如果没有找到有效的公共子串，返回所有唯一结果
        return unique_names
    
    # 统计公共子串在多少个提取结果中出现
    substring_count = sum(1 for name in unique_names if common_substring in name)
    
    # 如果公共子串在超过一个结果中出现，则返回公共子串
    # 否则返回所有唯一结果
    if substring_count > 1:
        # 进一步检查这个公共子串是否可能是姓名
        # 检查长度是否符合姓名要求
        if 2 <= len(common_substring) <= 4:
            return [common_substring]
        else:
            # 如果公共子串长度不符合姓名要求，返回所有唯一结果
            return unique_names
    else:
        # 公共子串只在一个结果中出现，返回所有唯一结果
        return unique_names

def extract_names_with_multiple_methods(content, surnames):
    """
    同时使用多种方法提取失信被执行人姓名
    """
    # 首先标准化内容，处理换行符问题
    normalized_content = normalize_content(content)
    
    all_names = []  # 收集所有方法提取的结果
    
    # 方法1: 查找"失信被执行人（姓名/名称）："模式
    method1_names = extract_names_method1(normalized_content, surnames)
    all_names.extend(method1_names)
    
    # 方法2: 查找"被纳入失信被执行人名单"前的姓名
    method2_names = extract_names_method2(normalized_content, surnames)
    all_names.extend(method2_names)
    
    # 方法3: 查找"失信被执行人名称："模式
    method3_names = extract_names_method3(normalized_content, surnames)
    all_names.extend(method3_names)
    
    # 方法4: 查找表格格式的失信人信息
    method4_names = extract_names_method4(normalized_content, surnames)
    all_names.extend(method4_names)
    
    # 方法5: 专门处理有换行情况的关键标识
    method5_names = extract_names_with_line_breaks(content, surnames)
    all_names.extend(method5_names)
    
    # 验证所有候选姓名
    validated_names = []
    for name in all_names:
        if is_valid_name(name, surnames, normalized_content):
            validated_names.append(name)
    
    # 如果验证后没有姓名，直接返回空列表
    if not validated_names:
        return []
    
    # 找出公共姓名部分
    common_names = find_common_name_from_extracted_results(validated_names)
    
    return common_names

def extract_names_method1(content, surnames):
    """
    方法1: 查找"失信被执行人（姓名/名称）："模式
    """
    names = []
    
    patterns = [
        r'失信被执行人（姓名/名称）[：:]\s*([^\s，。；;、]+)',
        r'失信被执行人\(姓名/名称\)[：:]\s*([^\s，。；;、]+)',
        r'失信被执行人（姓名/名称）[：:]\s*([^，。；;、]{2,4})',
        r'失信被执行人\(姓名/名称\)[：:]\s*([^，。；;、]{2,4})',
    ]
    
    for pattern in patterns:
        matches = re.finditer(pattern, content)
        for match in matches:
            name = match.group(1).strip()
            if is_valid_name(name, surnames, content):
                names.append(name)
    
    return names

def extract_names_method2(content, surnames):
    """
    方法2: 查找"被纳入失信被执行人名单"前的姓名
    处理标准化后的内容（已处理换行）
    """
    names = []
    
    # 处理标准化后的内容，换行已被替换为空格
    patterns = [
        # 标准模式，用于处理已标准化的内容
        r'公司董事\s*([\u4e00-\u9fa5]{2,4})\s*被纳入失信被执行人名单',
        r'董事\s*([\u4e00-\u9fa5]{2,4})\s*被纳入失信被执行人名单',
        r'([\u4e00-\u9fa5]{2,4})\s*被纳入失信被执行人名单',
    ]
    
    for pattern in patterns:
        matches = re.finditer(pattern, content)
        for match in matches:
            name = match.group(1).strip()
            if is_valid_name(name, surnames, content):
                names.append(name)
    
    return names

def extract_names_method3(content, surnames):
    """
    方法3: 查找"失信被执行人名称："模式
    """
    names = []
    
    patterns = [
        r'失信被执行人名称[：:]\s*([^\s，。；;、]+)',
        r'失信被执行人名称[：:]\s*([^，。；;、]{2,4})',
        r'失信被执行人姓名[：:]\s*([^\s，。；;、]+)',
    ]
    
    for pattern in patterns:
        matches = re.finditer(pattern, content)
        for match in matches:
            name = match.group(1).strip()
            if is_valid_name(name, surnames, content):
                names.append(name)
    
    return names

def extract_names_method4(content, surnames):
    """
    方法4: 查找表格格式的失信人信息
    """
    names = []
    
    pattern = r'失信被执行人（姓名/名称）[：:]\s*([^，。；;、\s]{2,4})'
    matches = re.finditer(pattern, content)
    for match in matches:
        name = match.group(1).strip()
        if is_valid_name(name, surnames, content):
            names.append(name)
    
    pattern2 = r'失信被执行人名称[：:]\s*([^，。；;、\s]{2,4})'
    matches = re.finditer(pattern2, content)
    for match in matches:
        name = match.group(1).strip()
        if is_valid_name(name, surnames, content):
            names.append(name)
    
    return names

def extract_names_with_line_breaks(original_content, surnames):
    """
    方法5: 专门处理关键标识中有换行的情况
    直接处理原始内容，不进行标准化，以处理各种换行情况
    """
    names = []
    
    # 关键标识的各个部分，允许任意换行
    # 匹配模式：姓名 + 任意空白/换行 + "被" + 任意空白/换行 + "纳入" + 任意空白/换行 + "失信" + 任意空白/换行 + "被" + 任意空白/换行 + "执" + 任意空白/换行 + "行人" + 任意空白/换行 + "名单"
    
    # 使用更灵活的正则表达式，允许任意空白字符（包括换行）
    patterns = [
        # 模式1: 董事 + 姓名 + 被纳入失信被执行人名单（允许任意位置换行）
        r'董事\s*([\u4e00-\u9fa5]{2,4})\s*被\s*纳\s*入\s*失\s*信\s*被\s*执\s*行\s*人\s*名\s*单',
        r'公司董事\s*([\u4e00-\u9fa5]{2,4})\s*被\s*纳\s*入\s*失\s*信\s*被\s*执\s*行\s*人\s*名\s*单',
        
        # 模式2: 简化的模式，匹配"被纳入失信"和"被执行人名单"之间有换行的情况
        r'([\u4e00-\u9fa5]{2,4})被\s*纳入\s*失信\s*被\s*执\s*行人\s*名单',
    ]
    
    for pattern in patterns:
        matches = re.finditer(pattern, original_content, re.DOTALL)
        for match in matches:
            # 获取捕获组中的姓名
            if match.lastindex:
                name = match.group(1).strip()
            else:
                # 如果没有捕获组，尝试从匹配文本中提取姓名
                matched_text = match.group(0).strip()
                # 查找2-4个中文字符作为姓名
                name_match = re.search(r'([\u4e00-\u9fa5]{2,4})', matched_text)
                name = name_match.group(1) if name_match else matched_text
            
            if is_valid_name(name, surnames, original_content):
                names.append(name)
    
    # 另外尝试提取董事姓名
    director_patterns = [
        # 查找"董事"后面的姓名，后面跟着"被纳入失信被执行人名单"的关键词
        r'董事\s*([\u4e00-\u9fa5]{2,4})\s*(?:(?:.{0,50}?被.{0,50}?纳入.{0,50}?失信.{0,50}?被.{0,50}?执.{0,50}?行人.{0,50}?名单))',
        r'公司董事\s*([\u4e00-\u9fa5]{2,4})\s*(?:(?:.{0,50}?被.{0,50}?纳入.{0,50}?失信.{0,50}?被.{0,50}?执.{0,50}?行人.{0,50}?名单))',
    ]
    
    for pattern in director_patterns:
        matches = re.finditer(pattern, original_content, re.DOTALL)
        for match in matches:
            if match.lastindex:
                name = match.group(1).strip()
                if is_valid_name(name, surnames, original_content):
                    names.append(name)
    
    return names

def extract_names_from_context(context_text, surnames, full_content):
    """
    从上下文文本中提取可能的姓名
    """
    names = []
    
    if is_valid_name(context_text, surnames, full_content):
        names.append(context_text)
    
    surname_regex = '|'.join([re.escape(surname) for surname in surnames])
    
    name_pattern = rf'({surname_regex}[\u4e00-\u9fa5]{{1,3}})'
    
    matches = re.finditer(name_pattern, context_text)
    for match in matches:
        name = match.group(1).strip()
        if is_valid_name(name, surnames, full_content):
            names.append(name)
    
    return names

def is_valid_name(name, surnames, content):
    """
    检查是否是有效的姓名
    """
    if not name or len(name) < 2 or len(name) > 4:
        return False
    
    # 检查是否以姓氏开头
    starts_with_surname = any(name.startswith(surname) for surname in surnames)
    if not starts_with_surname:
        return False
    
    # 只包含中文汉字
    if not re.match(r'^[\u4e00-\u9fa5]+$', name):
        return False
    
    # 过滤掉常见的非姓名词汇
    invalid_words = ['公司', '集团', '企业', '股份', '有限', '责任', '法院', '银行', '证券', '保险']
    if any(word in name for word in invalid_words):
        return False
    
    # 过滤掉职位名称
    positions = ['董事', '主席', '监事', '经理', '主任', '科长', '处长', '局长', '部长', '总裁', '总监']
    if name in positions:
        return False
    
    # 检查姓名在全文中的出现次数
    name_count = content.count(name)
    if name_count < 1:
        return False
    
    return True

def extract_info_from_file(file_path, surnames):
    """
    从单个txt文件中提取信息
    """
    with open(file_path, 'r', encoding='utf-8') as f:
        content = f.read()
    
    result = {
        '证券代码': '',
        '证券简称': '',
        '公告编号': '',
        '公告时间': '',
        '失信被执行人姓名': ''
    }
    
    # 提取证券代码
    code_match = re.search(r'证券代码[：:]\s*(\d{6})', content)
    if code_match:
        result['证券代码'] = code_match.group(1)
    
    # 提取证券简称
    name_match = re.search(r'证券简称[：:]\s*([^\s]+)', content)
    if name_match:
        result['证券简称'] = name_match.group(1)
    
    # 提取公告编号
    bulletin_match = re.search(r'公告编号[：:]\s*([^\s]+)', content)
    if bulletin_match:
        result['公告编号'] = bulletin_match.group(1)
    
    # 提取公告时间
    date_patterns = [
        r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日',
        r'(\d{4})-(\d{1,2})-(\d{1,2})',
        r'(\d{4})/(\d{1,2})/(\d{1,2})',
        r'(\d{4})\.(\d{1,2})\.(\d{1,2})'
    ]
    
    lines = content.strip().split('\n')
    for line in reversed(lines[-10:]):
        for pattern in date_patterns:
            date_match = re.search(pattern, line)
            if date_match:
                year, month, day = date_match.groups()
                try:
                    result['公告时间'] = f"{year}年{int(month)}月{int(day)}日"
                    break
                except:
                    continue
        if result['公告时间']:
            break
    
    # 提取失信被执行人姓名
    names = extract_names_with_multiple_methods(content, surnames)
    
    if names:
        unique_names = list(set(names))
        unique_names.sort()
        result['失信被执行人姓名'] = '，'.join(unique_names)
    
    return result

def extract_all_files(folder_path, surname_file_path):
    """
    提取文件夹中所有txt文件的信息
    """
    all_results = []
    
    surnames = load_surnames(surname_file_path)
    print(f"加载了 {len(surnames)} 个姓氏")
    
    txt_files = [f for f in os.listdir(folder_path) if f.endswith('.txt')]
    txt_files.sort()
    
    print(f"找到 {len(txt_files)} 个txt文件")
    
    for filename in txt_files:
        file_path = os.path.join(folder_path, filename)
        print(f"\n处理文件: {filename}")
        
        try:
            result = extract_info_from_file(file_path, surnames)
            
            record = {
                '文件名': filename,
                '证券代码': result['证券代码'],
                '证券简称': result['证券简称'],
                '公告编号': result['公告编号'],
                '公告时间': result['公告时间'],
                '失信被执行人姓名': result['失信被执行人姓名']
            }
            all_results.append(record)
            
            if result['失信被执行人姓名']:
                print(f"  提取到姓名: {result['失信被执行人姓名']}")
            else:
                print(f"  未提取到姓名")
                
        except Exception as e:
            print(f"处理文件 {filename} 时出错: {e}")
            all_results.append({
                '文件名': filename,
                '证券代码': '',
                '证券简称': '',
                '公告编号': '',
                '公告时间': '',
                '失信被执行人姓名': '',
                '错误': str(e)
            })
    
    return all_results

def save_to_csv(results, output_file):
    """
    将结果保存到CSV文件
    """
    if not results:
        print("没有数据可保存")
        return
    
    fieldnames = list(results[0].keys())
    
    try:
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            writer.writerows(results)
        
        print(f"\n结果已保存到: {output_file}")
    except PermissionError:
        print(f"\n[错误] 无法保存到 {output_file}，文件可能已被其他程序打开")
        print("请关闭Excel或其他可能正在使用该文件的程序，然后重试。")

def main():
    folder_path = input("请输入包含TXT文件的文件夹路径: ")
    
    if not os.path.exists(folder_path):
        print("文件夹不存在!")
        return
    
    surname_file_path = r"D:\This computer\desktop\姓氏表\姓氏表.docx"
    
    if not os.path.exists(surname_file_path):
        print(f"姓氏文件不存在: {surname_file_path}")
        print("请确保姓氏表文件存在，或修改代码中的路径")
        return
    
    results = extract_all_files(folder_path, surname_file_path)
    
    if results:
        output_file = os.path.join(folder_path, '失信人信息提取结果_改进版.csv')
        save_to_csv(results, output_file)
        
        print(f"\n处理完成!")
        print(f"共处理 {len(results)} 条记录")
        
        extracted_count = sum(1 for r in results if r.get('失信被执行人姓名'))
        print(f"成功提取失信被执行人姓名的记录: {extracted_count}")
        
        print("\n提取到的所有失信被执行人姓名:")
        all_names = []
        for r in results:
            if r.get('失信被执行人姓名'):
                names = r['失信被执行人姓名'].split('，')
                all_names.extend([name for name in names if name])
        
        unique_names = sorted(set(all_names))
        for name in unique_names:
            print(f"  {name}")
    else:
        print("没有提取到任何数据")

if __name__ == "__main__":
    main()
