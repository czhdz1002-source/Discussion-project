import os
import re
import docx
import csv

def load_surnames(surname_file_path):
    """
    从姓氏表docx文件中加载姓氏
    修改：支持多字姓氏，加载失败时直接提示，不使用备选姓氏
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
                    if surname:  # 保留所有姓氏，不限制单字
                        surnames.add(surname)
        
        # 如果从文件中提取失败，直接提示
        if not surnames:
            print(f"警告：姓氏文件为空或未提取到任何姓氏！")
            
    except Exception as e:
        print(f"加载姓氏文件时出错: {e}")
        return []  # 返回空列表，不使用备选姓氏
    
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

def extract_names_with_multiple_methods(content, surnames):
    """
    同时使用多种方法提取失信被执行人姓名
    """
    # 首先标准化内容，处理换行符问题
    normalized_content = normalize_content(content)
    
    names_set = set()
    
    # 方法1: 查找"失信被执行人（姓名/名称）："等模式
    method1_names = extract_names_method1(normalized_content, surnames)
    names_set.update(method1_names)
    
    # 方法2: 查找"被纳入失信被执行人名单"等模式前的姓名
    method2_names = extract_names_method2(normalized_content, surnames)
    names_set.update(method2_names)
    
    # 方法3: 查找"失信被执行人名称："等模式
    method3_names = extract_names_method3(normalized_content, surnames)
    names_set.update(method3_names)
    
    # 方法4: 查找表格格式的失信人信息
    method4_names = extract_names_method4(normalized_content, surnames)
    names_set.update(method4_names)
    
    # 方法5: 专门处理有换行情况的关键标识
    method5_names = extract_names_with_line_breaks(content, surnames)
    names_set.update(method5_names)
    
    # 方法6: 查找"被列入失信被执行名单"等模式前的姓名
    method6_names = extract_names_method6(normalized_content, surnames)
    names_set.update(method6_names)
    
    # 验证所有候选姓名
    validated_names = []
    for name in names_set:
        if is_valid_name(name, surnames, normalized_content):
            validated_names.append(name)
    
    return validated_names

def extract_names_method1(content, surnames):
    """
    方法1: 查找"失信被执行人（姓名/名称）："等模式
    修改：支持"失信被执行人姓名/名称："等多种表达
    """
    names = []
    
    # 多种表达模式
    patterns = [
        # 标准模式
        r'失信被执行人（姓名/名称）[：:]\s*([^\r\n]+)',
        r'失信被执行人\(姓名/名称\)[：:]\s*([^\r\n]+)',
        # 新增：失信被执行人姓名/名称：
        r'失信被执行人姓名/名称[：:]\s*([^\r\n]+)',
        # 新增：失信被执行人姓名：
        r'失信被执行人姓名[：:]\s*([^\r\n]+)',
    ]
    
    for pattern in patterns:
        matches = re.finditer(pattern, content)
        for match in matches:
            candidate = match.group(1).strip()
            
            # 跳过公司名称
            if is_company_name(candidate, content):
                continue
                
            # 检查是否是有效姓名
            if is_valid_name(candidate, surnames, content):
                names.append(candidate)
            else:
                # 尝试从候选文本中提取姓名
                extracted_names = extract_names_from_text(candidate, surnames, content)
                names.extend(extracted_names)
    
    # 原有模式也保留
    patterns2 = [
        r'失信被执行人（姓名/名称）[：:]\s*([^\s，。；;、]+)',
        r'失信被执行人\(姓名/名称\)[：:]\s*([^\s，。；;、]+)',
        r'失信被执行人姓名/名称[：:]\s*([^\s，。；;、]+)',
        r'失信被执行人姓名[：:]\s*([^\s，。；;、]+)',
    ]
    
    for pattern in patterns2:
        matches = re.finditer(pattern, content)
        for match in matches:
            name = match.group(1).strip()
            if is_valid_name(name, surnames, content):
                names.append(name)
    
    return names

def extract_names_method2(content, surnames):
    """
    方法2: 查找"被纳入失信被执行人名单"等模式前的姓名
    修改：支持多种表达，包括"纳入失信被执行人名单"
    """
    names = []
    
    # 处理标准化后的内容，换行已被替换为空格
    patterns = [
        # 原模式
        r'公司董事\s+([\u4e00-\u9fa5]{2,4})\s+被纳入失信被执行人名单',
        r'董事\s+([\u4e00-\u9fa5]{2,4})\s+被纳入失信被执行人名单',
        r'([\u4e00-\u9fa5]{2,4})\s+被纳入失信被执行人名单',
        r'公司董事([\u4e00-\u9fa5]{2,4})被纳入失信被执行人名单',
        r'董事([\u4e00-\u9fa5]{2,4})被纳入失信被执行人名单',
        # 新增：纳入失信被执行人名单（无"被"字）
        r'公司董事\s+([\u4e00-\u9fa5]{2,4})\s+纳入失信被执行人名单',
        r'董事\s+([\u4e00-\u9fa5]{2,4})\s+纳入失信被执行人名单',
        r'([\u4e00-\u9fa5]{2,4})\s+纳入失信被执行人名单',
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
    方法3: 查找"失信被执行人名称："等模式
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

def extract_names_method6(content, surnames):
    """
    方法6: 查找"被列入失信被执行名单"等模式前的姓名
    新增：支持"被列入失信被执行名单"等表达
    """
    names = []
    
    patterns = [
        # 被列入失信被执行名单
        r'公司董事\s+([\u4e00-\u9fa5]{2,4})\s+被列入失信被执行名单',
        r'董事\s+([\u4e00-\u9fa5]{2,4})\s+被列入失信被执行名单',
        r'([\u4e00-\u9fa5]{2,4})\s+被列入失信被执行名单',
        r'公司董事([\u4e00-\u9fa5]{2,4})被列入失信被执行名单',
        r'董事([\u4e00-\u9fa5]{2,4})被列入失信被执行名单',
        # 被列入失信被执行人名单
        r'公司董事\s+([\u4e00-\u9fa5]{2,4})\s+被列入失信被执行人名单',
        r'董事\s+([\u4e00-\u9fa5]{2,4})\s+被列入失信被执行人名单',
        r'([\u4e00-\u9fa5]{2,4})\s+被列入失信被执行人名单',
    ]
    
    for pattern in patterns:
        matches = re.finditer(pattern, content)
        for match in matches:
            name = match.group(1).strip()
            if is_valid_name(name, surnames, content):
                names.append(name)
    
    return names

def extract_names_with_line_breaks(original_content, surnames):
    """
    方法5: 专门处理关键标识中有换行的情况
    直接处理原始内容，不进行标准化，以处理各种换行情况
    修改：支持多种表达和换行符
    """
    names = []
    
    # 关键标识的各个部分，允许任意空白字符（包括换行）
    patterns = [
        # 模式1: 董事 + 姓名 + 被纳入失信被执行人名单（允许任意位置换行）
        r'董事\s*([\u4e00-\u9fa5]{2,4})\s*被\s*纳\s*入\s*失\s*信\s*被\s*执\s*行\s*人\s*名\s*单',
        r'公司\s*董事\s*([\u4e00-\u9fa5]{2,4})\s*被\s*纳\s*入\s*失\s*信\s*被\s*执\s*行\s*人\s*名\s*单',
        
        # 模式2: 姓名 + 被纳入失信被执行人名单
        r'([\u4e00-\u9fa5]{2,4})\s*被\s*纳\s*入\s*失\s*信\s*被\s*执\s*行\s*人\s*名\s*单',
        
        # 模式3: 被列入失信被执行名单
        r'董事\s*([\u4e00-\u9fa5]{2,4})\s*被\s*列\s*入\s*失\s*信\s*被\s*执\s*行\s*名\s*单',
        r'公司\s*董事\s*([\u4e00-\u9fa5]{2,4})\s*被\s*列\s*入\s*失\s*信\s*被\s*执\s*行\s*名\s*单',
        r'([\u4e00-\u9fa5]{2,4})\s*被\s*列\s*入\s*失\s*信\s*被\s*执\s*行\s*名\s*单',
        
        # 模式4: 纳入失信被执行人名单（无"被"字）
        r'董事\s*([\u4e00-\u9fa5]{2,4})\s*纳\s*入\s*失\s*信\s*被\s*执\s*行\s*人\s*名\s*单',
        r'公司\s*董事\s*([\u4e00-\u9fa5]{2,4})\s*纳\s*入\s*失\s*信\s*被\s*执\s*行\s*人\s*名\s*单',
        r'([\u4e00-\u9fa5]{2,4})\s*纳\s*入\s*失\s*信\s*被\s*执\s*行\s*人\s*名\s*单',
    ]
    
    for pattern in patterns:
        matches = re.finditer(pattern, original_content, re.DOTALL)
        for match in matches:
            if match.lastindex:
                name = match.group(1).strip()
                # 检查是否包含"司"字（避免"司董事"）
                if name.startswith('司'):
                    continue
                
                if is_valid_name(name, surnames, original_content):
                    names.append(name)
    
    return names

def extract_names_from_text(text, surnames, full_content):
    """
    从文本中提取可能的姓名
    修改：支持多字姓氏，名字部分为1-2个字
    """
    names = []
    
    # 如果是有效姓名，直接添加
    if is_valid_name(text, surnames, full_content):
        names.append(text)
    
    # 构建姓氏正则表达式
    if surnames:
        # 按长度从长到短排序，确保长姓氏优先匹配
        sorted_surnames = sorted(surnames, key=len, reverse=True)
        surname_regex = '|'.join([re.escape(surname) for surname in sorted_surnames])
        
        # 名字部分为1-2个字符（姓氏可能为多字，总长度2-4）
        name_pattern = rf'({surname_regex}[\u4e00-\u9fa5]{{1,2}})'
        
        matches = re.finditer(name_pattern, text)
        for match in matches:
            name = match.group(1).strip()
            if is_valid_name(name, surnames, full_content):
                names.append(name)
    
    return names

def is_company_name(candidate, content):
    """
    判断是否是公司名称
    """
    # 检查是否包含明显的公司关键词
    company_keywords = ['公司', '股份', '有限', '集团', '企业', '责任公司', '有限公司', '有限责任公司']
    if any(keyword in candidate for keyword in company_keywords):
        return True
    
    # 检查是否是较长字符串（超过8个字符，很可能是公司名称）
    if len(candidate) > 8:
        return True
    
    # 检查是否以常见地名开头
    location_prefixes = ['北京', '上海', '广州', '深圳', '天津', '重庆', '杭州', '南京', '武汉', '成都',
                        '西安', '郑州', '沈阳', '青岛', '大连', '宁波', '厦门', '苏州', '无锡', '常州',
                        '徐州', '南通', '扬州', '镇江', '泰州', '盐城', '淮安', '连云港', '宿迁', '温州',
                        '绍兴', '嘉兴', '湖州', '金华', '衢州', '台州', '丽水', '舟山', '福州', '泉州',
                        '漳州', '莆田', '宁德', '龙岩', '三明', '南平', '济南', '青岛', '烟台', '潍坊',
                        '临沂', '济宁', '淄博', '泰安', '枣庄', '日照', '莱芜', '聊城', '德州', '滨州',
                        '菏泽', '东营', '威海', '山东', '云南', '江苏', '江西', '安徽', '福建', '江阴', 
                        '广西', '广东', '甘肃', '焦作']
    
    for location in location_prefixes:
        if candidate.startswith(location):
            # 检查上下文中是否有完整的公司名称
            pattern = rf'{candidate}[^\s，。；;、]*公司'
            if re.search(pattern, content):
                return True
    
    return False

def is_valid_name(name, surnames, content):
    """
    检查是否是有效的姓名
    修改：支持多字姓氏，名字部分为1-2个字
    新增：排除包含"被"字的姓名，更好地处理"董事"、"董事长"等称呼
    """
    if not name:
        return False
    
    # 检查是否只包含中文汉字
    if not re.match(r'^[\u4e00-\u9fa5]+$', name):
        return False
    
    # 检查长度：总长度2-4个字符（姓氏1-2字 + 名字1-2字）
    if len(name) < 2 or len(name) > 4:
        return False
    
    # 新增：姓名中不能包含"被"字
    if '被' in name:
        return False
    
    # 如果没有姓氏列表，无法验证
    if not surnames:
        return False
    
    # 检查是否以姓氏开头（支持多字姓氏）
    starts_with_surname = False
    for surname in surnames:
        if name.startswith(surname):
            starts_with_surname = True
            # 检查名字部分长度是否为1-2
            given_name_part = name[len(surname):]
            if len(given_name_part) < 1 or len(given_name_part) > 2:
                return False
            break
    
    if not starts_with_surname:
        return False
    
    # 过滤掉常见的非姓名词汇
    invalid_words = ['公司', '集团', '企业', '股份', '有限', '责任', '法院', '银行', '证券', '保险', 
                    '董事', '董事长', '总经理', '高管', '被', '长', '主席', '监事', '经理', '主任', 
                    '科长', '处长', '局长', '部长', '总裁', '总监', '委员', '代表']
    if any(word in name for word in invalid_words):
        return False
    
    # 特别检查：姓名不能包含"董事"等字（即使不在开头）
    if '董事' in name:
        return False
    
    # 过滤掉职位名称
    positions = ['董事', '董事长', '高管', '主席', '监事', '经理', '主任', '科长', '处长', '局长', '部长', '总裁', '总监']
    if name in positions:
        return False
    
    # 检查是否是常见地名开头的公司名截断
    location_prefixes = ['北京', '上海', '广州', '深圳', '天津', '重庆', '杭州', '南京', '武汉', '成都',
                        '西安', '郑州', '沈阳', '青岛', '大连', '宁波', '厦门', '苏州', '无锡', '常州',
                        '徐州', '南通', '扬州', '镇江', '泰州', '盐城', '淮安', '连云港', '宿迁', '温州',
                        '绍兴', '嘉兴', '湖州', '金华', '衢州', '台州', '丽水', '舟山', '福州', '泉州',
                        '漳州', '莆田', '宁德', '龙岩', '三明', '南平', '济南', '青岛', '烟台', '潍坊',
                        '临沂', '济宁', '淄博', '泰安', '枣庄', '日照', '莱芜', '聊城', '德州', '滨州',
                        '菏泽', '东营', '威海', '山东', '云南', '江苏', '江西', '安徽', '福建', '江阴',
                        '广西', '广东', '甘肃', '焦作']
    
    # 检查是否以地名开头
    for location in location_prefixes:
        if name.startswith(location):
            # 检查上下文中是否有完整的公司名称
            pattern = rf'{name}[^\s，。；;、]*公司'
            if re.search(pattern, content):
                return False
    
    # 检查姓名在全文中的出现次数
    name_count = content.count(name)
    if name_count < 1:
        return False
    
    # 额外检查：确保姓名不是从"公司董事"这样的短语中错误截取的
    if len(name) == 2 and name[0] == '司':
        pattern = rf'公司{name}'
        if re.search(pattern, content):
            return False
    
    # 新增检查：避免提取到"被"字开头的错误姓名
    if name.startswith('被'):
        return False
    
    # 新增检查：检查上下文，避免将"董事XXX"中的"董事"部分错误包含
    # 检查姓名前面是否有"董事"字样（排除空格）
    preceding_context = content[:content.find(name)][-20:]  # 取姓名前20个字符
    if '董事' in preceding_context and not re.search(r'公司董事|独立董事|执行董事|非执行董事', preceding_context[-5:]):
        # 如果前面有"董事"但不是完整的"公司董事"等复合词，可能是错误的
        return False
    
    return True

def extract_case_numbers(content):
    """
    提取案号信息
    案号通常出现在"案号："后，结构一般为"（年份）法院简称+案件类型+编号+号"
    修改：适配实际文件中的格式（半角括号和空格）
    """
    case_numbers = []
    
    # 案号正则表达式模式 - 修改为适配实际格式
    # 匹配格式如：(2023)京0101民初12345号 或 (2023)京0101执12345号
    # 注意：实际文件中是半角括号，并且可能有空格
    pattern = r"""
        案号[：:]\s*                # 案号关键词后跟冒号和可能的空格
        (                           # 开始捕获案号
        [\(（]                     # 左括号（半角或全角）
        \d{4}                       # 4位年份
        [\)）]                     # 右括号（半角或全角）
        \s*                         # 可能的空格
        [\u4e00-\u9fa5\d\s]+?      # 法院简称（汉字+数字+空格，非贪婪）
        (?:                         # 案件类型分组
            民[初终再审再]?\s*     # 民事案件：民初、民终、民再、民审等
            |刑[初终再]?\s*       # 刑事案件：刑初、刑终、刑再等
            |行[初终]?\s*         # 行政案件：行初、行终等
            |执(?:恢|异|保|行)?\s* # 执行案件：执、执恢、执异、执保、执行等
            |知民[初终]?\s*       # 知识产权民事
            |知刑[初终]?\s*       # 知识产权刑事
            |商[初终]?\s*         # 商事案件
        )
        \s*                         # 可能的空格
        \d+                         # 案件序号
        \s*                         # 可能的空格
        号                          # "号"字
        )                           # 结束捕获
    """
    
    # 编译正则表达式，忽略空格和注释
    case_regex = re.compile(pattern, re.VERBOSE | re.UNICODE)
    
    # 查找所有匹配
    matches = case_regex.findall(content)
    if matches:
        case_numbers.extend(matches)
    
    # 另一种可能的关键词 - 简化的模式，更灵活匹配
    pattern2 = r"""
        (?:案件案号|案由案号|案号)[：:]\s*
        (
        [\(（]
        \d{4}
        [\)）]
        [^，。；;、号]{5,30}?  # 非贪婪匹配，最多30个字符直到"号"
        号
        )
    """
    
    case_regex2 = re.compile(pattern2, re.VERBOSE | re.UNICODE)
    matches2 = case_regex2.findall(content)
    if matches2:
        case_numbers.extend(matches2)
    
    # 额外模式：简单匹配"案号："后面的括号格式
    pattern3 = r'案号[：:]\s*(\(?\d{4}\)?[^，。；;、]{5,30}?号)'
    matches3 = re.findall(pattern3, content)
    if matches3:
        case_numbers.extend(matches3)
    
    # 清理结果：去除多余空格
    cleaned_numbers = []
    for number in case_numbers:
        # 合并多个空格为一个空格
        cleaned = re.sub(r'\s+', ' ', number.strip())
        cleaned_numbers.append(cleaned)
    
    # 去重并返回
    return list(set(cleaned_numbers))

def extract_enforcement_numbers(content):
    """
    提取执行文号信息
    执行文号通常出现在"执行依据文号："等表达后，结构一般为"（年份）法院简称+执行类型+序号"
    修改：适配实际文件中的格式
    """
    enforcement_numbers = []
    
    # 执行文号正则表达式模式 - 修改为适配实际格式
    # 匹配格式如：(2023)京0101执12345号
    # 注意：实际文件中可能包含审判案号作为执行依据文号
    pattern = r"""
        (?:执行依据文号|执行文号|执行案号|依据文号)[：:]\s*  # 多种可能的关键词
        (                                           # 开始捕获执行文号
        [\(（]                                     # 左括号（半角或全角）
        \d{4}                                       # 4位年份
        [\)）]                                     # 右括号（半角或全角）
        \s*                                         # 可能的空格
        [\u4e00-\u9fa5\d\s]+?                      # 法院简称（汉字+数字+空格，非贪婪）
        (?:                                         # 案件类型分组
            执(?:恢|异|保|行)?\s*                 # 执行案件：执、执恢、执异、执保、执行等
            |民[初终再审再]?\s*                   # 民事案件（执行依据可能是审判案号）
            |刑[初终再]?\s*                       # 刑事案件
            |行[初终]?\s*                         # 行政案件
        )
        \s*                                         # 可能的空格
        \d+                                         # 序号
        \s*                                         # 可能的空格
        号                                          # "号"字
        )                                           # 结束捕获
    """
    
    # 编译正则表达式，忽略空格和注释
    enforcement_regex = re.compile(pattern, re.VERBOSE | re.UNICODE)
    
    # 查找所有匹配
    matches = enforcement_regex.findall(content)
    if matches:
        enforcement_numbers.extend(matches)
    
    # 另一种模式：更灵活的匹配，匹配"执行依据文号："后面的括号格式
    pattern2 = r"""
        (?:执行依据文号|执行文号|执行案号)[：:]\s*
        (
        [\(（]
        \d{4}
        [\)）]
        [^，。；;、号]{5,30}?  # 非贪婪匹配，最多30个字符直到"号"
        号
        )
    """
    
    enforcement_regex2 = re.compile(pattern2, re.VERBOSE | re.UNICODE)
    matches2 = enforcement_regex2.findall(content)
    if matches2:
        enforcement_numbers.extend(matches2)
    
    # 额外模式：简单匹配执行文号
    pattern3 = r'(?:执行依据文号|执行文号)[：:]\s*(\(?\d{4}\)?[^，。；;、]{5,30}?号)'
    matches3 = re.findall(pattern3, content)
    if matches3:
        enforcement_numbers.extend(matches3)
    
    # 清理结果：去除多余空格
    cleaned_numbers = []
    for number in enforcement_numbers:
        # 合并多个空格为一个空格
        cleaned = re.sub(r'\s+', ' ', number.strip())
        cleaned_numbers.append(cleaned)
    
    # 去重并返回
    return list(set(cleaned_numbers))

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
        '失信被执行人姓名': '',
        '案号': '',
        '执行文号': ''
    }
    
    # 第一步：提取失信被执行人姓名
    names = extract_names_with_multiple_methods(content, surnames)
    
    # 无论是否提取到姓名，都提取证券代码、证券简称、公告编号、公告时间
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
    
    # 如果提取到姓名，则继续提取案号和执行文号信息
    if names:
        unique_names = list(set(names))
        unique_names.sort()
        result['失信被执行人姓名'] = '，'.join(unique_names)
        
        # 提取案号信息
        case_numbers = extract_case_numbers(content)
        if case_numbers:
            result['案号'] = '，'.join(case_numbers)
        
        # 提取执行文号信息
        enforcement_numbers = extract_enforcement_numbers(content)
        if enforcement_numbers:
            result['执行文号'] = '，'.join(enforcement_numbers)
    
    return result

def extract_all_files(folder_path, surname_file_path):
    """
    提取文件夹中所有txt文件的信息
    """
    all_results = []
    
    surnames = load_surnames(surname_file_path)
    
    if not surnames:
        print(f"警告：未能加载姓氏表，姓名提取功能将受限！")
    
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
                '失信被执行人姓名': result['失信被执行人姓名'],
                '案号': result['案号'],
                '执行文号': result['执行文号']
            }
            all_results.append(record)
            
            # 显示提取结果
            if result['失信被执行人姓名']:
                print(f"  提取到姓名: {result['失信被执行人姓名']}")
                if result['案号']:
                    print(f"  提取到案号: {result['案号']}")
                if result['执行文号']:
                    print(f"  提取到执行文号: {result['执行文号']}")
            else:
                print(f"  未提取到姓名")
                print(f"  提取到证券代码: {result['证券代码']}")
                print(f"  提取到证券简称: {result['证券简称']}")
                print(f"  提取到公告编号: {result['公告编号']}")
                print(f"  提取到公告时间: {result['公告时间']}")
                
        except Exception as e:
            print(f"处理文件 {filename} 时出错: {e}")
            all_results.append({
                '文件名': filename,
                '证券代码': '',
                '证券简称': '',
                '公告编号': '',
                '公告时间': '',
                '失信被执行人姓名': '',
                '案号': '',
                '执行文号': '',
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
    
    # 确定字段名（排除错误字段）
    fieldnames = []
    for result in results:
        fieldnames = list(result.keys())
        # 移除'错误'字段
        if '错误' in fieldnames:
            fieldnames.remove('错误')
        break
    
    try:
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for result in results:
                # 移除'错误'字段
                row_data = {k: v for k, v in result.items() if k != '错误'}
                writer.writerow(row_data)
        
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
        output_file = os.path.join(folder_path, '失信人信息提取结果_完整版.csv')
        save_to_csv(results, output_file)
        
        print(f"\n处理完成!")
        print(f"共处理 {len(results)} 条记录")
        
        extracted_count = sum(1 for r in results if r.get('失信被执行人姓名'))
        print(f"成功提取失信被执行人姓名的记录: {extracted_count}")
        
        case_count = sum(1 for r in results if r.get('案号'))
        print(f"成功提取案号的记录: {case_count}")
        
        enforcement_count = sum(1 for r in results if r.get('执行文号'))
        print(f"成功提取执行文号的记录: {enforcement_count}")
        
        # 显示所有提取到的姓名
        if extracted_count > 0:
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
