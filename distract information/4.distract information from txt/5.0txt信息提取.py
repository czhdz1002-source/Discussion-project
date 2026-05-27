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

def extract_names_with_multiple_methods(content, surnames):
    """
    同时使用多种方法提取失信被执行人姓名
    """
    # 首先标准化内容，处理换行符问题
    normalized_content = normalize_content(content)
    
    names_set = set()
    
    # 方法1: 查找"失信被执行人（姓名/名称）："模式 - 现在会过滤公司名称
    method1_names = extract_names_method1(normalized_content, surnames)
    names_set.update(method1_names)
    
    # 方法2: 查找"被纳入失信被执行人名单"前的姓名
    method2_names = extract_names_method2(normalized_content, surnames)
    names_set.update(method2_names)
    
    # 方法3: 查找"失信被执行人名称："模式
    method3_names = extract_names_method3(normalized_content, surnames)
    names_set.update(method3_names)
    
    # 方法4: 查找表格格式的失信人信息
    method4_names = extract_names_method4(normalized_content, surnames)
    names_set.update(method4_names)
    
    # 方法5: 专门处理有换行情况的关键标识
    method5_names = extract_names_with_line_breaks(content, surnames)
    names_set.update(method5_names)
    
    # 验证所有候选姓名
    validated_names = []
    for name in names_set:
        if is_valid_name(name, surnames, normalized_content):
            validated_names.append(name)
    
    # 如果没有从方法1提取到姓名，但其他方法有提取，使用其他方法的结果
    if not method1_names and validated_names:
        # 方法1可能过滤掉了正确的姓名，检查是否有其他方法提取到的姓名
        return validated_names
    
    return validated_names

def extract_names_method1(content, surnames):
    """
    方法1: 查找"失信被执行人（姓名/名称）："模式
    改进：识别并过滤公司名称
    """
    names = []
    
    # 首先，匹配整行内容，以便分析后续文字
    patterns = [
        r'失信被执行人（姓名/名称）[：:]\s*([^\r\n]+)',  # 匹配整行
        r'失信被执行人\(姓名/名称\)[：:]\s*([^\r\n]+)',  # 匹配整行
    ]
    
    for pattern in patterns:
        matches = re.finditer(pattern, content)
        for match in matches:
            candidate = match.group(1).strip()
            
            # 判断是否是公司名称 - 多种检查方法
            
            # 1. 检查是否包含明显的公司关键词
            company_keywords = ['公司', '股份', '有限', '集团', '企业', '责任公司', '有限公司', '有限责任公司']
            if any(keyword in candidate for keyword in company_keywords):
                continue  # 跳过公司名称
            
            # 2. 检查是否是较长字符串（超过6个字符，很可能是公司名称或公司名截断）
            if len(candidate) > 6:
                continue  # 跳过长字符串
            
            # 3. 检查是否以常见地名开头（如"聊城"）
            # 常见的地名前缀
            location_prefixes = ['北京', '上海', '广州', '深圳', '天津', '重庆', '杭州', '南京', '武汉', '成都',
                                '西安', '郑州', '沈阳', '青岛', '大连', '宁波', '厦门', '苏州', '无锡', '常州',
                                '徐州', '南通', '扬州', '镇江', '泰州', '盐城', '淮安', '连云港', '宿迁', '温州',
                                '绍兴', '嘉兴', '湖州', '金华', '衢州', '台州', '丽水', '舟山', '福州', '泉州',
                                '漳州', '莆田', '宁德', '龙岩', '三明', '南平', '济南', '青岛', '烟台', '潍坊',
                                '临沂', '济宁', '淄博', '泰安', '枣庄', '日照', '莱芜', '聊城', '德州', '滨州',
                                '菏泽', '东营', '威海', '合肥', '芜湖', '马鞍山', '安庆', '黄山', '阜阳', '蚌埠',
                                '淮南', '淮北', '铜陵', '六安', '滁州', '池州', '宣城', '亳州', '南昌', '九江',
                                '赣州', '吉安', '宜春', '抚州', '上饶', '景德镇', '萍乡', '新余', '鹰潭', '长沙',
                                '株洲', '湘潭', '衡阳', '邵阳', '岳阳', '常德', '张家界', '益阳', '郴州', '永州',
                                '怀化', '娄底', '湘西', '石家庄', '唐山', '秦皇岛', '邯郸', '邢台', '保定', '张家口',
                                '承德', '沧州', '廊坊', '衡水', '太原', '大同', '阳泉', '长治', '晋城', '朔州',
                                '晋中', '运城', '忻州', '临汾', '吕梁', '呼和浩特', '包头', '乌海', '赤峰', '通辽',
                                '鄂尔多斯', '呼伦贝尔', '巴彦淖尔', '乌兰察布', '兴安', '锡林郭勒', '阿拉善', '长春',
                                '吉林', '四平', '辽源', '通化', '白山', '松原', '白城', '延边', '哈尔滨', '齐齐哈尔',
                                '鸡西', '鹤岗', '双鸭山', '大庆', '伊春', '佳木斯', '七台河', '牡丹江', '黑河', '绥化',
                                '大兴安岭']
            
            # 检查是否以地名开头且后面跟着非姓名常用字
            for location in location_prefixes:
                if candidate.startswith(location):
                    # 如果以地名开头，很可能是公司名称
                    continue
            
            # 4. 检查是否是2-3个中文字符，并且以姓氏开头
            if re.match(r'^[\u4e00-\u9fa5]{2,3}$', candidate):
                if is_valid_name(candidate, surnames, content):
                    names.append(candidate)
    
    # 原有的模式也保留，用于其他情况
    patterns2 = [
        r'失信被执行人（姓名/名称）[：:]\s*([^\s，。；;、]+)',
        r'失信被执行人\(姓名/名称\)[：:]\s*([^\s，。；;、]+)',
        r'失信被执行人（姓名/名称）[：:]\s*([^，。；;、]{2,3})',
        r'失信被执行人\(姓名/名称\)[：:]\s*([^，。；;、]{2,3})',
    ]
    
    for pattern in patterns2:
        matches = re.finditer(pattern, content)
        for match in matches:
            name = match.group(1).strip()
            # 添加公司名称检查
            if len(name) > 3:  # 如果长度超过3，可能是公司名称
                continue
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
    # 修改模式：确保姓名不是"董事"或包含"董事"二字
    patterns = [
        # 标准模式，用于处理已标准化的内容
        r'公司董事\s+([\u4e00-\u9fa5]{2,3})\s+被纳入失信被执行人名单',  # 修改：增加\s+确保完整匹配
        r'董事\s+([\u4e00-\u9fa5]{2,3})\s+被纳入失信被执行人名单',  # 修改：增加\s+确保完整匹配
        r'([\u4e00-\u9fa5]{2,3})\s+被纳入失信被执行人名单',  # 修改：2-3个字符，增加\s+
        r'公司董事([\u4e00-\u9fa5]{2,3})被纳入失信被执行人名单',  # 新增：无空格的匹配
        r'董事([\u4e00-\u9fa5]{2,3})被纳入失信被执行人名单',  # 新增：无空格的匹配
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
        r'失信被执行人名称[：:]\s*([^，。；;、]{2,3})',  # 修改：2-3个字符
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
    
    pattern = r'失信被执行人（姓名/名称）[：:]\s*([^，。；;、\s]{2,3})'  # 修改：2-3个字符
    matches = re.finditer(pattern, content)
    for match in matches:
        name = match.group(1).strip()
        if is_valid_name(name, surnames, content):
            names.append(name)
    
    pattern2 = r'失信被执行人名称[：:]\s*([^，。；;、\s]{2,3})'  # 修改：2-3个字符
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
    # 使用更灵活的正则表达式，允许任意空白字符（包括换行）
    # 修改模式：避免匹配到"司董事"这样的错误结果
    patterns = [
        # 模式1: 董事 + 姓名 + 被纳入失信被执行人名单（允许任意位置换行）
        r'董事[\s，、]*([\u4e00-\u9fa5]{2,3})[\s，、]*被[\s，、]*纳[\s，、]*入[\s，、]*失[\s，、]*信[\s，、]*被[\s，、]*执[\s，、]*行[\s，、]*人[\s，、]*名[\s，、]*单',  # 修改：添加[，、]分隔符
        r'公司董事[\s，、]*([\u4e00-\u9fa5]{2,3})[\s，、]*被[\s，、]*纳[\s，、]*入[\s，、]*失[\s，、]*信[\s，、]*被[\s，、]*执[\s，、]*行[\s，、]*人[\s，、]*名[\s，、]*单',  # 修改：添加[，、]分隔符
        
        # 新增模式：避免截取"公司董事"中的"司董事"
        r'(?:^|[^\u4e00-\u9fa5])董事[\s，、]*([\u4e00-\u9fa5]{2,3})(?![司])[\s，、]*被纳入失信被执行人名单',  # 新增：使用否定前瞻确保不匹配"司董事"
        r'(?:^|[^\u4e00-\u9fa5])公司董事[\s，、]*([\u4e00-\u9fa5]{2,3})[\s，、]*被纳入失信被执行人名单',  # 新增：确保"公司董事"是完整的
        
        # 模式2: 简化的模式，匹配"被纳入失信"和"被执行人名单"之间有换行的情况
        r'(?:^|[^\u4e00-\u9fa5])([\u4e00-\u9fa5]{2,3})(?![司])[\s，、]*被[\s，、]*纳[\s，、]*入[\s，、]*失[\s，、]*信[\s，、]*被[\s，、]*执[\s，、]*行[\s，、]*人[\s，、]*名[\s，、]*单',
    ]
    
    for pattern in patterns:
        matches = re.finditer(pattern, original_content, re.DOTALL)
        for match in matches:
            # 获取捕获组中的姓名
            if match.lastindex:
                name = match.group(1).strip()
                # 额外检查：确保姓名不以"司"开头（避免"司董事"）
                if name.startswith('司'):
                    continue
                
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
    
    # 修改：名字部分改为1-2个字符（姓氏1个字符 + 名字1-2个字符 = 总共2-3个字符）
    name_pattern = rf'({surname_regex}[\u4e00-\u9fa5]{{1,2}})'  # 修改：名字部分1-2个字符
    
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
    # 修改：将长度条件从2-4改为2-3
    if not name or len(name) < 2 or len(name) > 3:  # 修改：2-3个字符
        return False
    
    # 检查是否以姓氏开头
    starts_with_surname = any(name.startswith(surname) for surname in surnames)
    if not starts_with_surname:
        return False
    
    # 只包含中文汉字
    if not re.match(r'^[\u4e00-\u9fa5]+$', name):
        return False
    
    # 过滤掉常见的非姓名词汇
    invalid_words = ['公司', '集团', '企业', '股份', '有限', '责任', '法院', '银行', '证券', '保险', '董事']
    if any(word in name for word in invalid_words):
        return False
    
    # 特别检查：姓名不能包含"董事"二字（即使不在开头）
    if '董事' in name:
        return False
    
    # 过滤掉职位名称
    positions = ['董事', '主席', '监事', '经理', '主任', '科长', '处长', '局长', '部长', '总裁', '总监']
    if name in positions:
        return False
    
    # 新增：检查是否是常见地名开头的公司名截断（如"聊城继"）
    # 常见的地名前缀
    location_prefixes = ['北京', '上海', '广州', '深圳', '天津', '重庆', '杭州', '南京', '武汉', '成都',
                        '西安', '郑州', '沈阳', '青岛', '大连', '宁波', '厦门', '苏州', '无锡', '常州',
                        '徐州', '南通', '扬州', '镇江', '泰州', '盐城', '淮安', '连云港', '宿迁', '温州',
                        '绍兴', '嘉兴', '湖州', '金华', '衢州', '台州', '丽水', '舟山', '福州', '泉州',
                        '漳州', '莆田', '宁德', '龙岩', '三明', '南平', '济南', '青岛', '烟台', '潍坊',
                        '临沂', '济宁', '淄博', '泰安', '枣庄', '日照', '莱芜', '聊城', '德州', '滨州',
                        '菏泽', '东营', '威海']
    
    # 检查是否以地名开头
    for location in location_prefixes:
        if name.startswith(location):
            # 检查上下文中是否有完整的公司名称
            pattern = rf'{name}[^\s，。；;、]*公司'
            if re.search(pattern, content):
                return False  # 很可能是公司名称的截断
    
    # 检查姓名在全文中的出现次数
    name_count = content.count(name)
    if name_count < 1:
        return False
    
    # 额外检查：确保姓名不是从"公司董事"这样的短语中错误截取的
    # 如果姓名是"司"加上其他字，可能是从"公司董事"截取的
    if len(name) == 2 and name[0] == '司':
        # 检查上下文中是否有"公司董事"或类似模式
        pattern = rf'公司{name}'
        if re.search(pattern, content):
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
