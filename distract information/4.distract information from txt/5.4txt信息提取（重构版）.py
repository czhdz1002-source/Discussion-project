import os
import re
import docx
import csv

# ── jieba 分词（可选依赖，未安装时自动降级） ──────────────────────────────
try:
    import jieba
    import jieba.posseg as pseg
    JIEBA_AVAILABLE = True
except ImportError:
    JIEBA_AVAILABLE = False
    print("提示：未检测到 jieba 库，姓名分词优化功能将关闭。"
          "可通过 pip install jieba 安装以获得更精准的姓名提取。")

# ── 职位词表：加入 jieba 用户词典，防止"董事长武力"→"长武力" ───────────────
POSITION_WORDS = [
    '董事长', '董事会', '执行董事', '独立董事', '非执行董事', '董事',
    '监事长', '监事会', '监事',
    '总经理', '副总经理', '总裁', '副总裁', '总监', '副总监',
    '首席执行官', '首席财务官', '首席运营官',
    '法定代表人', '实际控制人', '控股股东',
    '主席', '副主席', '委员', '秘书',
]

def _init_jieba(surnames):
    """初始化 jieba：注入职位词和姓氏，确保分词边界正确"""
    if not JIEBA_AVAILABLE:
        return
    for word in POSITION_WORDS:
        jieba.add_word(word, freq=50000, tag='n')
    # 将姓氏作为单字词加入，辅助姓名边界识别
    for s in surnames:
        jieba.add_word(s, freq=10000, tag='nr')
    jieba.initialize()


# ── 同义词/近义说法扩展表 ────────────────────────────────────────────────────
# 结构：{ 规范化表达: [所有等价说法列表] }
# 用于在正则模式构建时自动扩展，覆盖公告中各种不统一的措辞
SYNONYM_GROUPS = {
    # "被纳入失信被执行人名单" 的各种说法
    'inducted': [
        '被纳入失信被执行人名单',
        '纳入失信被执行人名单',
        '被纳入失信被执行名单',
        '纳入失信被执行名单',
        '被列入失信被执行人名单',
        '列入失信被执行人名单',
        '被列入失信被执行名单',
        '列入失信被执行名单',
        '被认定为失信被执行人',
        '认定为失信被执行人',
        '被纳入失信人名单',
        '纳入失信人名单',
        '被列为失信被执行人',
        '列为失信被执行人',
        '被加入失信被执行人名单',
        '加入失信被执行人名单',
        '被登记为失信被执行人',
        '登记为失信被执行人',
        '被录入失信被执行人名单',
        '录入失信被执行人名单',
        '被纳入失信名单',
        '纳入失信名单',
        '列入失信名单',
        '被列入失信名单',
    ],
    # "失信被执行人（姓名/名称）：" 的标题字段说法
    'field_name': [
        '失信被执行人（姓名/名称）',
        '失信被执行人(姓名/名称)',
        '失信被执行人姓名/名称',
        '失信被执行人姓名',
        '失信被执行人名称',
        '失信人姓名',
        '失信人名称',
        '被执行人姓名',
        '被执行人名称',
    ],
}


def build_synonym_pattern(group_key, allow_spaces=False):
    """
    根据同义词组生成单个正则表达式，支持各词之间允许任意空白/换行。
    allow_spaces=True 时每个汉字间插入 \\s*，用于处理换行散字情况。
    """
    synonyms = SYNONYM_GROUPS.get(group_key, [])
    parts = []
    for phrase in synonyms:
        if allow_spaces:
            escaped = r'\s*'.join(re.escape(ch) for ch in phrase)
        else:
            escaped = re.escape(phrase)
        parts.append(escaped)
    # 按长度降序，确保长模式优先匹配
    parts.sort(key=len, reverse=True)
    return '(?:' + '|'.join(parts) + ')'

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

def extract_name_by_jieba(context_text, surnames):
    """
    使用 jieba 对"职位+姓名"的上下文片段进行分词，提取紧跟职位词后的人名。
    例如："董事长武力" → ['董事长', '武力'] → 返回 '武力'
    如果 jieba 不可用则返回空列表。
    """
    if not JIEBA_AVAILABLE:
        return []

    names = []
    words = list(jieba.cut(context_text))
    for i, word in enumerate(words):
        if word in POSITION_WORDS:
            # 取紧跟职位词之后的词
            if i + 1 < len(words):
                candidate = words[i + 1].strip()
                if is_valid_name(candidate, surnames, context_text):
                    names.append(candidate)
            # 也检查后两个词拼合（应对分词粒度问题）
            if i + 2 < len(words):
                candidate2 = (words[i + 1] + words[i + 2]).strip()
                if is_valid_name(candidate2, surnames, context_text):
                    names.append(candidate2)
    return names


def extract_names_with_multiple_methods(content, surnames):
    """
    同时使用多种方法提取失信被执行人姓名。
    方法1-6：原有正则方法。
    方法7：基于 jieba 分词修正职位+姓名粘连问题（如"董事长武力"→"武力"）。
    方法8：基于同义词表的通用扩展正则，覆盖未在方法1-6中枚举的同义说法。
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

    # ── 方法7：jieba 分词修正职位粘连问题 ──────────────────────────────────
    method7_names = extract_names_method7_jieba(normalized_content, surnames)
    names_set.update(method7_names)

    # ── 方法8：同义词通用扩展正则 ─────────────────────────────────────────
    method8_names = extract_names_method8_synonym(normalized_content, surnames)
    names_set.update(method8_names)

    # ── 方法9：处理姓名本身被换行拆断（如"刘中↵庆"→"刘中庆"）─────────────
    # 方法9已在正则中确认了失信上下文，直接加入，跳过后面的 is_valid_name 全文存在性检查
    method9_names = extract_names_method9_linebreak_name(content, surnames)
    # 单独收集，避免被normalized_content里找不到该姓名而过滤掉
    method9_validated = [n for n in method9_names if not is_only_xiao_fei_person(n, normalized_content)]

    # 验证所有候选姓名
    validated_names = []
    for name in names_set:
        if is_valid_name(name, surnames, normalized_content):
            # 额外：语义上下文过滤——排除仅出现在"限制消费"语境中的姓名
            if not is_only_xiao_fei_person(name, normalized_content):
                validated_names.append(name)

    # 合并方法9结果（去重）
    all_names = list(dict.fromkeys(validated_names + method9_validated))
    return all_names


def is_only_xiao_fei_person(name, content):
    """
    判断某姓名是否只出现在"限制消费"相关语境，而非失信被执行人语境。
    若是，则该姓名不应被提取为失信被执行人。

    判断逻辑：
    - 在文中找出所有包含该姓名的上下文窗口（前后各40字）
    - 若每一处都含有"限制消费/限消/消费令"等关键词，
      且没有任何一处含有"失信被执行/纳入名单"等关键词，
      则判断为"仅限消人员"，返回 True（应被过滤）
    """
    # 失信相关关键词
    shixin_keywords = ['失信被执行', '纳入失信', '列入失信', '失信名单', '失信人名单']
    # 限制消费相关关键词
    xiaofei_keywords = ['限制消费', '限消', '消费令', '限消人员']

    found_shixin = False
    # 注：found_only_xiaofei 已移除（原为死变量，从未被赋值或读取）

    start = 0
    occurrence_count = 0
    xiaofei_occurrence_count = 0

    while True:
        idx = content.find(name, start)
        if idx == -1:
            break
        occurrence_count += 1
        window = content[max(0, idx - 40): idx + len(name) + 40]

        has_shixin = any(kw in window for kw in shixin_keywords)
        has_xiaofei = any(kw in window for kw in xiaofei_keywords)

        if has_shixin:
            found_shixin = True
        if has_xiaofei and not has_shixin:
            xiaofei_occurrence_count += 1

        start = idx + 1

    # 如果出现过"失信"相关上下文，不过滤
    if found_shixin:
        return False

    # 如果所有出现都在"限制消费"语境中，过滤
    if occurrence_count > 0 and xiaofei_occurrence_count == occurrence_count:
        return True

    return False

def extract_names_method1(content, surnames):
    """
    方法1: 查找"失信被执行人（姓名/名称）："等模式
    修改：支持"失信被执行人姓名/名称："等多种表达
    """
    names = []
    
    # 多种表达模式
    patterns = [
        # 标准模式（带"失信"前缀）
        r'失信被执行人（姓名/名称）[：:]\s*([^\r\n]+)',
        r'失信被执行人\(姓名/名称\)[：:]\s*([^\r\n]+)',
        r'失信被执行人姓名/名称[：:]\s*([^\r\n]+)',
        r'失信被执行人姓名[：:]\s*([^\r\n]+)',
        # 新增：券商公告等格式，字段为"被执行人姓名/名称"（无"失信"前缀）
        r'被执行人姓名/名称[：:]\s*([^\r\n]+)',
        r'被执行人（姓名/名称）[：:]\s*([^\r\n]+)',
        r'被执行人\(姓名/名称\)[：:]\s*([^\r\n]+)',
        r'被执行人姓名[：:]\s*([^\r\n]+)',
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
    
    # 原有模式也保留（精确匹配，不含后续文本）
    patterns2 = [
        r'失信被执行人（姓名/名称）[：:]\s*([^\s，。；;、]+)',
        r'失信被执行人\(姓名/名称\)[：:]\s*([^\s，。；;、]+)',
        r'失信被执行人姓名/名称[：:]\s*([^\s，。；;、]+)',
        r'失信被执行人姓名[：:]\s*([^\s，。；;、]+)',
        r'被执行人姓名/名称[：:]\s*([^\s，。；;、]+)',
        r'被执行人（姓名/名称）[：:]\s*([^\s，。；;、]+)',
        r'被执行人姓名[：:]\s*([^\s，。；;、]+)',
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
    修复：允许"被"与"纳入"之间有空格（normalize后换行变空格的情况）
    修复(Bug1)：将 induct_suffix 拆为"含被"/"不含被"两个分支，
              彻底解决 {2,4} 贪婪把"被"吃进姓名的问题。
              原来用负向前瞻 (?![被纳列]) 的方案会同时阻止
              "李栋被纳入"中正确匹配"李栋"（因为"李栋"后一字恰好是"被"），
              两分支方案则让正则引擎自然回溯，无此副作用。
    新增：支持"及董事XXX"前缀
    """
    names = []

    # 两个后缀分支：明确区分"被"是否存在，避免贪婪吃掉"被"字
    induct_with_bei    = r'被\s*(?:纳入|列入)\s*失信被执行人?名单'   # 姓名后跟"被纳/被列"
    induct_without_bei = r'(?:纳入|列入)\s*失信被执行人?名单'         # 姓名后直接跟"纳/列"

    name_plain = r'([\u4e00-\u9fa5]{2,4})'
    pos_prefix = r'(?:(?:公司\s*)?(?:及\s*)?(?:董事长?|总经理|监事长?|法定代表人|实际控制人)\s*)?'

    patterns = []
    for suffix in [induct_with_bei, induct_without_bei]:
        patterns.append(pos_prefix + name_plain + r'\s*' + suffix)

    # 原有严格模式（保留兜底，这些模式中"被"单独在外，不存在贪婪问题）
    patterns += [
        r'公司董事\s+([\u4e00-\u9fa5]{2,4})\s+被纳入失信被执行人名单',
        r'董事\s+([\u4e00-\u9fa5]{2,4})\s+被纳入失信被执行人名单',
        r'公司董事\s+([\u4e00-\u9fa5]{2,4})\s+纳入失信被执行人名单',
        r'董事\s+([\u4e00-\u9fa5]{2,4})\s+纳入失信被执行人名单',
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, content):
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

    # 通用后缀：允许任意换行/空格散布在字符间
    def _spaced(phrase):
        """将短语中每个字符间插入 \\s*，适配跨行散字"""
        return r'\s*'.join(re.escape(c) for c in phrase)

    induct_suffixes = [
        _spaced('被纳入失信被执行人名单'),
        _spaced('纳入失信被执行人名单'),
        _spaced('被列入失信被执行人名单'),
        _spaced('列入失信被执行人名单'),
        _spaced('被纳入失信被执行名单'),
        _spaced('纳入失信被执行名单'),
        _spaced('被列入失信被执行名单'),
        _spaced('列入失信被执行名单'),
    ]
    name_re = r'([\u4e00-\u9fa5]{2,4})'

    # 前缀类型
    position_prefixes = [
        r'(?:公司\s*)?(?:及\s*)?(?:董事长?|监事长?|总经理|法定代表人|实际控制人)\s*',
        r'(?:公司\s*)?(?:及\s*)?董事\s*',
        r'',  # 无前缀
    ]

    patterns = []
    for prefix in position_prefixes:
        for suffix in induct_suffixes:
            patterns.append(prefix + name_re + r'\s*' + suffix)
    
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


def is_valid_name_loose(name, surnames):
    """
    宽松版姓名验证：仅做基本检查（无需在全文中查找上下文），
    专供 extract_names_method9_linebreak_name 使用，
    因为该方法已在正则匹配时确认了失信相关上下文。
    """
    if not name or not re.match(r'^[\u4e00-\u9fa5]+$', name):
        return False
    if len(name) < 2 or len(name) > 4:
        return False
    if '被' in name:
        return False

    invalid_words = ['公司', '集团', '企业', '股份', '有限', '责任', '法院', '银行', '证券', '保险',
                     '董事', '总经理', '高管', '主席', '监事', '经理', '主任',
                     '科长', '处长', '局长', '部长', '总裁', '总监', '委员', '代表']
    if any(w in name for w in invalid_words):
        return False

    if not surnames:
        return False

    invalid_given_chars = set('已及而且并或则虽然但是因为所以如果仍再也都很更加非常较比较何其某些各该此这那其每')
    for surname in surnames:
        if name.startswith(surname):
            given = name[len(surname):]
            if 1 <= len(given) <= 2 and not any(c in invalid_given_chars for c in given):
                return True
    return False


def extract_names_method9_linebreak_name(original_content, surnames):
    """
    方法9：专门处理姓名本身被换行拆断的情况。
    例如："公司董事刘中↵庆被纳入" → normalize后变"刘中 庆"，直接匹配失败。
    本方法在原始内容中，允许姓名各字之间出现任意空白/换行，
    匹配后去除空格还原纯汉字姓名再进行验证。
    """
    names = []

    # 允许被纳入/列入等关键词散字（各字间可有换行）
    induct_suffix_loose = r'(?:被\s*)?(?:纳\s*入|列\s*入)\s*失\s*信\s*被\s*执\s*行\s*人?\s*名\s*单'

    # 允许姓名2-4个汉字之间夹杂空白
    # 修复：每个后续字用负向前瞻排除"被/纳/列"，防止吃掉induct_suffix开头的字
    # 末字后加正向前瞻，确保紧跟着的是induct_suffix的起点
    name_with_spaces = r'([\u4e00-\u9fa5](?:\s*(?![被纳列])[\u4e00-\u9fa5]){1,3})(?=\s*(?:被\s*)?(?:纳|列))'

    position_prefix_loose = r'(?:(?:公司\s*)?(?:控股股东\s*[、，,]?\s*)?(?:及\s*)?(?:董事长?|监事长?|总经理|法定代表人|实际控制人|董事|监事)\s*)'

    patterns = [
        position_prefix_loose + name_with_spaces + r'\s*' + induct_suffix_loose,
        name_with_spaces + r'\s*' + induct_suffix_loose,
    ]

    for pattern in patterns:
        for match in re.finditer(pattern, original_content, re.DOTALL):
            raw_name = match.group(1)
            # 去掉姓名中间的空白，还原为纯汉字
            clean_name = re.sub(r'\s+', '', raw_name)
            if clean_name.startswith('司'):
                continue
            if is_valid_name_loose(clean_name, surnames):
                names.append(clean_name)

    return names


def extract_names_method7_jieba(content, surnames):
    """
    方法7：使用 jieba 分词修正职位词与姓名粘连问题。
    针对"董事长武力"/"总经理张三"/"法定代表人李四"等情况，
    利用 jieba 的分词边界正确切割出姓名部分。
    如果 jieba 未安装则跳过，不影响其他方法运行。
    """
    if not JIEBA_AVAILABLE:
        return []

    names = []
    # 提取包含职位词的上下文窗口（前后各40字），避免对整篇文本分词效率低
    for pos_word in POSITION_WORDS:
        start = 0
        while True:
            idx = content.find(pos_word, start)
            if idx == -1:
                break
            # 取窗口
            window_start = max(0, idx - 10)
            window_end = min(len(content), idx + len(pos_word) + 15)
            window = content[window_start:window_end]

            # jieba 分词
            seg_list = list(jieba.cut(window))
            for i, seg in enumerate(seg_list):
                if seg == pos_word or seg in POSITION_WORDS:
                    # 检查紧随其后的词（单词或双词）
                    for look_ahead in range(1, 3):
                        if i + look_ahead < len(seg_list):
                            candidate = ''.join(seg_list[i + 1: i + 1 + look_ahead]).strip()
                            if candidate and is_valid_name(candidate, surnames, content):
                                names.append(candidate)
            start = idx + 1

    return names


def extract_names_method8_synonym(content, surnames):
    """
    方法8：同义词扩展通用正则。
    自动将 SYNONYM_GROUPS['inducted'] 中所有说法合并为单一正则，
    无需逐一手动枚举，新增同义说法只需在 SYNONYM_GROUPS 中添加即可。
    匹配结构：[职位词?] + [姓名] + [同义词组中任意一种说法]
    """
    names = []

    # 构造"被纳入/被列入/..."的同义词合并正则，拆为两组
    # 含"被"开头的说法 和 不含"被"开头的说法
    synonyms_inducted = SYNONYM_GROUPS.get('inducted', [])
    inducted_with_bei    = [s for s in synonyms_inducted if s.startswith('被')]
    inducted_without_bei = [s for s in synonyms_inducted if not s.startswith('被')]

    def _build_alt(phrases):
        parts = sorted([re.escape(p) for p in phrases], key=len, reverse=True)
        return '(?:' + '|'.join(parts) + ')'

    pattern_with_bei    = _build_alt(inducted_with_bei)    # 以"被"开头的所有说法
    pattern_without_bei = _build_alt(inducted_without_bei) # 直接以"纳/列"等开头的所有说法

    # 同时构造字段名称类型的同义词合并正则（用于 "字段名：姓名" 格式）
    field_pattern = build_synonym_pattern('field_name')

    # --- 模式A：姓名 + 同义词组（覆盖所有"被纳入/被列入..."说法）---
    # 修复(Bug1)：拆为含被/不含被两个分支，彻底解决贪婪吃掉"被"字的问题。
    # 原来用 (?![被纳列\s]) 负向前瞻，副作用是阻止了"李栋被纳入"中正确匹配"李栋"
    # （因为"李栋"后一字"被"恰好在排除集合中）。
    # 两分支方案让引擎自然回溯，无此副作用。
    position_prefix = r'(?:(?:公司\s*)?(?:控股股东[、，,]?\s*)?(?:及\s*)?(?:董事长|执行董事|独立董事|非执行董事|监事长|总经理|副总经理|总裁|法定代表人|实际控制人|董事|监事|主席)\s*)?'
    name_plain = r'([\u4e00-\u9fa5]{2,4})'

    for suffix in [pattern_with_bei, pattern_without_bei]:
        pattern_a = position_prefix + name_plain + r'\s*' + suffix
        for match in re.finditer(pattern_a, content):
            candidate = match.group(1).strip()
            if not candidate.startswith('司') and is_valid_name(candidate, surnames, content):
                names.append(candidate)

    # --- 模式A2：多人连列格式（如"李栋及公司董事李鲁方被纳入"）---
    # 用姓氏锚定两个捕获组，防止匹配到职位末字
    # 同样用两分支：name1后跟"及..."，name2后跟inducted（含被/不含被均可）
    if surnames:
        sorted_s = sorted(surnames, key=len, reverse=True)
        sp = '|'.join(re.escape(s) for s in sorted_s)
        # name1：姓氏锚定，后接"及/和/与"连词（不存在贪婪问题，因为连词不是汉字）
        name1_anchored = rf'((?:{sp})[\u4e00-\u9fa5]{{1,2}})'
        name2_anchored = rf'((?:{sp})[\u4e00-\u9fa5]{{1,2}})'
        inducted_all = build_synonym_pattern('inducted')
        multi_person_pattern = (name1_anchored
            + r'\s*(?:及|和|与)\s*(?:公司\s*)?(?:[^\s，。]{0,10})?\s*'
            + name2_anchored + r'\s*' + inducted_all)
        for match in re.finditer(multi_person_pattern, content):
            for i in (1, 2):
                candidate = match.group(i).strip()
                if candidate and not candidate.startswith('司') and is_valid_name(candidate, surnames, content):
                    names.append(candidate)

    # --- 模式B：字段名称 + 冒号 + 姓名 ---
    pattern_b = field_pattern + r'[：:]\s*([\u4e00-\u9fa5]{2,4})'
    for match in re.finditer(pattern_b, content):
        candidate = match.group(1).strip()
        if is_valid_name(candidate, surnames, content):
            names.append(candidate)

    return names


def extract_names_from_text(text, surnames, full_content):
    """
    从文本中提取可能的姓名
    修改：支持多字姓氏，名字部分为1-2个字
    修复：若候选文本包含"限制消费"等非失信段落的标志词，截断到该词之前，
         避免将限制消费人员姓名误扫为失信被执行人。
    """
    names = []

    # 若长串跨越了"限制消费"等无关段落，只取其之前的部分
    cross_section_markers = ['限制消费', '被限制消费', '限消']
    for marker in cross_section_markers:
        marker_idx = text.find(marker)
        if marker_idx != -1:
            text = text[:marker_idx]
            break

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
    # 修复(Bug3)：将长度不合法时的 return False 改为 continue，
    # 防止某个较长姓氏先匹配但名字部分长度不合法时直接退出，
    # 导致后续更短的姓氏无法被尝试（如姓氏表含"李鲁"会阻止"李"匹配"李鲁方"）
    starts_with_surname = False
    matched_surname = None
    for surname in sorted(surnames, key=len, reverse=True):  # 长姓氏优先，但不短路
        if name.startswith(surname):
            given_name_part = name[len(surname):]
            if len(given_name_part) < 1 or len(given_name_part) > 2:
                continue  # 名字部分长度不合法，跳过此姓氏，继续尝试更短的
            starts_with_surname = True
            matched_surname = surname
            break

    if not starts_with_surname:
        return False

    # 新增：名字部分（去掉姓氏后的字符）不能包含常见虚词、动词、副词、介词
    # 这类字不会出现在人名的"名"部分，能过滤"王兵已"/"李栋及"等误识别
    invalid_given_chars = set('已及而且并或则虽然但是因为所以如果虽仍再也都很更加非常较比较何其某些各该此这那其每')
    if matched_surname:
        given_part = name[len(matched_surname):]
        if any(c in invalid_given_chars for c in given_part):
            return False
    
    # 过滤掉常见的非姓名词汇
    # 修复(Bug2)：去掉单字 '被' 和 '长'——
    #   '被' 已在前面专门检查（if '被' in name），重复多余
    #   '长' 会误杀 "张长明"/"王长海" 等含"长"字的合法人名
    # 修复(Bug4)：将下方三处重复的 '董事' 检查统一到此处，删除冗余代码
    invalid_words = ['公司', '集团', '企业', '股份', '有限', '责任', '法院', '银行', '证券', '保险',
                    '董事', '董事长', '总经理', '高管', '主席', '监事', '经理', '主任',
                    '科长', '处长', '局长', '部长', '总裁', '总监', '委员', '代表']
    if any(word in name for word in invalid_words):
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
    
    # 检查上下文：过滤"长武力"/"席张三"这类职位词末字误粘入姓名首字的情况。
    # 判断依据：姓名第一个字是职位词末尾字（如"长"、"席"），且文中该姓名的
    # 每一处出现，其紧前字+姓名首字恰好构成某个职位词的末两字（如"事长"、"事席"）。
    # 注意：像"董事朱佑兰"是正常写法，此时姓名首字"朱"不是职位词末字，不过滤。
    
    # 构建职位词末两字集合，用于精确判断是否粘连
    position_word_endings = set()
    for pw in POSITION_WORDS:
        if len(pw) >= 2:
            position_word_endings.add(pw[-2:])  # 如"事长"、"事席"、"理长"

    # 检查文中每一处姓名出现，只要有一处前一字不构成职位末两字，就认为合法
    search_start = 0
    found_valid_occurrence = False
    while True:
        idx = content.find(name, search_start)
        if idx == -1:
            break
        if idx > 0:
            prev_char = content[idx - 1]
            # 取前一字+姓名首字，看是否命中职位词末两字
            two_char = prev_char + name[0]
            if two_char in position_word_endings:
                # 这是粘连情况（如"长武"对应"董事长武力"），跳过
                search_start = idx + 1
                continue
        # 前面不是职位末两字，说明是合法出现（如"董事 朱佑兰"或句首）
        found_valid_occurrence = True
        break

    if not found_valid_occurrence:
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
    # 修复(Bug9)：原逻辑直接取最后10行中最靠近末尾的日期，
    # 但公告正文末尾可能有引用的法律文书日期，会被误取。
    # 修复策略：优先在"董事会"/"特此公告"等发文落款标志词之后的行中查找日期；
    # 若找不到，再退化到原来的最后10行兜底逻辑。
    date_patterns = [
        r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日',
        r'(\d{4})-(\d{1,2})-(\d{1,2})',
        r'(\d{4})/(\d{1,2})/(\d{1,2})',
        r'(\d{4})\.(\d{1,2})\.(\d{1,2})'
    ]

    def _try_extract_date(text_segment):
        """在给定文本段中查找最靠近末尾的日期，返回格式化字符串或空串"""
        lines_seg = text_segment.strip().split('\n')
        for line in reversed(lines_seg):
            for pat in date_patterns:
                m = re.search(pat, line)
                if m:
                    try:
                        y, mo, d = m.groups()
                        return f"{y}年{int(mo)}月{int(d)}日"
                    except Exception:
                        continue
        return ''

    # 阶段1：在"董事会"/"特此公告"/"此致"等落款标志词出现后的文本中查找
    sign_markers = ['董事会', '特此公告', '此致', '主办券商', '发布日期']
    announce_date = ''
    for marker in sign_markers:
        marker_idx = content.rfind(marker)   # 取最后一次出现（避免正文引用）
        if marker_idx != -1:
            tail_text = content[marker_idx:]
            announce_date = _try_extract_date(tail_text)
            if announce_date:
                break

    # 阶段2：若落款区域未找到，退化到原始最后10行兜底
    if not announce_date:
        lines = content.strip().split('\n')
        announce_date = _try_extract_date('\n'.join(lines[-10:]))

    result['公告时间'] = announce_date
    
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

    # ── 初始化 jieba 分词器（注入职位词和姓氏） ──
    _init_jieba(surnames)
    if JIEBA_AVAILABLE:
        print("jieba 分词已启用（方法7：职位粘连修正）")
    
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
    修复(Bug5)：DictWriter 添加 extrasaction='ignore'，防止出错记录含额外'错误'键时
              writerow 抛出 ValueError（原代码仅对第一条记录判断fieldnames，存在漏洞）
    """
    if not results:
        print("没有数据可保存")
        return

    # 固定字段顺序（不依赖任意一条记录的keys，避免顺序不稳定）
    fieldnames = ['文件名', '证券代码', '证券简称', '公告编号', '公告时间',
                  '失信被执行人姓名', '案号', '执行文号']

    try:
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
            # extrasaction='ignore'：遇到 fieldnames 之外的键（如'错误'）直接忽略，不抛异常
            writer = csv.DictWriter(f, fieldnames=fieldnames, extrasaction='ignore')
            writer.writeheader()
            for result in results:
                writer.writerow(result)

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
