"""
失信被执行人信息提取工具 v2
新增/修改：
  1. 删除执行文号提取，所有 PDF 均无条件提取案号
  2. 提取失信被执行人对应职位（姓名前后出现职位词时记录）
  3. 失信被执行人公司名称独立识别并单独存放一列（与姓名共用定位词逻辑）
  4. 案号提取不再依赖是否存在姓名
"""

import logging
import os
import re
import docx
import csv
from datetime import datetime

# ── 日志配置 ──────────────────────────────────────────────────────────────────
_log_dir = os.path.dirname(os.path.abspath(__file__))
_log_file = os.path.join(_log_dir, f"extract_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log")

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s [%(levelname)s] %(message)s',
    handlers=[
        logging.FileHandler(_log_file, encoding='utf-8'),
        logging.StreamHandler(),
    ]
)
logger = logging.getLogger(__name__)

# ── pdfplumber ────────────────────────────────────────────────────────────────
try:
    import pdfplumber
    PDFPLUMBER_AVAILABLE = True
except ImportError:
    pdfplumber = None
    PDFPLUMBER_AVAILABLE = False
    logger.warning("未检测到 pdfplumber，无法处理 PDF。请 pip install pdfplumber。")

# ── jieba（可选） ──────────────────────────────────────────────────────────────
try:
    import jieba
    import jieba.posseg as pseg  # noqa: F401
    JIEBA_AVAILABLE = True
except ImportError:
    jieba = None
    pseg = None
    JIEBA_AVAILABLE = False
    logger.warning("未检测到 jieba，姓名分词优化将关闭。可 pip install jieba。")

# ══════════════════════════════════════════════════════════════════════════════
# 模块级常量
# ══════════════════════════════════════════════════════════════════════════════

# 职位词表（按长度降序，确保长词优先匹配）
POSITION_WORDS: list[str] = sorted([
    '董事长', '董事会', '执行董事', '独立董事', '非执行董事', '董事',
    '监事长', '监事会', '监事',
    '总经理', '副总经理', '总裁', '副总裁', '总监', '副总监',
    '首席执行官', '首席财务官', '首席运营官',
    '法定代表人', '实际控制人', '控股股东',
    '主席', '副主席', '委员', '秘书',
], key=len, reverse=True)

# 职位词末两字集合：检测"董事长武力"→"长武"粘连
POSITION_WORD_ENDINGS: frozenset[str] = frozenset(
    pw[-2:] for pw in POSITION_WORDS if len(pw) >= 2
)

# 职位词正则（用于前后文职位提取）
_POSITION_PAT: re.Pattern = re.compile(
    '(' + '|'.join(re.escape(pw) for pw in POSITION_WORDS) + ')',
    re.UNICODE,
)

# 无效姓名词汇
INVALID_NAME_WORDS: frozenset[str] = frozenset([
    '公司', '集团', '企业', '股份', '有限', '责任', '法院', '银行', '证券', '保险',
    '董事', '董事长', '总经理', '高管', '被', '长', '主席', '监事', '经理', '主任',
    '科长', '处长', '局长', '部长', '总裁', '总监', '委员', '代表',
])

# 公司相关关键词
COMPANY_KEYWORDS: frozenset[str] = frozenset([
    '公司', '股份', '有限', '集团', '企业', '责任公司', '有限公司', '有限责任公司',
])

# 地名前缀
LOCATION_PREFIXES: tuple[str, ...] = (
    '北京', '上海', '广州', '深圳', '天津', '重庆', '杭州', '南京', '武汉', '成都',
    '西安', '郑州', '沈阳', '青岛', '大连', '宁波', '厦门', '苏州', '无锡', '常州',
    '徐州', '南通', '扬州', '镇江', '泰州', '盐城', '淮安', '连云港', '宿迁', '温州',
    '绍兴', '嘉兴', '湖州', '金华', '衢州', '台州', '丽水', '舟山', '福州', '泉州',
    '漳州', '莆田', '宁德', '龙岩', '三明', '南平', '济南', '烟台', '潍坊',
    '临沂', '济宁', '淄博', '泰安', '枣庄', '日照', '莱芜', '聊城', '德州', '滨州',
    '菏泽', '东营', '威海', '山东', '云南', '江苏', '江西', '安徽', '福建', '江阴',
    '广西', '广东', '甘肃', '焦作',
)

# 失信相关关键词
SHIXIN_KEYWORDS: tuple[str, ...] = (
    '失信被执行', '纳入失信', '列入失信', '失信名单', '失信人名单',
)

# 限制消费相关关键词
XIAOFEI_KEYWORDS: tuple[str, ...] = (
    '限制消费', '限消', '消费令', '限消人员',
)

# 跨段落截断标记
CROSS_SECTION_MARKERS: tuple[str, ...] = ('限制消费', '被限制消费', '限消')

# ── 同义词/近义说法扩展表 ────────────────────────────────────────────────────
SYNONYM_GROUPS: dict[str, list[str]] = {
    'inducted': [
        '被纳入失信被执行人名单', '纳入失信被执行人名单',
        '被纳入失信被执行名单',   '纳入失信被执行名单',
        '被列入失信被执行人名单', '列入失信被执行人名单',
        '被列入失信被执行名单',   '列入失信被执行名单',
        '被认定为失信被执行人',   '认定为失信被执行人',
        '被纳入失信人名单',       '纳入失信人名单',
        '被列为失信被执行人',     '列为失信被执行人',
        '被加入失信被执行人名单', '加入失信被执行人名单',
        '被登记为失信被执行人',   '登记为失信被执行人',
        '被录入失信被执行人名单', '录入失信被执行人名单',
        '被纳入失信名单',         '纳入失信名单',
        '列入失信名单',           '被列入失信名单',
        '被纳扩失信被执行人',     '被纳扩失信',
    ],
    'field_name': [
        '失信被执行人（姓名/名称）', '失信被执行人(姓名/名称)',
        '失信被执行人姓名/名称',     '失信被执行人姓名',
        '失信被执行人名称',          '失信人姓名',
        '失信人名称',                '被执行人姓名',
        '被执行人名称',
    ],
}


# ══════════════════════════════════════════════════════════════════════════════
# 工具函数
# ══════════════════════════════════════════════════════════════════════════════

def build_synonym_pattern(group_key: str, allow_spaces: bool = False) -> str:
    """根据同义词组生成合并正则，按长度降序确保长模式优先匹配。"""
    synonyms = SYNONYM_GROUPS.get(group_key, [])
    parts = []
    for phrase in synonyms:
        if allow_spaces:
            escaped = r'\s*'.join(re.escape(ch) for ch in phrase)
        else:
            escaped = re.escape(phrase)
        parts.append(escaped)
    parts.sort(key=len, reverse=True)
    return '(?:' + '|'.join(parts) + ')'


def load_surnames(surname_file_path: str) -> list[str]:
    """从姓氏表 docx 文件中加载姓氏（支持多字姓氏）。"""
    surnames: set[str] = set()
    try:
        doc = docx.Document(surname_file_path)
        for para in doc.paragraphs:
            for part in para.text.strip().split('、'):
                surname = part.strip()
                if surname:
                    surnames.add(surname)
        if not surnames:
            logger.warning("姓氏文件为空或未提取到任何姓氏！")
    except Exception as exc:
        logger.error("加载姓氏文件时出错: %s", exc)
        return []
    return list(surnames)


def normalize_content(content: str) -> str:
    """将换行符替换为空格并合并连续空白。"""
    content = content.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
    return re.sub(r'\s+', ' ', content)


def _init_jieba(surnames: list[str]) -> None:
    """初始化 jieba：注入职位词和姓氏，确保分词边界正确。"""
    if not JIEBA_AVAILABLE:
        return
    for word in POSITION_WORDS:
        jieba.add_word(word, freq=50000, tag='n')
    for s in surnames:
        jieba.add_word(s, freq=10000, tag='nr')
    jieba.initialize()


# ══════════════════════════════════════════════════════════════════════════════
# 姓名 / 公司验证
# ══════════════════════════════════════════════════════════════════════════════

def is_company_name(candidate: str, content: str) -> bool:
    """判断候选字符串是否为公司名称（宽松判断，用于姓名过滤）。"""
    if any(kw in candidate for kw in COMPANY_KEYWORDS):
        return True
    if len(candidate) > 8:
        return True
    for loc in LOCATION_PREFIXES:
        if candidate.startswith(loc):
            if re.search(rf'{re.escape(candidate)}[^\s，。；;、]*公司', content):
                return True
    return False


def is_valid_shixin_company(candidate: str) -> bool:
    """
    判断候选字符串是否为合法的失信被执行人公司名称。
    要求：
      - 含有公司类关键词
      - 长度在 4~30 字之间（过滤极短误匹配）
      - 不以"的"等虚词结尾
    """
    if not candidate:
        return False
    if not any(kw in candidate for kw in COMPANY_KEYWORDS):
        return False
    if not 4 <= len(candidate) <= 30:
        return False
    if candidate.endswith(('的', '及', '与', '和', '或')):
        return False
    return True


def is_valid_name(name: str, surnames: list[str], content: str) -> bool:
    """验证候选字符串是否为合法中文姓名。"""
    if not name:
        return False
    if not re.match(r'^[\u4e00-\u9fa5]+$', name):
        return False
    if not 2 <= len(name) <= 4:
        return False
    if '被' in name or name.startswith('被'):
        return False
    if not surnames:
        return False

    starts_with_surname = False
    for surname in surnames:
        if name.startswith(surname):
            given = name[len(surname):]
            if not 1 <= len(given) <= 2:
                return False
            starts_with_surname = True
            break
    if not starts_with_surname:
        return False

    if any(word in name for word in INVALID_NAME_WORDS):
        return False
    if '董事' in name:
        return False

    for loc in LOCATION_PREFIXES:
        if name.startswith(loc):
            if re.search(rf'{re.escape(name)}[^\s，。；;、]*公司', content):
                return False

    if content.count(name) < 1:
        return False

    if len(name) == 2 and name[0] == '司':
        if re.search(rf'公司{re.escape(name)}', content):
            return False

    # 职位词末字粘连检测
    search_start = 0
    found_valid = False
    while True:
        idx = content.find(name, search_start)
        if idx == -1:
            break
        if idx > 0:
            two_char = content[idx - 1] + name[0]
            if two_char in POSITION_WORD_ENDINGS:
                search_start = idx + 1
                continue
        found_valid = True
        break
    if not found_valid:
        return False

    return True


def is_only_xiao_fei_person(name: str, content: str) -> bool:
    """判断姓名是否仅出现在"限制消费"语境而非失信被执行人语境。"""
    found_shixin = False
    occurrence_count = 0
    xiaofei_count = 0
    start = 0

    while True:
        idx = content.find(name, start)
        if idx == -1:
            break
        occurrence_count += 1
        window = content[max(0, idx - 40): idx + len(name) + 40]
        has_shixin = any(kw in window for kw in SHIXIN_KEYWORDS)
        has_xiaofei = any(kw in window for kw in XIAOFEI_KEYWORDS)
        if has_shixin:
            found_shixin = True
        if has_xiaofei and not has_shixin:
            xiaofei_count += 1
        start = idx + 1

    if found_shixin:
        return False
    if occurrence_count > 0 and xiaofei_count == occurrence_count:
        return True
    return False


# ══════════════════════════════════════════════════════════════════════════════
# 职位提取
# ══════════════════════════════════════════════════════════════════════════════

def extract_position_for_name(name: str, content: str) -> str:
    """
    在全文中查找姓名的所有出现位置，检测前后各 15 字窗口内是否存在职位词。
    只要任一处出现职位词即记录并返回（取首个匹配到的职位词）。
    找不到则返回空字符串。
    """
    window_size = 15
    start = 0
    while True:
        idx = content.find(name, start)
        if idx == -1:
            break
        left = content[max(0, idx - window_size): idx]
        right = content[idx + len(name): idx + len(name) + window_size]
        context = left + right
        m = _POSITION_PAT.search(context)
        if m:
            return m.group(1)
        start = idx + 1
    return ''


# ══════════════════════════════════════════════════════════════════════════════
# 正则模式预编译（姓名 + 公司名称共用定位词）
# ══════════════════════════════════════════════════════════════════════════════

def _build_extraction_patterns():
    """
    构造所有提取模式，统一编译。
    姓名模式与公司名称模式共用相同的定位词（字段标签 / 纳入列入 / 同义词组），
    区别仅在于捕获组的内容长度与公司词验证。
    """
    # 人名捕获组：2-4 汉字
    name_re = r'([\u4e00-\u9fa5]{2,4})'
    # 公司名称捕获组：4-30 汉字（宽松，后续由 is_valid_shixin_company 精筛）
    company_re = r'([\u4e00-\u9fa5]{4,30}(?:公司|集团|企业))'

    position_prefix = (
        r'(?:(?:公司\s*)?(?:及\s*)?'
        r'(?:董事长|执行董事|独立董事|非执行董事|监事长|总经理|副总经理|'
        r'总裁|法定代表人|实际控制人|控股股东|董事|监事|主席)\s*)?'
    )

    # ── 字段标签变体 ──────────────────────────────────────────────────────────
    field_label_vars = [
        r'失信被执行人[（(]姓名/名称[）)][：:]',
        r'失信被执行人姓名/名称[：:]',
        r'失信被执行人姓名[：:]',
        r'失信被执行人名称[：:]',
        r'失信人姓名[：:]',
        r'被执行人姓名[：:]',
    ]

    # A. 字段标签 → 姓名（短，直接捕获 2-4 字）
    pat_a_name = [
        re.compile(rf'{lbl}\s*{name_re}', re.UNICODE)
        for lbl in field_label_vars
    ]
    # A_LONG. 字段标签 → 整行（二次提取姓名）
    pat_a_long = [
        re.compile(rf'{lbl}\s*([^\r\n]+)', re.UNICODE)
        for lbl in field_label_vars
    ]
    # A_COMPANY. 字段标签 → 公司名称
    pat_a_company = [
        re.compile(rf'{lbl}\s*{company_re}', re.UNICODE)
        for lbl in field_label_vars
    ]

    # ── 纳入/列入模式 ─────────────────────────────────────────────────────────
    induct_suffix_norm = r'(?:被\s*)?(?:纳入|列入)\s*失信被执行人?名单'
    induct_suffix_spaced = build_synonym_pattern('inducted', allow_spaces=True)

    # B. 归一化文本 → 姓名
    pat_b_name = re.compile(
        position_prefix + name_re + r'\s*' + induct_suffix_norm, re.UNICODE)
    # B_SPACED. 原始文本 → 姓名（散字版）
    pat_b_spaced_name = re.compile(
        position_prefix + name_re + r'\s*' + induct_suffix_spaced,
        re.UNICODE | re.DOTALL)
    # B_COMPANY. 归一化文本 → 公司名称
    pat_b_company = re.compile(
        r'(?:公司\s*)?' + company_re + r'\s*' + induct_suffix_norm, re.UNICODE)

    # ── 同义词扩展通用正则 ────────────────────────────────────────────────────
    inducted_pat = build_synonym_pattern('inducted')
    field_pat = build_synonym_pattern('field_name')

    # C. 同义词 → 姓名
    pat_c_induct_name = re.compile(
        position_prefix + name_re + r'\s*' + inducted_pat, re.UNICODE)
    pat_c_field_name = re.compile(
        field_pat + r'[：:]\s*' + name_re, re.UNICODE)
    # C_COMPANY. 同义词 → 公司名称
    pat_c_induct_company = re.compile(
        r'(?:公司\s*)?' + company_re + r'\s*' + inducted_pat, re.UNICODE)
    pat_c_field_company = re.compile(
        field_pat + r'[：:]\s*' + company_re, re.UNICODE)

    return (
        pat_a_name, pat_a_long, pat_a_company,
        pat_b_name, pat_b_spaced_name, pat_b_company,
        pat_c_induct_name, pat_c_field_name,
        pat_c_induct_company, pat_c_field_company,
    )


(
    _PAT_A_NAME, _PAT_A_LONG, _PAT_A_COMPANY,
    _PAT_B_NAME, _PAT_B_SPACED_NAME, _PAT_B_COMPANY,
    _PAT_C_INDUCT_NAME, _PAT_C_FIELD_NAME,
    _PAT_C_INDUCT_COMPANY, _PAT_C_FIELD_COMPANY,
) = _build_extraction_patterns()


# ══════════════════════════════════════════════════════════════════════════════
# 姓名 & 公司名称提取（统一入口）
# ══════════════════════════════════════════════════════════════════════════════

def _extract_names_from_long_match(text: str, surnames: list[str], content: str) -> list[str]:
    """从字段标签后的长串中二次提取姓名，跨限消段落时先截断。"""
    for marker in CROSS_SECTION_MARKERS:
        idx = text.find(marker)
        if idx != -1:
            text = text[:idx]
            break

    names: list[str] = []
    if is_valid_name(text.strip(), surnames, content):
        names.append(text.strip())
        return names

    if not surnames:
        return names
    sorted_surnames = sorted(surnames, key=len, reverse=True)
    surname_regex = '|'.join(re.escape(s) for s in sorted_surnames)
    for m in re.finditer(rf'({surname_regex}[\u4e00-\u9fa5]{{1,2}})', text):
        candidate = m.group(1)
        if is_valid_name(candidate, surnames, content):
            names.append(candidate)
    return names


def _extract_names_jieba(content: str, surnames: list[str]) -> list[str]:
    """使用 jieba 分词修正职位词+姓名粘连问题（仅在职位词窗口内运行）。"""
    if not JIEBA_AVAILABLE:
        return []
    names: list[str] = []
    for pos_word in POSITION_WORDS:
        start = 0
        while True:
            idx = content.find(pos_word, start)
            if idx == -1:
                break
            window = content[max(0, idx - 10): min(len(content), idx + len(pos_word) + 15)]
            seg_list = list(jieba.cut(window))
            for i, seg in enumerate(seg_list):
                if seg == pos_word or seg in POSITION_WORDS:
                    for look_ahead in range(1, 3):
                        if i + look_ahead < len(seg_list):
                            candidate = ''.join(seg_list[i + 1: i + 1 + look_ahead]).strip()
                            if candidate and is_valid_name(candidate, surnames, content):
                                names.append(candidate)
            start = idx + 1
    return names


def extract_shixin_entities(
    content: str, surnames: list[str]
) -> tuple[list[str], list[str], dict[str, str]]:
    """
    统一提取失信被执行人的：
      - 自然人姓名列表
      - 公司名称列表
      - 自然人 → 职位 映射字典（无职位则值为空字符串）

    流程：
      1. 归一化文本上运行所有正则（单次遍历）
      2. jieba 辅助修正粘连
      3. 验证 + 过滤仅限消姓名
      4. 对每个有效姓名提取职位
    """
    normalized = normalize_content(content)
    name_candidates: set[str] = set()
    company_candidates: set[str] = set()

    # ── A. 字段标签 → 姓名 ────────────────────────────────────────────────
    for pat in _PAT_A_NAME:
        for m in pat.finditer(normalized):
            c = m.group(1).strip()
            if not is_company_name(c, normalized):
                name_candidates.add(c)

    # ── A_LONG. 字段标签 → 长串（二次提取姓名） ──────────────────────────
    for pat in _PAT_A_LONG:
        for m in pat.finditer(normalized):
            raw = m.group(1).strip()
            if is_company_name(raw, normalized):
                continue
            if is_valid_name(raw, surnames, normalized):
                name_candidates.add(raw)
            else:
                name_candidates.update(
                    _extract_names_from_long_match(raw, surnames, normalized)
                )

    # ── A_COMPANY. 字段标签 → 公司名称 ───────────────────────────────────
    for pat in _PAT_A_COMPANY:
        for m in pat.finditer(normalized):
            c = m.group(1).strip()
            if is_valid_shixin_company(c):
                company_candidates.add(c)

    # ── B. 纳入/列入 → 姓名（归一化） ───────────────────────────────────
    for m in _PAT_B_NAME.finditer(normalized):
        c = m.group(1).strip()
        if not c.startswith('司'):
            name_candidates.add(c)

    # ── B_SPACED. 换行散字 → 姓名（原始文本） ────────────────────────────
    for m in _PAT_B_SPACED_NAME.finditer(content):
        c = m.group(1).strip()
        if not c.startswith('司'):
            name_candidates.add(c)

    # ── B_COMPANY. 纳入/列入 → 公司名称 ─────────────────────────────────
    for m in _PAT_B_COMPANY.finditer(normalized):
        c = m.group(1).strip()
        if is_valid_shixin_company(c):
            company_candidates.add(c)

    # ── C. 同义词扩展 → 姓名 ─────────────────────────────────────────────
    for m in _PAT_C_INDUCT_NAME.finditer(normalized):
        c = m.group(1).strip()
        if not c.startswith('司'):
            name_candidates.add(c)
    for m in _PAT_C_FIELD_NAME.finditer(normalized):
        name_candidates.add(m.group(1).strip())

    # ── C_COMPANY. 同义词扩展 → 公司名称 ────────────────────────────────
    for m in _PAT_C_INDUCT_COMPANY.finditer(normalized):
        c = m.group(1).strip()
        if is_valid_shixin_company(c):
            company_candidates.add(c)
    for m in _PAT_C_FIELD_COMPANY.finditer(normalized):
        c = m.group(1).strip()
        if is_valid_shixin_company(c):
            company_candidates.add(c)

    # ── jieba 职位粘连修正 → 姓名 ────────────────────────────────────────
    if JIEBA_AVAILABLE:
        name_candidates.update(_extract_names_jieba(normalized, surnames))

    # ── 验证姓名 + 过滤仅限消 ────────────────────────────────────────────
    valid_names: list[str] = []
    for name in name_candidates:
        if is_valid_name(name, surnames, normalized):
            if not is_only_xiao_fei_person(name, normalized):
                valid_names.append(name)

    # ── 为每个有效姓名提取职位 ────────────────────────────────────────────
    name_to_position: dict[str, str] = {
        name: extract_position_for_name(name, normalized)
        for name in valid_names
    }

    # ── 去重公司名称（去掉与姓名重复的误匹配） ────────────────────────────
    valid_companies: list[str] = [
        c for c in company_candidates if c not in valid_names
    ]

    return valid_names, valid_companies, name_to_position


# ══════════════════════════════════════════════════════════════════════════════
# 案号提取（无条件，不依赖姓名是否存在）
# ══════════════════════════════════════════════════════════════════════════════

_CASE_PATTERNS: list[re.Pattern] = [
    re.compile(r"""
        案号[：:]\s*
        (
        [\(（]\d{4}[\)）]\s*
        [\u4e00-\u9fa5\d\s]+?
        (?:民[初终再审再]?\s*|刑[初终再]?\s*|行[初终]?\s*
          |执(?:恢|异|保|行)?\s*|知民[初终]?\s*|知刑[初终]?\s*|商[初终]?\s*)
        \s*\d+\s*号
        )
    """, re.VERBOSE | re.UNICODE),
    re.compile(r"""
        (?:案件案号|案由案号|案号)[：:]\s*
        ([\(（]\d{4}[\)）][^，。；;、号]{5,30}?号)
    """, re.VERBOSE | re.UNICODE),
    re.compile(r'案号[：:]\s*(\(?\d{4}\)?[^，。；;、]{5,30}?号)', re.UNICODE),
]


def extract_case_numbers(content: str) -> list[str]:
    """提取全文案号（无条件执行，所有 PDF 均调用）。"""
    results: set[str] = set()
    for pat in _CASE_PATTERNS:
        for m in pat.finditer(content):
            cleaned = re.sub(r'\s+', ' ', m.group(1).strip())
            results.add(cleaned)
    return list(results)


# ══════════════════════════════════════════════════════════════════════════════
# PDF 提取 / 文件处理 / 输出
# ══════════════════════════════════════════════════════════════════════════════

def extract_text_from_pdf(file_path: str) -> tuple[str, list[int]]:
    """从 PDF 提取文本，返回 (全文字符串, 扫描页编号列表)。"""
    if not PDFPLUMBER_AVAILABLE:
        raise RuntimeError("pdfplumber 未安装，无法提取 PDF 文本。")

    all_text: list[str] = []
    scanned_pages: list[int] = []

    with pdfplumber.open(file_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text and text.strip():
                all_text.append(text)
            else:
                scanned_pages.append(i)

    if scanned_pages:
        logger.info("  [提示] 以下页面疑似扫描件，已跳过：第 %s 页", scanned_pages)

    return '\n'.join(all_text), scanned_pages


_DATE_PATTERNS: list[re.Pattern] = [
    re.compile(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日'),
    re.compile(r'(\d{4})-(\d{1,2})-(\d{1,2})'),
    re.compile(r'(\d{4})/(\d{1,2})/(\d{1,2})'),
    re.compile(r'(\d{4})\.(\d{1,2})\.(\d{1,2})'),
]


def extract_info_from_file(file_path: str, surnames: list[str]) -> dict:
    """
    从单个 PDF 文件提取所有字段。
    输出字段：
      证券代码 / 证券简称 / 公告编号 / 公告时间
      失信被执行人姓名 / 失信被执行人职位 / 失信被执行人公司
      案号
    """
    content, _ = extract_text_from_pdf(file_path)
    empty = {
        '证券代码': '', '证券简称': '', '公告编号': '', '公告时间': '',
        '失信被执行人姓名': '', '失信被执行人职位': '', '失信被执行人公司': '',
        '案号': '',
    }

    if not content.strip():
        logger.warning("  [警告] 文件无可提取文本（可能全为扫描件），跳过。")
        return empty

    result = dict(empty)

    # ── 证券代码 ──────────────────────────────────────────────────────────
    m = re.search(r'证券代码[：:]\s*(\d{6})', content)
    if m:
        result['证券代码'] = m.group(1)

    # ── 证券简称 ──────────────────────────────────────────────────────────
    m = re.search(r'证券简称[：:]\s*([^\s]+)', content)
    if m:
        result['证券简称'] = m.group(1)

    # ── 公告编号 ──────────────────────────────────────────────────────────
    m = re.search(r'公告编号[：:]\s*([^\s]+)', content)
    if m:
        result['公告编号'] = m.group(1)

    # ── 公告时间（取文末10行最后一个日期） ───────────────────────────────
    lines = content.strip().split('\n')
    for line in reversed(lines[-10:]):
        for dpat in _DATE_PATTERNS:
            dm = dpat.search(line)
            if dm:
                year, month, day = dm.groups()
                try:
                    result['公告时间'] = f"{year}年{int(month)}月{int(day)}日"
                    break
                except ValueError:
                    continue
        if result['公告时间']:
            break

    # ── 案号（无条件提取） ────────────────────────────────────────────────
    case_nums = extract_case_numbers(content)
    if case_nums:
        result['案号'] = '，'.join(case_nums)

    # ── 失信被执行人：姓名 / 职位 / 公司 ────────────────────────────────
    names, companies, name_to_pos = extract_shixin_entities(content, surnames)

    if names:
        sorted_names = sorted(set(names))
        result['失信被执行人姓名'] = '，'.join(sorted_names)

        # 职位列：与姓名一一对应，无职位输出空值，多人用顿号分隔
        # 格式："姓名1:职位1，姓名2:（无）"——此处选择仅输出职位字符串与姓名顺序对应
        positions = [name_to_pos.get(n, '') for n in sorted_names]
        result['失信被执行人职位'] = '，'.join(positions)

    if companies:
        result['失信被执行人公司'] = '，'.join(sorted(set(companies)))

    return result


def extract_all_files(folder_path: str, surname_file_path: str) -> list[dict]:
    """提取文件夹内所有 PDF 文件的信息。"""
    surnames = load_surnames(surname_file_path)
    if not surnames:
        logger.warning("未能加载姓氏表，姓名提取功能将受限！")
    logger.info("加载了 %d 个姓氏", len(surnames))

    _init_jieba(surnames)
    if JIEBA_AVAILABLE:
        logger.info("jieba 分词已启用（职位粘连修正）")

    pdf_files = sorted(f for f in os.listdir(folder_path) if f.lower().endswith('.pdf'))
    logger.info("找到 %d 个 PDF 文件", len(pdf_files))

    all_results: list[dict] = []
    for filename in pdf_files:
        file_path = os.path.join(folder_path, filename)
        logger.info("\n处理文件: %s", filename)
        try:
            result = extract_info_from_file(file_path, surnames)
            record = {'文件名': filename, **result}
            all_results.append(record)

            logger.info("  证券代码: %s  证券简称: %s  公告时间: %s",
                        result['证券代码'], result['证券简称'], result['公告时间'])
            if result['失信被执行人姓名']:
                logger.info("  失信被执行人（自然人）: %s", result['失信被执行人姓名'])
                logger.info("  对应职位: %s",
                            result['失信被执行人职位'] or '（未提取到职位）')
            else:
                logger.info("  未提取到失信被执行人自然人姓名")
            if result['失信被执行人公司']:
                logger.info("  失信被执行人（公司）: %s", result['失信被执行人公司'])
            if result['案号']:
                logger.info("  案号: %s", result['案号'])
            else:
                logger.info("  未提取到案号")

        except Exception as exc:
            logger.error("处理文件 %s 时出错: %s", filename, exc)
            all_results.append({
                '文件名': filename,
                '证券代码': '', '证券简称': '', '公告编号': '', '公告时间': '',
                '失信被执行人姓名': '', '失信被执行人职位': '', '失信被执行人公司': '',
                '案号': '',
                '错误': str(exc),
            })

    return all_results


def save_to_csv(results: list[dict], output_file: str) -> None:
    """将结果保存到 CSV 文件。"""
    if not results:
        logger.info("没有数据可保存")
        return

    fieldnames = [k for k in results[0].keys() if k != '错误']
    try:
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for row in results:
                writer.writerow({k: v for k, v in row.items() if k != '错误'})
        logger.info("\n结果已保存到: %s", output_file)
    except PermissionError:
        logger.error("[错误] 无法保存到 %s，文件可能已被其他程序打开", output_file)


# ══════════════════════════════════════════════════════════════════════════════
# 主程序
# ══════════════════════════════════════════════════════════════════════════════

def main() -> None:
    folder_path = input("请输入包含PDF文件的文件夹路径: ")
    if not os.path.exists(folder_path):
        logger.error("文件夹不存在: %s", folder_path)
        return

    surname_file_path = r"D:\This computer\desktop\姓氏表\姓氏表.docx"
    if not os.path.exists(surname_file_path):
        logger.error("姓氏文件不存在: %s", surname_file_path)
        logger.error("请确保姓氏表文件存在，或修改代码中的路径")
        return

    results = extract_all_files(folder_path, surname_file_path)

    if results:
        output_file = os.path.join(folder_path, '失信人信息提取结果_完整版.csv')
        save_to_csv(results, output_file)

        logger.info("\n处理完成！共处理 %d 条记录", len(results))
        logger.info("成功提取失信被执行人姓名: %d 条",
                    sum(1 for r in results if r.get('失信被执行人姓名')))
        logger.info("成功提取失信被执行人公司: %d 条",
                    sum(1 for r in results if r.get('失信被执行人公司')))
        logger.info("成功提取案号: %d 条",
                    sum(1 for r in results if r.get('案号')))

        extracted_count = sum(1 for r in results if r.get('失信被执行人姓名'))
        if extracted_count > 0:
            all_names: list[str] = []
            for r in results:
                if r.get('失信被执行人姓名'):
                    all_names.extend(r['失信被执行人姓名'].split('，'))
            logger.info("\n提取到的所有失信被执行人姓名:")
            for name in sorted(set(filter(None, all_names))):
                logger.info("  %s", name)
    else:
        logger.info("没有提取到任何数据")


if __name__ == "__main__":
    main()