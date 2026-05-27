from __future__ import annotations

"""
失信被执行人信息提取工具 v10 (终极分类强化版)
新增/修改：
  1. 分类隔离：对“文件名”与“正文”进行分开检索。优先匹配文件名，文件名无明确结论再匹配正文，彻底切断上下文跨域污染。
  2. 优先级重构：优先判定“撤销”，防止撤销公告中复述“曾被纳入”导致的误判。
  3. 精准排雷：职务库使用负向断言 `董事(?!会)` 及 `监事(?!会)`，避开落款干扰。
  4. 文本清洗：对正文执行“标准免责声明”剔除与“限制消费”语句剥离。
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
    logger.warning("未检测到 pdfplumber，无法处理 PDF。")

# ── jieba ─────────────────────────────────────────────────────────────────────
try:
    import jieba
    import jieba.posseg as pseg

    JIEBA_AVAILABLE = True
except ImportError:
    jieba = None
    pseg = None
    JIEBA_AVAILABLE = False

# ── LAC ───────────────────────────────────────────────────────────────────────
try:
    from LAC import LAC as _LAC

    _lac_instance = _LAC(mode='lac')
    LAC_AVAILABLE = True
    logger.info("LAC 已加载")
except Exception as e:
    _lac_instance = None
    LAC_AVAILABLE = False

# ══════════════════════════════════════════════════════════════════════════════
# 模块级常量
# ══════════════════════════════════════════════════════════════════════════════

POSITION_WORDS: list[str] = sorted([
    '执行董事', '独立董事', '非执行董事', '董事长', '董事', '监事长', '监事',
    '总经理', '副总经理', '总裁', '副总裁', '总监', '副总监',
    '首席执行官', '首席财务官', '首席运营官', '法定代表人', '实际控制人', '控股股东',
    '主席', '副主席', '委员', '秘书',
], key=len, reverse=True)

_INSTITUTION_WORDS: frozenset[str] = frozenset(['董事会', '监事会', '股东会', '委员会'])

POSITION_WORD_ENDINGS: frozenset[str] = frozenset(
    pw[-2:] for pw in POSITION_WORDS if len(pw) >= 2
)

_POSITION_PAT: re.Pattern = re.compile(
    r'(?<!' + r'(?<!'.join(['']) + r')'
    if False else
    '(' + '|'.join(re.escape(pw) for pw in POSITION_WORDS) + ')',
    re.UNICODE,
)

INVALID_NAME_WORDS: frozenset[str] = frozenset([
    '公司', '集团', '企业', '股份', '有限', '责任', '法院', '银行', '证券', '保险',
    '董事', '董事长', '总经理', '高管', '被', '长', '主席', '监事', '经理', '主任',
    '科长', '处长', '局长', '部长', '总裁', '总监', '委员', '代表',
    '任免', '任命', '辞职', '离职', '决定', '选举', '通知', '解除'
])

COMPANY_KEYWORDS: frozenset[str] = frozenset([
    '公司', '股份', '有限', '集团', '企业', '责任公司', '有限公司', '有限责任公司', '厂'
])

LOCATION_PREFIXES: tuple[str, ...] = (
    '北京', '上海', '广州', '深圳', '天津', '重庆', '杭州', '南京', '武汉', '成都',
    '西安', '郑州', '沈阳', '青岛', '大连', '宁波', '厦门', '苏州', '无锡', '常州',
    '徐州', '南通', '扬州', '镇江', '泰州', '盐城', '淮安', '连云港', '宿迁', '温州',
    '绍兴', '嘉兴', '湖州', '金华', '衢州', '台州', '丽水', '舟山', '福州', '泉州',
    '漳州', '莆田', '宁德', '龙岩', '三明', '南平', '济南', '烟台', '潍坊', '临沂',
    '济宁', '淄博', '泰安', '枣庄', '日照', '莱芜', '聊城', '德州', '滨州', '菏泽',
    '东营', '威海', '山东', '云南', '江苏', '江西', '安徽', '福建', '江阴', '广西',
    '广东', '甘肃', '焦作', '浙江', '河南', '河北', '湖南', '湖北'
)

SHIXIN_KEYWORDS: tuple[str, ...] = ('失信被执行', '纳入失信', '列入失信', '失信名单', '失信人名单')
XIAOFEI_KEYWORDS: tuple[str, ...] = ('限制消费', '限消', '消费令', '限消人员')
CROSS_SECTION_MARKERS: tuple[str, ...] = ('限制消费', '被限制消费', '限消')

# ── 同义词扩展（包含新语料全量扩充） ─────────────────────────────
SYNONYM_GROUPS: dict[str, list[str]] = {
    'inducted': [
        '被纳入失信被执行人名单', '纳入失信被执行人名单', '被纳入失信被执行名单', '纳入失信被执行名单',
        '被列入失信被执行人名单', '列入失信被执行人名单', '被列入失信被执行名单', '列入失信被执行名单',
        '被认定为失信被执行人', '认定为失信被执行人', '被纳入失信人名单', '纳入失信人名单',
        '被列为失信被执行人', '列为失信被执行人', '被加入失信被执行人名单', '加入失信被执行人名单',
        '被纳入失信名单', '纳入失信名单', '列入失信名单', '被列入失信名单', '被纳扩失信被执行人',
        '被纳入失信被执行人', '纳入失信被执行人', '被列入失信被执行人', '列入失信被执行人',
        '被纳为失信被执行人名单', '纳为失信被执行人名单', '被纳为失信被执行人', '纳为失信被执行人',
        '被纳为失信', '纳为失信', '被纳入为失信', '被纳入失信执行人名单', '纳入失信执行人名单',
        '被纳入失信执行人', '纳入失信执行人', '被列入失信联合惩戒对象名单', '列入失信联合惩戒对象名单',
        '被列入失信联合惩戒对象', '列入失信联合惩戒对象', '失信联合惩戒对象', '惩戒对象',
        '增加失信信息', '新增失信信息'
    ],
    'revoked': [
        '撤销失信被执行人', '撤销失信', '撤销被纳入失信', '移出失信被执行人名单', '移出失信',
        '删除失信', '屏蔽失信', '退出失信',
        '从失信被执行人名单库中删除', '从失信被执行人名单中删除', '已撤销纳入失信被执行人名单',
        '已被撤销纳入失信被执行人名单', '撤销纳入失信被执行人名单', '已被撤销纳入失信', '撤销纳入失信',
        '已撤销被纳入失信被执行人名单', '撤销被纳入失信被执行人名单', '已撤销被纳入失信',
        '已被撤销失信被执行人名单', '已撤销失信被执行人名单', '从失信被执行人名单中撤销',
        '从失信名单中撤销', '撤销公司被纳入失信被执行人名单', '撤销公司控股股东被纳入失信被执行人名单',
        '删除失信信息', '删除公司失信信息', '删除董事失信信息', '法院删除', '撤销失信被执行人信息'
    ],
    'not_dishonest': [
        '不属于失信', '不为失信', '不是失信',
        '未被纳入失信被执行人名单', '未纳入失信被执行人名单', '未被列入失信被执行人名单', '未列入失信被执行人名单',
        '未被认定为失信被执行人', '未认定为失信被执行人', '未被纳入失信名单', '未纳入失信名单',
        '未被列入失信名单', '未列入失信名单', '不属于失信被执行人', '非失信被执行人',
        '未构成失信', '不存在失信', '未被列为失信被执行人', '不在失信名单',
        '不属于失信执行人', '未被纳入失信执行人名单'
    ],
    'field_name': [
        '失信被执行人（姓名/名称）', '失信被执行人(姓名/名称)', '失信被执行人姓名/名称', '失信被执行人姓名',
        '失信被执行人名称', '失信人姓名', '失信人名称', '被执行人姓名', '被执行人名称',
    ],
}


def build_synonym_pattern(group_key: str, allow_spaces: bool = False) -> str:
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
    surnames: set[str] = set()
    try:
        doc = docx.Document(surname_file_path)
        for para in doc.paragraphs:
            for part in para.text.strip().split('、'):
                if part.strip(): surnames.add(part.strip())
    except Exception as exc:
        logger.error("加载姓氏文件时出错: %s", exc)
        return []
    return list(surnames)


def normalize_content(content: str) -> str:
    content = content.replace('\r\n', ' ').replace('\n', ' ').replace('\r', ' ')
    return re.sub(r'\s+', ' ', content)


def _init_jieba(surnames: list[str]) -> None:
    if not JIEBA_AVAILABLE: return
    for word in POSITION_WORDS:
        jieba.add_word(word, freq=50000, tag='n')
    for s in surnames:
        jieba.add_word(s, freq=10000, tag='nr')
    jieba.initialize()


# ══════════════════════════════════════════════════════════════════════════════
# 标题及分类识别模块 (隔离判定机制 v10)
# ══════════════════════════════════════════════════════════════════════════════

def classify_document(filename: str, content: str) -> str:
    # ── 1. 预编译结构化组件库 ──
    not_dish_pat = re.compile(build_synonym_pattern('not_dishonest'))
    induct_pat = re.compile(build_synonym_pattern('inducted'))
    revoke_pat = re.compile(build_synonym_pattern('revoked'))

    # 严密职务库：负向断言排雷 董事(?!会)
    roles = r'(董事长|董事(?!会)|监事会主席|监事(?!会)|高级管理人员|高管|总裁|副总裁|总经理|副总经理|实际控制人|实控人|控股股东|法定代表人|法代|董事会秘书|董秘)'
    comps = r'(?:公司|本公司|上市公司|本企业)'
    conns = r'(?:及|与|和|及其|、|，|,)'

    # 严格的双主体匹配逻辑
    pat_a = rf'{comps}\s*{conns}\s*(?:原)?(?:{comps})?(?:部分)?{roles}'
    pat_b = rf'{roles}\s*{conns}\s*(?:原)?{comps}'
    both_pat = re.compile(rf'(?:{pat_a})|(?:{pat_b})')

    # 封装核心判定器
    def _eval_text_category(text: str) -> str:
        if not_dish_pat.search(text):
            return '不属于失信人'
        has_role = re.search(roles, text)

        # 【撤销优先】防止正文历史回顾污染
        if revoke_pat.search(text):
            if both_pat.search(text):
                return '公司及个人被撤销'
            elif has_role:
                return '个人被撤销'
            else:
                return '公司被撤销'

        if induct_pat.search(text):
            if both_pat.search(text):
                return '公司及个人被纳入'
            elif has_role:
                return '个人被纳入'
            else:
                return '公司被纳入'

        return '其他'

    # ── 2. 第一重隔离：仅对文件名进行审查 ──
    base_filename = os.path.splitext(filename)[0]
    # 清洗掉文件名中可能附带的“限消”语句干扰
    xiaofei_pat = re.compile(r'[,，、\s]*[^,，。；;\n]*?(?:限制消费|限消|消费令)[^,，。；;\n]*', re.UNICODE)
    clean_filename = xiaofei_pat.sub('', base_filename)

    file_cat = _eval_text_category(clean_filename)
    if file_cat != '其他':
        return file_cat  # 如果文件名判定成功，直接返回，绝不看正文

    # ── 3. 第二重后备：对正文前300字进行审查 ──
    header_text = content[:300].replace('\n', ' ').replace('\r', ' ')

    # 清洗：标准免责套话（防止“公司及董事会全体成员...”导致双主体误判）
    boilerplate_pat = re.compile(
        r'(?:本)?公司及(?:董事(?:会全体成员)?|全体董事|监事(?:会)?|高级管理人员|高管).*?(?:保证|承诺|确认).*?(?:真实|准确|完整|虚假记载|误导性陈述|重大遗漏)[^。]*。?',
        re.UNICODE
    )
    header_text = boilerplate_pat.sub('', header_text)

    # 清洗：限消隔离
    header_text = xiaofei_pat.sub('', header_text)

    return _eval_text_category(header_text)


# ══════════════════════════════════════════════════════════════════════════════
# LAC NER & 提取核心逻辑
# ══════════════════════════════════════════════════════════════════════════════
_SENT_SPLIT_PAT: re.Pattern = re.compile(r'[。；;\n，,]+')


def _sentences_with_shixin(content: str) -> list[str]:
    sentences = _SENT_SPLIT_PAT.split(content)
    result: list[str] = []
    for sent in sentences:
        if '失信' in sent:
            result.append(sent.strip())
    return result


def _lac_extract_names(sentences: list[str]) -> list[str]:
    if not LAC_AVAILABLE or not sentences: return []
    names: set[str] = set()
    try:
        results = _lac_instance.run(sentences)
        for words, tags in results:
            for word, tag in zip(words, tags):
                if tag == 'PER' and 2 <= len(word) <= 4:
                    names.add(word)
    except Exception as exc:
        pass
    return list(names)


_ANCHOR_SUFFIX_PAT: re.Pattern = re.compile(r'\s*' + build_synonym_pattern('inducted'), re.UNICODE)
_ANCHOR_PREFIX_PAT: re.Pattern = re.compile(build_synonym_pattern('field_name') + r'[：:]\s*', re.UNICODE)
_ANCHOR_POSITION_PREFIX_PAT: re.Pattern = re.compile(
    r'(?:' + '|'.join(re.escape(pw) for pw in POSITION_WORDS) + r')\s*', re.UNICODE)
_ANCHOR_WINDOW = 40


def _has_anchor(name: str, content: str) -> bool:
    start = 0
    while True:
        idx = content.find(name, start)
        if idx == -1: break
        left = content[max(0, idx - _ANCHOR_WINDOW): idx]
        right = content[idx + len(name): idx + len(name) + _ANCHOR_WINDOW]
        if (_ANCHOR_SUFFIX_PAT.match(right) or _ANCHOR_PREFIX_PAT.search(left) or _ANCHOR_POSITION_PREFIX_PAT.search(
                left)):
            return True
        start = idx + 1
    return False


# ══════════════════════════════════════════════════════════════════════════════
# 实体验证与清理
# ══════════════════════════════════════════════════════════════════════════════

def clean_company_name(name: str) -> str:
    for loc in LOCATION_PREFIXES:
        idx = name.find(loc)
        if 0 < idx < 6: return name[idx:]
    return name


def is_valid_shixin_company(candidate: str) -> bool:
    if not candidate: return False
    if candidate.startswith('关于'): return False
    if not 4 <= len(candidate) <= 30: return False
    if not candidate.endswith(('公司', '厂', '中心', '企业', '集团', '合伙')): return False
    return True


def is_company_name(candidate: str, content: str) -> bool:
    if any(kw in candidate for kw in COMPANY_KEYWORDS): return True
    if len(candidate) > 8: return True
    return False


def is_valid_name(name: str, surnames: list[str], content: str) -> bool:
    if not name: return False
    if not re.match(r'^[\u4e00-\u9fa5]+$', name): return False
    if not 2 <= len(name) <= 4: return False

    if name[-1] in ('及', '系', '将', '等', '的', '与', '和', '了', '人', '为', '被'):
        return False
    if '被' in name or name.startswith('被'): return False
    if not surnames: return False

    starts_with_surname = False
    for surname in surnames:
        if name.startswith(surname):
            given = name[len(surname):]
            if not 1 <= len(given) <= 2: return False
            starts_with_surname = True
            break
    if not starts_with_surname: return False

    if any(word in name for word in INVALID_NAME_WORDS): return False
    if '董事' in name: return False
    for loc in LOCATION_PREFIXES:
        if name.startswith(loc): return False

    if content.count(name) < 1: return False
    if len(name) == 2 and name[0] == '司': return False

    search_start = 0
    found_valid = False
    while True:
        idx = content.find(name, search_start)
        if idx == -1: break
        if idx > 0:
            two_char = content[idx - 1] + name[0]
            if two_char in POSITION_WORD_ENDINGS:
                search_start = idx + 1
                continue
        found_valid = True
        break
    if not found_valid: return False

    return True


def is_only_xiao_fei_person(name: str, content: str) -> bool:
    found_shixin = False
    occurrence_count = 0
    xiaofei_count = 0
    start = 0
    while True:
        idx = content.find(name, start)
        if idx == -1: break
        occurrence_count += 1
        window = content[max(0, idx - 40): idx + len(name) + 40]
        has_shixin = any(kw in window for kw in SHIXIN_KEYWORDS)
        has_xiaofei = any(kw in window for kw in XIAOFEI_KEYWORDS)
        if has_shixin: found_shixin = True
        if has_xiaofei and not has_shixin: xiaofei_count += 1
        start = idx + 1
    if found_shixin: return False
    if occurrence_count > 0 and xiaofei_count == occurrence_count: return True
    return False


def extract_position_for_name(name: str, content: str) -> str:
    window_size = 15
    start = 0
    while True:
        idx = content.find(name, start)
        if idx == -1: break
        left = content[max(0, idx - window_size): idx]
        right = content[idx + len(name): idx + len(name) + window_size]
        for m in _POSITION_PAT.finditer(left + right):
            pos = m.group(1)
            if pos not in _INSTITUTION_WORDS: return pos
        start = idx + 1
    return ''


def _build_extraction_patterns():
    name_re = r'([\u4e00-\u9fa5]{2,4})'
    company_re = r'([\u4e00-\u9fa5]{2,30}(?:公司|集团|企业|中心|厂|合伙))'

    position_prefix = (
            r'(?:(?:公司\s*)?(?:及\s*)?'
            r'(?:' + '|'.join(re.escape(pw) for pw in POSITION_WORDS) + r')\s*)?'
    )

    field_label_vars = [
        r'失信被执行人[（(]姓名/名称[）)][：:]', r'失信被执行人姓名/名称[：:]',
        r'失信被执行人姓名[：:]', r'失信被执行人名称[：:]',
        r'失信人姓名[：:]', r'被执行人姓名[：:]',
    ]

    pat_a_name = [re.compile(rf'{lbl}\s*{name_re}', re.UNICODE) for lbl in field_label_vars]
    pat_a_long = [re.compile(rf'{lbl}\s*([^\r\n]+)', re.UNICODE) for lbl in field_label_vars]
    pat_a_company = [re.compile(rf'{lbl}\s*{company_re}', re.UNICODE) for lbl in field_label_vars]

    induct_suffix_norm = build_synonym_pattern('inducted')
    induct_suffix_spaced = build_synonym_pattern('inducted', allow_spaces=True)

    pat_b_name = re.compile(position_prefix + name_re + r'\s*' + induct_suffix_norm, re.UNICODE)
    pat_b_spaced = re.compile(position_prefix + name_re + r'\s*' + induct_suffix_spaced, re.UNICODE | re.DOTALL)
    pat_b_company = re.compile(r'(?:公司\s*)?' + company_re + r'\s*' + induct_suffix_norm, re.UNICODE)

    inducted_pat = build_synonym_pattern('inducted')
    field_pat = build_synonym_pattern('field_name')

    pat_c_induct_name = re.compile(position_prefix + name_re + r'\s*' + inducted_pat, re.UNICODE)
    pat_c_field_name = re.compile(field_pat + r'[：:]\s*' + name_re, re.UNICODE)
    pat_c_induct_company = re.compile(r'(?:公司\s*)?' + company_re + r'\s*' + inducted_pat, re.UNICODE)
    pat_c_field_company = re.compile(field_pat + r'[：:]\s*' + company_re, re.UNICODE)

    return (
        pat_a_name, pat_a_long, pat_a_company,
        pat_b_name, pat_b_spaced, pat_b_company,
        pat_c_induct_name, pat_c_field_name,
        pat_c_induct_company, pat_c_field_company,
    )


(
    _PAT_A_NAME, _PAT_A_LONG, _PAT_A_COMPANY,
    _PAT_B_NAME, _PAT_B_SPACED, _PAT_B_COMPANY,
    _PAT_C_INDUCT_NAME, _PAT_C_FIELD_NAME,
    _PAT_C_INDUCT_COMPANY, _PAT_C_FIELD_COMPANY,
) = _build_extraction_patterns()


def _extract_names_from_long_match(text: str, surnames: list[str], content: str) -> list[str]:
    for marker in CROSS_SECTION_MARKERS:
        idx = text.find(marker)
        if idx != -1: text = text[:idx]; break
    names: list[str] = []
    if is_valid_name(text.strip(), surnames, content):
        names.append(text.strip())
        return names
    if not surnames: return names
    sorted_surnames = sorted(surnames, key=len, reverse=True)
    surname_regex = '|'.join(re.escape(s) for s in sorted_surnames)
    for m in re.finditer(rf'({surname_regex}[\u4e00-\u9fa5]{{1,2}})', text):
        candidate = m.group(1)
        if is_valid_name(candidate, surnames, content): names.append(candidate)
    return names


def _extract_names_jieba(content: str, surnames: list[str]) -> list[str]:
    if not JIEBA_AVAILABLE: return []
    names: list[str] = []
    for pos_word in POSITION_WORDS:
        start = 0
        while True:
            idx = content.find(pos_word, start)
            if idx == -1: break
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
        content: str, surnames: list[str], doc_category: str
) -> tuple[list[str], list[str], list[str], dict[str, str]]:
    normalized = normalize_content(content)
    regex_candidates: set[str] = set()
    company_candidates: set[str] = set()

    def add_company(c: str):
        c = clean_company_name(c)
        if is_valid_shixin_company(c):
            company_candidates.add(c)

    for pat in _PAT_A_COMPANY:
        for m in pat.finditer(normalized): add_company(m.group(1).strip())
    for m in _PAT_B_COMPANY.finditer(normalized): add_company(m.group(1).strip())
    for m in _PAT_C_INDUCT_COMPANY.finditer(normalized): add_company(m.group(1).strip())
    for m in _PAT_C_FIELD_COMPANY.finditer(normalized): add_company(m.group(1).strip())

    extract_person = doc_category not in ('公司被纳入', '公司被撤销')

    if extract_person:
        for pat in _PAT_A_NAME:
            for m in pat.finditer(normalized):
                c = m.group(1).strip()
                if not is_company_name(c, normalized): regex_candidates.add(c)
        for pat in _PAT_A_LONG:
            for m in pat.finditer(normalized):
                raw = m.group(1).strip()
                if is_company_name(raw, normalized): continue
                if is_valid_name(raw, surnames, normalized):
                    regex_candidates.add(raw)
                else:
                    regex_candidates.update(_extract_names_from_long_match(raw, surnames, normalized))
        for m in _PAT_B_NAME.finditer(normalized):
            c = m.group(1).strip()
            if not c.startswith('司'): regex_candidates.add(c)
        for m in _PAT_B_SPACED.finditer(content):
            c = m.group(1).strip()
            if not c.startswith('司'): regex_candidates.add(c)
        for m in _PAT_C_INDUCT_NAME.finditer(normalized):
            c = m.group(1).strip()
            if not c.startswith('司'): regex_candidates.add(c)
        for m in _PAT_C_FIELD_NAME.finditer(normalized):
            regex_candidates.add(m.group(1).strip())
        if JIEBA_AVAILABLE: regex_candidates.update(_extract_names_jieba(normalized, surnames))

    regex_confirmed: set[str] = set()
    for name in regex_candidates:
        if is_valid_name(name, surnames, normalized):
            if not is_only_xiao_fei_person(name, normalized):
                regex_confirmed.add(name)

    lac_confirmed: set[str] = set()
    lac_suspicious: set[str] = set()

    if LAC_AVAILABLE and extract_person:
        shixin_sentences = _sentences_with_shixin(normalized)
        lac_names = _lac_extract_names(shixin_sentences)
        for name in lac_names:
            if name in regex_confirmed: continue
            if not is_valid_name(name, surnames, normalized): continue
            if is_only_xiao_fei_person(name, normalized): continue
            if _has_anchor(name, normalized):
                lac_confirmed.add(name)
            else:
                lac_suspicious.add(name)

    if not extract_person:
        all_confirmed = []
        suspicious = []
        name_to_position = {}
    else:
        raw_confirmed = sorted(regex_confirmed | lac_confirmed, key=len, reverse=True)
        filtered_confirmed = []
        for name in raw_confirmed:
            if not any(name in other for other in filtered_confirmed):
                filtered_confirmed.append(name)
        all_confirmed = sorted(filtered_confirmed)
        suspicious = sorted(lac_suspicious - regex_confirmed)
        name_to_position = {name: extract_position_for_name(name, normalized) for name in all_confirmed}

    valid_companies = sorted(c for c in company_candidates if c not in regex_confirmed)

    return all_confirmed, suspicious, valid_companies, name_to_position


# ══════════════════════════════════════════════════════════════════════════════
# 案号及外围逻辑
# ══════════════════════════════════════════════════════════════════════════════
_CASE_PATTERNS: list[re.Pattern] = [
    re.compile(
        r"""案号[：:]\s*([\(（]\d{4}[\)）]\s*[\u4e00-\u9fa5\d\s]+?(?:民[初终再审再]?\s*|刑[初终再]?\s*|行[初终]?\s*|执(?:恢|异|保|行)?\s*|知民[初终]?\s*|知刑[初终]?\s*|商[初终]?\s*)\s*\d+\s*号)""",
        re.VERBOSE | re.UNICODE),
    re.compile(r"""(?:案件案号|案由案号|案号)[：:]\s*([\(（]\d{4}[\)）][^，。；;、号]{5,30}?号)""", re.VERBOSE | re.UNICODE),
    re.compile(r'案号[：:]\s*(\(?\d{4}\)?[^，。；;、]{5,30}?号)', re.UNICODE),
]


def extract_case_numbers(content: str) -> list[str]:
    results: set[str] = set()
    for pat in _CASE_PATTERNS:
        for m in pat.finditer(content): results.add(re.sub(r'\s+', ' ', m.group(1).strip()))
    return list(results)


def extract_text_from_pdf(file_path: str) -> tuple[str, list[int]]:
    if not PDFPLUMBER_AVAILABLE: raise RuntimeError("pdfplumber 未安装")
    all_text: list[str] = []
    scanned_pages: list[int] = []
    with pdfplumber.open(file_path) as pdf:
        for i, page in enumerate(pdf.pages, start=1):
            text = page.extract_text()
            if text and text.strip():
                all_text.append(text)
            else:
                scanned_pages.append(i)
    return '\n'.join(all_text), scanned_pages


_DATE_PATTERNS: list[re.Pattern] = [
    re.compile(r'(\d{4})\s*年\s*(\d{1,2})\s*月\s*(\d{1,2})\s*日'),
    re.compile(r'(\d{4})-(\d{1,2})-(\d{1,2})'),
    re.compile(r'(\d{4})/(\d{1,2})/(\d{1,2})'),
    re.compile(r'(\d{4})\.(\d{1,2})\.(\d{1,2})'),
]


def extract_info_from_file(file_path: str, surnames: list[str]) -> dict:
    filename = os.path.basename(file_path)
    content, _ = extract_text_from_pdf(file_path)

    empty = {
        '证券代码': '', '证券简称': '', '公告编号': '', '公告时间': '',
        '文件分类': '', '失信被执行人姓名': '', '失信被执行人职位': '', '可疑姓名': '',
        '失信被执行人公司': '', '案号': '',
    }
    if not content.strip(): return empty
    result = dict(empty)

    m = re.search(r'证券代码[：:]\s*(\d{6})', content)
    if m: result['证券代码'] = m.group(1)
    m = re.search(r'证券简称[：:]\s*([^\s]+)', content)
    if m: result['证券简称'] = m.group(1)
    m = re.search(r'公告编号[：:]\s*([^\s]+)', content)
    if m: result['公告编号'] = m.group(1)

    lines = content.strip().split('\n')
    for line in reversed(lines[-10:]):
        for dpat in _DATE_PATTERNS:
            dm = dpat.search(line)
            if dm:
                try:
                    result['公告时间'] = f"{dm.group(1)}年{int(dm.group(2))}月{int(dm.group(3))}日"
                    break
                except ValueError:
                    continue
        if result['公告时间']: break

    doc_category = classify_document(filename, content)
    result['文件分类'] = doc_category

    case_nums = extract_case_numbers(content)
    if case_nums: result['案号'] = '，'.join(case_nums)

    confirmed, suspicious, companies, name_to_pos = extract_shixin_entities(content, surnames, doc_category)
    if confirmed:
        result['失信被执行人姓名'] = '，'.join(confirmed)
        result['失信被执行人职位'] = '，'.join([name_to_pos.get(n, '') for n in confirmed])
    if suspicious: result['可疑姓名'] = '，'.join(suspicious)
    if companies: result['失信被执行人公司'] = '，'.join(companies)

    return result


def extract_all_files(folder_path: str, surname_file_path: str) -> list[dict]:
    surnames = load_surnames(surname_file_path)
    _init_jieba(surnames)
    pdf_files = sorted(f for f in os.listdir(folder_path) if f.lower().endswith('.pdf'))
    all_results: list[dict] = []

    for filename in pdf_files:
        file_path = os.path.join(folder_path, filename)
        logger.info("\n处理: %s", filename)
        try:
            result = extract_info_from_file(file_path, surnames)
            record = {'文件名': filename, **result}
            all_results.append(record)
            logger.info("  分类: %s | 公司: %s | 姓名: %s",
                        result['文件分类'], result['失信被执行人公司'], result['失信被执行人姓名'])
        except Exception as exc:
            all_results.append({'文件名': filename, '错误': str(exc)})

    return all_results


def save_to_csv(results: list[dict], output_file: str) -> None:
    if not results: return
    fieldnames = [k for k in results[0].keys() if k != '错误']
    try:
        with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.DictWriter(f, fieldnames=fieldnames)
            writer.writeheader()
            for row in results: writer.writerow({k: v for k, v in row.items() if k != '错误'})
        logger.info("\n结果已保存: %s", output_file)
    except PermissionError:
        logger.error("[错误] 文件被占用: %s", output_file)


def main() -> None:
    folder_path = input("请输入包含PDF文件的文件夹路径: ")
    if not os.path.exists(folder_path): return

    surname_file_path = r"D:\This computer\desktop\姓氏表\姓氏表.docx"
    if not os.path.exists(surname_file_path):
        logger.error("请确保姓氏表文件存在")
        return

    results = extract_all_files(folder_path, surname_file_path)
    if results:
        output_file = os.path.join(folder_path, '失信人信息提取结果_分类版.csv')
        save_to_csv(results, output_file)
        logger.info("\n完成！共 %d 条记录", len(results))
        for cat in ['公司被纳入', '公司被撤销', '个人被纳入', '个人被撤销', '公司及个人被纳入', '公司及个人被撤销',
                    '不属于失信人', '其他']:
            count = sum(1 for r in results if r.get('文件分类') == cat)
            if count > 0: logger.info("分类[%s]: %d 条", cat, count)


if __name__ == "__main__":
    main()