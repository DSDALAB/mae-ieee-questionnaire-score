import re
import os
import argparse
import pandas as pd
import numpy as np
from collections import OrderedDict

# 1. 讀取原始檔名（支援命令列參數 -f，或互動式輸入）
parser = argparse.ArgumentParser(description='問卷量化整理腳本')
parser.add_argument('-f', '--file', help='要處理的 Excel 檔案名稱（相對或絕對路徑）')
args = parser.parse_args()

default_name = '回收問卷_大學部應屆畢業生_ReWu.xlsx'
if args.file:
    inp = args.file.strip('"\'')
else:
    # 若沒有提供參數，提示使用者輸入（可按 Enter 使用預設檔名）
    try:
        user_inp = input(f'請輸入要處理的檔案名（直接 Enter 使用預設: {default_name}）：').strip()
    except EOFError:
        user_inp = ''
    inp = user_inp or default_name

# 使用 header=1 以正確抓到問卷的欄位名稱
# 檢查檔案是否存在，若不存在嘗試去除多餘引號後再檢查
if not os.path.exists(inp):
    alt = inp.strip('"\'')
    if alt != inp and os.path.exists(alt):
        inp = alt
    else:
        raise FileNotFoundError(f"找不到輸入檔案: {inp}")

df = pd.read_excel(inp, header=1)

# 備份一份原始資料
orig = df.copy()

# 2. 建立各種量表的文字→分數對照
map_importance = {  # 重要性
    "極重要": 5,
    "重要": 4,
    "普通": 3,
    "不重要": 2,
    "極不重要": 1,
}

map_fit = {  # 符合程度
    "非常符合": 5,
    "符合": 4,
    "普通": 3,
    "不符合": 2,
    "非常不符合": 1,
}

map_agree = {  # 認同程度
    "極認同": 5,
    "認同": 4,
    "普通": 3,
    "不認同": 2,
    "極不認同": 1,
}

map_satis = {  # 滿意度 / 品質
    "極佳": 5,
    "好": 4,
    "普通": 3,
    "差": 2,
    "極差": 1,
}

map_help = {  # 幫助程度（如果之後要用得到）
    "非常有幫助": 5,
    "有幫助": 4,
    "普通": 3,
    "沒幫助": 2,
    "沒有": 2,   # 原檔中也有「沒有」，這裡暫時視同 2 分
}

# 3. 依據欄位名稱（標題列）抓出要量化的題目欄
# 3. 依據欄位名稱（標題列）抓出要量化的題目欄
#    為了支援不同問卷版本，我們改用題號開頭的 regex 匹配（例如 ^1\.1, ^1\.2, ^2\.1 等）
cols = list(df.columns)


# 正規化欄位名稱，統一全形/半形標點與多餘空白，方便更寬鬆且精準的比對

def normalize_label(s):
    if not isinstance(s, str):
        return s
    # 統一各種點類符號為半形點
    s = s.replace('\uFF0E', '.').replace('．', '.').replace('。', '.')
    # 全形空白改為半形
    s = s.replace('\u3000', ' ')
    # 移除點前後多餘空白（例如 '8 .', '1 . 1' -> '8.' / '1.1'）
    s = re.sub(r"\s*\.\s*", '.', s)
    # 把多個空白壓縮為單一空白
    s = re.sub(r"\s+", ' ', s).strip()
    return s


# 建立正規化欄位列表，與原始欄位一一對應
norm_cols = [normalize_label(c) if isinstance(c, str) else c for c in cols]

# 根據輸入檔名判斷問卷類型（大學部 / 研究所 / 碩專班），以設定期望題號長度
base_name = os.path.basename(inp)
is_grad = ('研究所' in base_name) or ('碩專' in base_name) or ('碩專班' in base_name)

# 決定 mapping 區域名稱（對應 mapping_report 裡的區塊標題）
if '大學部' in base_name:
    region_name = '大學部'
elif '研究所' in base_name:
    region_name = '研究所'
elif '碩專' in base_name or '碩專班' in base_name:
    region_name = '碩專'
else:
    region_name = '研究所' if is_grad else '大學部'

# 內嵌的直接對應表（由 mapping_report.txt 內容嵌入，鍵為正規化後的期望題目）
mapping_table = {
    '大學部': {
        normalize_label('6. 如果您求職的面試由學校相關單位輔導安排，您認為有幫助嗎'):
            '6. 如果您求職的面試由學校相關單位輔導安排，您認為有幫助嗎',
        normalize_label('8. 目前，您認為您是否學到了專業或技能課程'):
            '8. 目前，您認為您是否學到了專業或技能課程',
        normalize_label('9. 一般的通識課程是否對您有正面幫助'):
            '9. 一般的通識課程是否對您有正面幫助',
        normalize_label('1.1 您認為具備第 1 項之核心能力對於您求學及就業的重要性'):
            '1.1 您認為具備第 1 項之核心能力對於您求學及就業的重要性',
        normalize_label('1.2 您認為具備第 1 項之核心能力對於您求學及就業的重要性'):
            '1.2 您認為具備第 1 項之核心能力對於您求學及就業的重要性',
        normalize_label('2.1 您認為具備第 2 項之核心能力對於您求學及就業的重要性'):
            '2.1 您認為具備第 2 項之核心能力對於您求學及就業的重要性',
        normalize_label('2.2 您認為求學期間本系是否已幫助你達成第 2 項之核心能力'):
            '2.2 您認為求學期間本系是否已幫助你達成第 2 項之核心能力',
        normalize_label('3.1 您認為具備第 3 項之核心能力對於您求學及就業的重要性'):
            '3.1 您認為具備第 3 項之核心能力對於您求學及就業的重要性',
        normalize_label('3.2 您認為求學期間本系是否已幫助你達成第 3 項之核心能力'):
            '3.2 您認為求學期間本系是否已幫助你達成第 3 項之核心能力',
        normalize_label('4.1 您認為具備第 4 項之核心能力對於您求學及就業的重要性'):
            '4.1 您認為具備第 4 項之核心能力對於您求學及就業的重要性',
        normalize_label('4.2 您認為求學期間本系是否已幫助你達成第 4 項之核心能力'):
            '4.2 您認為求學期間本系是否已幫助你達成第 4 項之核心能力',
        normalize_label('5.1 您認為具備第 5 項之核心能力對於您求學及就業的重要性'):
            '5.1 您認為具備第 5 項之核心能力對於您求學及就業的重要性',
        normalize_label('5.2 您認為求學期間本系是否已幫助你達成第 5 項之核心能力'):
            '5.2 您認為求學期間本系是否已幫助你達成第 5 項之核心能力',
        normalize_label('1. 您現在工作的性質與機電專業領域相關'):
            '1. 您現在工作的性質與機電專業領域相關',
        normalize_label('2. 請問您在學校所學的知識與技能在畢業後對您的工作有幫助'):
            '2. 請問您在學校所學的知識與技能在畢業後對您的工作有幫助',
        normalize_label('3. 當您畢業時，已學到基本的專業技能'):
            '3. 當您畢業時，已學到基本的專業技能',
        normalize_label('4. 一般的通識課程讓您對人生前景有所期待'):
            '4. 一般的通識課程讓您對人生前景有所期待',
        normalize_label('5. 在學期間所學已培養出獨立思考能力'):
            '5. 在學期間所學已培養出獨立思考能力',
        normalize_label('6. 如果有機會，您願意回母校與學弟妹分享學習、人生及工作上的經驗'):
            '6. 如果有機會，您願意回母校與學弟妹分享學習、人生及工作上的經驗',
        normalize_label('7. 我對於大學就讀機電工程系滿意'):
            '7. 我對於大學就讀機電工程系滿意',
        normalize_label('1. 本系專任教師教學品質'):
            '1. 本系專任教師教學品質',
        normalize_label('2. 本系兼任教師教學品質'):
            '2. 本系兼任教師教學品質',
        normalize_label('3. 本系開設之應用性課程對我的職涯有幫助'):
            '3. 本系開設之應用性課程對我的職涯有幫助',
        normalize_label('4. 本系開設之基礎課程對我的職涯有幫助'):
            '4. 本系開設之基礎課程對我的職涯有幫助',
        normalize_label('5. 本系實驗室設備與器材對我的就業有助益'):
            '5. 本系實驗室設備與器材對我的就業有助益',
        normalize_label('6. 一般通識課程對我的就業有助益'):
            '6. 一般通識課程對我的就業有助益',
        normalize_label('7. 求學期間是否善用實驗室的設備與器材'):
            '7. 求學期間是否善用實驗室的設備與器材',
        normalize_label('8 .系上教師之指導與教誨，對我的工作職能有幫助'):
            '8 .系上教師之指導與教誨，對我的工作職能有幫助',
        normalize_label('9. 您對本系的整體品質滿意'):
            '9. 您對本系的整體品質滿意'
    },
    '研究所': {
        normalize_label('6. 如果您求職的面試由學校相關單位輔導安排，您認為有幫助嗎'):
            '6. 如果您求職的面試由學校相關單位輔導安排，您認為有幫助嗎',
        normalize_label('8. 目前，您認為您是否學到了專業或技能課程'):
            '8. 目前，您認為您是否學到了專業或技能課程',
        normalize_label('9. 在學期間所參與演講或相關研討會對您有正面幫助'):
            '9. 在學期間所參與演講或相關研討會對您有正面幫助',
        normalize_label('1.1 您認為具備第 1 項之核心能力對於您求學及就業的重要性'):
            '1.1 您認為具備第 1 項之核心能力對於您求學及就業的重要性',
        normalize_label('1.2 您認為求學期間本系是否已幫助你達成第 1 項之核心能力'):
            '1.2 您認為求學期間本系是否已幫助你達成第 1 項之核心能力',
        normalize_label('2.1 您認為具備第 2 項之核心能力對於您求學及就業的重要性'):
            '2.1 您認為具備第 2 項之核心能力對於您求學及就業的重要性',
        normalize_label('2.2 您認為求學期間本系是否已幫助你達成第 2 項之核心能力'):
            '2.2 您認為求學期間本系是否已幫助你達成第 2 項之核心能力',
        normalize_label('3.1 您認為具備第 3 項之核心能力對於您求學及就業的重要性'):
            '3.1 您認為具備第 3 項之核心能力對於您求學及就業的重要性',
        normalize_label('3.2 您認為求學期間本系是否已幫助你達成第 3 項之核心能力'):
            '3.2 您認為求學期間本系是否已幫助你達成第 3 項之核心能力',
        normalize_label('1. 您現在工作的性質與機電專業領域相關'):
            '1. 您現在工作的性質與機電專業領域相關',
        normalize_label('2. 請問您在學校所學的知識與技能在畢業後對您的工作有幫助'):
            '2. 請問您在學校所學的知識與技能在畢業後對您的工作有幫助',
        normalize_label('3. 當您畢業時，是否學到了專業技能'):
            '3. 當您畢業時，是否學到了專業技能',
        normalize_label('4. 在學期間所學是否培養您的獨立思考能力'):
            '4. 在學期間所學是否培養您的獨立思考能力',
        normalize_label('5. 如果有機會，您願意回母校與學弟妹分享學習、人生及工作上的經驗'):
            '5. 如果有機會，您願意回母校與學弟妹分享學習、人生及工作上的經驗',
        normalize_label('6. 我對於研究所就讀機電工程系滿意'):
            '6. 我對於研究所就讀機電工程系滿意',
        normalize_label('1. 本系專任教師教學品質'):
            '1. 本系專任教師教學品質',
        normalize_label('2. 本系兼任教師教學品質'):
            '2. 本系兼任教師教學品質',
        normalize_label('3. 本系開設課程之實用性對我的職涯有幫助'):
            '3. 本系開設課程之實用性對我的職涯有幫助',
        normalize_label('4. 本系開設課程之品質對我的職涯有幫助'):
            '4. 本系開設課程之品質對我的職涯有幫助',
        normalize_label('5. 本系實驗室的設備與器材對我的就業有助益'):
            '5. 本系實驗室的設備與器材對我的就業有助益',
        normalize_label('6. 求學期間是否善用實驗室的設備與器材'):
            '6. 求學期間是否善用實驗室的設備與器材',
        normalize_label('7. 系上教師之指導與教誨，對我的工作職能有幫助'):
            '7. 系上教師之指導與教誨，對我的工作職能有幫助',
        normalize_label('8. 您對本系的整體品質滿意'):
            '8. 您對本系的整體品質滿意'
    },
    '碩專': {
        normalize_label('6. 如果您求職的面試由學校相關單位輔導安排，您認為有幫助嗎'):
            '6. 如果您求職的面試由學校相關單位輔導安排，您認為有幫助嗎',
        normalize_label('8. 目前，您認為您是否學到了專業或技能課程'):
            '8. 目前，您認為您是否學到了專業或技能課程',
        normalize_label('9. 在學期間所參與演講或相關研討會對您有正面幫助'):
            '9. 在學期間所參與演講或相關研討會對您有正面幫助',
        normalize_label('1.1 您認為具備第 1 項之核心能力對於您求學及就業的重要性'):
            '1.1 您認為具備第 1 項之核心能力對於您求學及就業的重要性',
        normalize_label('1.2 您認為求學期間本系是否已幫助你達成第 1 項之核心能力'):
            '1.2 您認為求學期間本系是否已幫助你達成第 1 項之核心能力',
        normalize_label('2.1 您認為具備第 2 項之核心能力對於您求學及就業的重要性'):
            '2.1 您認為具備第 2 項之核心能力對於您求學及就業的重要性',
        normalize_label('2.2 您認為求學期間本系是否已幫助你達成第 2 項之核心能力'):
            '2.2 您認為求學期間本系是否已幫助你達成第 2 項之核心能力',
        normalize_label('3.1 您認為具備第 3 項之核心能力對於您求學及就業的重要性'):
            '3.1 您認為具備第 3 項之核心能力對於您求學及就業的重要性',
        normalize_label('3.2 您認為求學期間本系是否已幫助你達成第 3 項之核心能力'):
            '3.2 您認為求學期間本系是否已幫助你達成第 3 項之核心能力',
        normalize_label('1. 您現在工作的性質與機電專業領域相關'):
            '1. 您現在工作的性質與機電專業領域相關',
        normalize_label('2. 請問您在學校所學的知識與技能在畢業後對您的工作有幫助'):
            '2. 請問您在學校所學的知識與技能在畢業後對您的工作有幫助',
        normalize_label('3. 當您畢業時，是否學到了專業技能'):
            '3. 當您畢業時，是否學到了專業技能',
        normalize_label('4. 在學期間所學是否培養您的獨立思考能力'):
            '4. 在學期間所學是否培養您的獨立思考能力',
        normalize_label('5. 如果有機會，您願意回母校與學弟妹分享學習、人生及工作上的經驗'):
            '5. 如果有機會，您願意回母校與學弟妹分享學習、人生及工作上的經驗',
        normalize_label('6. 我對於研究所就讀機電工程系滿意'):
            '6. 我對於研究所就讀機電工程系滿意',
        normalize_label('1. 本系專任教師教學品質'):
            '1. 本系專任教師教學品質',
        normalize_label('2. 本系兼任教師教學品質'):
            '2. 本系兼任教師教學品質',
        normalize_label('3. 本系開設課程之實用性對我的職涯有幫助'):
            '3. 本系開設課程之實用性對我的職涯有幫助',
        normalize_label('4. 本系開設課程之品質對我的職涯有幫助'):
            '4. 本系開設課程之品質對我的職涯有幫助',
        normalize_label('5. 本系實驗室的設備與器材對我的就業有助益'):
            '5. 本系實驗室的設備與器材對我的就業有助益',
        normalize_label('6. 求學期間是否善用實驗室的設備與器材'):
            '6. 求學期間是否善用實驗室的設備與器材',
        normalize_label('7. 系上教師之指導與教誨，對我的工作職能有幫助'):
            '7. 系上教師之指導與教誨，對我的工作職能有幫助',
        normalize_label('8. 您對本系的整體品質滿意'):
            '8. 您對本系的整體品質滿意'
    }
}


def get_mapped_col(prefix):
    """在 mapping_table[region_name] 中尋找第一個 key 以 prefix 開頭的對應欄位，回傳原始欄名或 None。"""
    pref = normalize_label(prefix)
    region_map = mapping_table.get(region_name, {})
    for k, v in region_map.items():
        if k.startswith(pref):
            return v
    return None


def find_col_by_prefix(prefix, used=set()):
    """以正規化後的欄位名尋找第一個以 prefix 開頭的欄位（允許空白或全形點差異）。
    回傳原始欄位名稱並將其加入 used。
    """
    # 正規化 prefix（把可能的空白與全形點處理一致）
    prefix_norm = normalize_label(prefix)
    pat = re.compile(r'^' + re.escape(prefix_norm))
    for orig, norm in zip(cols, norm_cols):
        if orig in used:
            continue
        if isinstance(norm, str) and pat.search(norm):
            used.add(orig)
            return orig
    return None


def fallback_find_by_number(main, sub, used):
    """備援：在正規化後的欄位名中做寬鬆存在性匹配（例如 '4.1', '4.', '第4項', '4項' 等）。"""
    targets = [f'{main}.{sub}', f'{main}.', f'第{main}項', f'{main}項']
    for orig, norm in zip(cols, norm_cols):
        if orig in used:
            continue
        if not isinstance(norm, str):
            continue
        for t in targets:
            if t in norm:
                used.add(orig)
                return orig
    return None


# 以題號為主的期望列表（會嘗試找對應的欄位名稱）
used_cols = set()
# 先偵測「個人基本資料與綜合問題」常見題號（例如 6., 8., 9.），放在前面避免被其他搜尋佔用
personal_prefixes = ['6.', '8.', '9.']
personal_cols = []
for p in personal_prefixes:
    col = get_mapped_col(p)
    if col and col not in used_cols:
        personal_cols.append(col)
        used_cols.add(col)
if is_grad:
    core_imp_prefixes = ['1.1', '2.1', '3.1']
    core_fit_prefixes = ['1.2', '2.2', '3.2']
    edu_prefixes = ['1.', '2.', '3.', '4.', '5.', '6.']
    teach_prefixes = ['1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.']
else:
    core_imp_prefixes = ['1.1', '2.1', '3.1', '4.1', '5.1']
    core_fit_prefixes = ['1.2', '2.2', '3.2', '4.2', '5.2']
    edu_prefixes = ['1.', '2.', '3.', '4.', '5.', '6.', '7.']
    teach_prefixes = ['1.', '2.', '3.', '4.', '5.', '6.', '7.', '8.', '9.']

# 使用精確 prefix 匹配，找不到時使用備援的數字匹配
core_imp_cols = []
for p in core_imp_prefixes:
    col = get_mapped_col(p)
    if col and col not in used_cols:
        core_imp_cols.append(col)
        used_cols.add(col)

core_fit_cols = []
for p in core_fit_prefixes:
    col = get_mapped_col(p)
    if col and col not in used_cols:
        core_fit_cols.append(col)
        used_cols.add(col)


def find_section_start(keywords):
    for idx, norm in enumerate(norm_cols):
        if not isinstance(norm, str):
            continue
        for kw in keywords:
            if kw in norm:
                return idx
    return None


def find_col_in_range(prefix, start_idx, used):
    prefix_norm = normalize_label(prefix)
    pat = re.compile(r'^' + re.escape(prefix_norm))
    for orig, norm in zip(cols[start_idx+1:], norm_cols[start_idx+1:]):
        if orig in used:
            continue
        if isinstance(norm, str) and pat.search(norm):
            used.add(orig)
            return orig
    return None


# 先嘗試依節段（找到「參、本系教育成效調查」或相似文字）再在該節後面尋找題號
edu_start = find_section_start(['教育成效調查', '本系教育成效', '參、本系教育成效'])
teach_start = find_section_start(['教學及設備', '教學及設備滿意度', '本系教學及設備'])

edu_cols = []
for p in edu_prefixes:
    col = None
    if edu_start is not None:
        col = find_col_in_range(p, edu_start, used_cols)
    if not col:
        col = find_col_by_prefix(p, used_cols)
    if not col:
        main = p.rstrip('.')
        col = fallback_find_by_number(main, '1', used_cols)
    if col:
        edu_cols.append(col)
    col = get_mapped_col(p)
    if col and col not in used_cols:
        edu_cols.append(col)
        used_cols.add(col)

teach_cols = []
for p in teach_prefixes:
    col = None
    if teach_start is not None:
        col = find_col_in_range(p, teach_start, used_cols)
    if not col:
        col = find_col_by_prefix(p, used_cols)
    if not col:
        main = p.rstrip('.')
        col = fallback_find_by_number(main, '1', used_cols)
    if col:
        teach_cols.append(col)
    col = get_mapped_col(p)
    if col and col not in used_cols:
        teach_cols.append(col)
        used_cols.add(col)

# 根據問卷類型，設定每個構面應該有的題數（用於保證摘要列數）
if is_grad:
    expected_counts = {'personal': 3, 'core_imp': 3,
                       'core_fit': 3, 'edu': 6, 'teach': 8}
else:
    expected_counts = {'personal': 3, 'core_imp': 5,
                       'core_fit': 5, 'edu': 7, 'teach': 9}

# 過濾明顯的個資/欄位標題避免誤抓入教育成效
meta_keywords = ['姓名', '聯絡', 'E-Mail', '電話', '地址', '性別', '出生', '學號']


def filter_meta(cols_list):
    cleaned = [c for c in cols_list if not any(
        mk in str(c) for mk in meta_keywords)]
    return cleaned if cleaned else cols_list


edu_cols = filter_meta(edu_cols)
teach_cols = filter_meta(teach_cols)

# 裁切或補齊至預期題數（若缺題則加入預留占位字串，方便人工對照）


def normalize_list(lst, key):
    want = expected_counts.get(key, len(lst))
    if len(lst) > want:
        return lst[:want]
    elif len(lst) < want:
        # 補上 placeholder
        out = list(lst)
        for i in range(len(lst)+1, want+1):
            out.append(f'未對應_{key}_{i}')
        return out
    return lst


personal_cols = normalize_list(personal_cols, 'personal')
core_imp_cols = normalize_list(core_imp_cols, 'core_imp')
core_fit_cols = normalize_list(core_fit_cols, 'core_fit')
edu_cols = normalize_list(edu_cols, 'edu')
teach_cols = normalize_list(teach_cols, 'teach')

# 移除研究所/碩專問卷中不必要或亂入的題目（例如含有「研究所畢業後」之類的雜項）


def is_noise_question(label):
    if not isinstance(label, str):
        return False
    nl = normalize_label(label)
    if '研究所畢業後' in nl or '研究所畢業' in nl:
        return True
    return False


edu_cols = [c for c in edu_cols if not is_noise_question(c)]
teach_cols = [c for c in teach_cols if not is_noise_question(c)]
core_imp_cols = [c for c in core_imp_cols if not is_noise_question(c)]
core_fit_cols = [c for c in core_fit_cols if not is_noise_question(c)]
personal_cols = [c for c in personal_cols if not is_noise_question(c)]

# 4. 建立一份含數值的新 DataFrame
num_df = df.copy()

# 核心能力：重要性 → map_importance
for c in core_imp_cols:
    if c in num_df.columns:
        num_df[c + '_num'] = num_df[c].map(map_importance)
    else:
        num_df[c + '_num'] = np.nan

# 核心能力：達成度 → map_fit
for c in core_fit_cols:
    if c in num_df.columns:
        num_df[c + '_num'] = num_df[c].map(map_fit)
    else:
        num_df[c + '_num'] = np.nan

# 教育成效 + 教學設備：有些題目是認同量表、有些比較像滿意度
# 先試 map_agree，沒對到再試 map_satis
for c in edu_cols + teach_cols:
    if c in num_df.columns:
        num_df[c +
               '_num'] = num_df[c].map(map_agree).fillna(num_df[c].map(map_satis))
    else:
        num_df[c + '_num'] = np.nan

# 個人基本題目：題號 6 使用 map_help，題號 8/9 使用 map_satis
for idx, prefix in enumerate(['6.', '8.', '9.']):
    try:
        c = personal_cols[idx]
    except Exception:
        c = None
    if not c:
        continue
    if c in num_df.columns:
        if prefix == '6.':
            num_df[c + '_num'] = num_df[c].map(map_help)
        else:
            num_df[c + '_num'] = num_df[c].map(map_satis)
    else:
        num_df[c + '_num'] = np.nan

# 5. 計算每位受測者的構面平均分數


def row_mean(cols):
    """對一組欄位做列平均，如果 list 為空就全部給 NaN。"""
    if not cols:
        return pd.Series(np.nan, index=num_df.index)
    return num_df[cols].mean(axis=1)


# helper: 將 1-based 欄位編號轉成 Excel 欄字母 (1 -> A, 27 -> AA)
def col_idx_to_excel_col(idx):
    letters = ''
    while idx > 0:
        idx, rem = divmod(idx - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# 對應各構面的題目（num 欄）
core_imp_num = [c + '_num' for c in core_imp_cols]
core_fit_num = [c + '_num' for c in core_fit_cols]
edu_prof_num = [edu_cols[0] + '_num', edu_cols[1] + '_num',
                edu_cols[2] + '_num'] if len(edu_cols) >= 3 else []
edu_gen_num = [edu_cols[3] + '_num', edu_cols[4] +
               '_num'] if len(edu_cols) >= 5 else []
edu_alumni_num = [edu_cols[5] + '_num', edu_cols[6] +
                  '_num'] if len(edu_cols) >= 7 else []
teach_quality_num = [teach_cols[0] + '_num', teach_cols[1] +
                     '_num', teach_cols[7] + '_num'] if len(teach_cols) >= 8 else []
teach_course_num = [teach_cols[2] + '_num', teach_cols[3] +
                    '_num', teach_cols[5] + '_num'] if len(teach_cols) >= 6 else []
teach_equipment_num = [teach_cols[4] + '_num', teach_cols[6] +
                       '_num', teach_cols[8] + '_num'] if len(teach_cols) >= 9 else []

# 寫入構面分數欄
num_df['核心能力_重要性'] = row_mean(core_imp_num)
num_df['核心能力_達成度'] = row_mean(core_fit_num)
num_df['教育成效_專業連結'] = row_mean(edu_prof_num)
num_df['教育成效_通識思辨'] = row_mean(edu_gen_num)
num_df['教育成效_校友認同'] = row_mean(edu_alumni_num)
num_df['教學品質'] = row_mean(teach_quality_num)
num_df['課程實用性'] = row_mean(teach_course_num)
num_df['設備整體滿意'] = row_mean(teach_equipment_num)

# 總量表平均：把所有有定義的 num 題一起平均
num_cols_all = (
    core_imp_num + core_fit_num +
    edu_prof_num + edu_gen_num + edu_alumni_num +
    teach_quality_num + teach_course_num + teach_equipment_num
)
num_df['總量表平均'] = row_mean(num_cols_all)

# 從「資料建立日期」欄提取年度資訊
try:
    date_col_candidates = [c for c in orig.columns if isinstance(c, str) and '資料建立日期' in c]
    if date_col_candidates:
        date_series = orig[date_col_candidates[0]]
    else:
        raise KeyError('找不到「資料建立日期」欄位')
    date_parsed = pd.to_datetime(date_series, errors='coerce')
    num_df['_年度'] = date_parsed.dt.year.astype('Int64')
    unique_years = sorted([int(y) for y in num_df['_年度'].dropna().unique()])
except Exception as e:
    print(f'警告：無法取得年度資訊（{e}），將不產生年度分類工作表。')
    num_df['_年度'] = pd.NA
    unique_years = []

# 根據問卷類型調整構面標籤：研究所 / 碩專 使用三個研究導向構面
if is_grad:
    comp_labels = [
        '具備精進專業知能與研究方法之能力',
        '具備規劃及執行專題研究與成果撰述之能力',
        '具備蒐尋與評述國際文獻之能力'
    ]
else:
    comp_labels = [
        '具備數理與工程知能',
        '具備設計製造與自動化專業知能',
        '具備整合與實踐工程實務之能力',
        '具備溝通表達與團隊合作之能力',
        '具備體察環境與持續成長之能力'
    ]


# 6. 題目級統計摘要：以函式封裝，支援全資料與年度篩選
def generate_stats(data_df):
    """根據指定的資料框（可為年度篩選後的子集）產生 (summary_df, comp_df)。"""
    local_rows = []

    def _append_header(title):
        row = OrderedDict()
        row['題目'] = title
        for col in ['有效樣本數', '平均數', '標準差', '最小值', '最大值', '4分以上比例']:
            row[col] = np.nan
        local_rows.append(row)

    def _append_questions(cols_list):
        for c in cols_list:
            if isinstance(c, str) and c.startswith('未對應_'):
                continue
            num_col = c + '_num'
            row = OrderedDict()
            row['題目'] = c
            if num_col in data_df.columns:
                s = data_df[num_col].dropna()
                if s.empty:
                    row['有效樣本數'] = np.nan
                    row['平均數'] = np.nan
                    row['標準差'] = np.nan
                    row['最小值'] = np.nan
                    row['最大值'] = np.nan
                    row['4分以上比例'] = np.nan
                else:
                    row['有效樣本數'] = int(s.count())
                    row['平均數'] = float(s.mean())
                    row['標準差'] = float(s.std(ddof=1)) if s.count() > 1 else np.nan
                    row['最小值'] = float(s.min())
                    row['最大值'] = float(s.max())
                    row['4分以上比例'] = float((s >= 4).mean())
            else:
                row['有效樣本數'] = np.nan
                row['平均數'] = np.nan
                row['標準差'] = np.nan
                row['最小值'] = np.nan
                row['最大值'] = np.nan
                row['4分以上比例'] = np.nan
            score_labels = ['極重要', '重要', '普通', '不認同', '極不認同']
            score_vals   = [5, 4, 3, 2, 1]
            try:
                if num_col in data_df.columns and not data_df[num_col].dropna().empty:
                    s = data_df[num_col].dropna()
                    counts = s.value_counts()
                    total = float(s.count())
                    for lbl, val in zip(score_labels, score_vals):
                        row[f'{lbl}數量'] = int(counts.get(val, 0))
                    for lbl, val in zip(score_labels, score_vals):
                        row[f'{lbl}比例'] = float(counts.get(val, 0)) / total if total > 0 else np.nan
                else:
                    for lbl in score_labels:
                        row[f'{lbl}數量'] = np.nan
                    for lbl in score_labels:
                        row[f'{lbl}比例'] = np.nan
            except Exception:
                for lbl in score_labels:
                    row[f'{lbl}數量'] = np.nan
                for lbl in score_labels:
                    row[f'{lbl}比例'] = np.nan
            local_rows.append(row)

    _append_header('壹、個人基本資料與綜合問題')
    _append_questions(personal_cols)

    _append_header('貳、核心能力')
    max_pairs = max(len(core_imp_cols), len(core_fit_cols))
    for i in range(max_pairs):
        if i < len(core_imp_cols):
            c_imp = core_imp_cols[i]
            if not (isinstance(c_imp, str) and c_imp.startswith('未對應_')):
                _append_questions([c_imp])
        if i < len(core_fit_cols):
            c_fit = core_fit_cols[i]
            if not (isinstance(c_fit, str) and c_fit.startswith('未對應_')):
                _append_questions([c_fit])

    _append_header('參、本系教學成效調查')
    _append_questions(edu_cols)

    _append_header('肆、本系教學及設備滿意度調查表')
    _append_questions(teach_cols)

    s_df = pd.DataFrame(local_rows)

    # 構面彙整
    c_rows = []
    for idx, lbl in enumerate(comp_labels):
        imp_col = core_imp_num[idx] if idx < len(core_imp_num) else None
        fit_col = core_fit_num[idx] if idx < len(core_fit_num) else None

        if imp_col and imp_col in data_df.columns and fit_col and fit_col in data_df.columns:
            series = data_df[[imp_col, fit_col]].mean(axis=1).dropna()
        elif imp_col and imp_col in data_df.columns:
            series = data_df[imp_col].dropna()
        elif fit_col and fit_col in data_df.columns:
            series = data_df[fit_col].dropna()
        else:
            series = pd.Series(dtype=float)

        row = OrderedDict()
        row['構面'] = lbl
        if series.empty:
            row['有效樣本數'] = np.nan
            row['平均數'] = np.nan
            row['標準差'] = np.nan
            row['最小值'] = np.nan
            row['最大值'] = np.nan
            row['4分以上比例'] = np.nan
        else:
            row['有效樣本數'] = int(series.count())
            row['平均數'] = float(series.mean())
            row['標準差'] = float(series.std(ddof=1)) if series.count() > 1 else np.nan
            row['最小值'] = float(series.min())
            row['最大值'] = float(series.max())
            row['4分以上比例'] = float((series >= 4).mean())
        # 五級分布（對平均後的分數四捨五入後計算）
        score_labels = ['極重要', '重要', '普通', '不認同', '極不認同']
        score_vals   = [5, 4, 3, 2, 1]
        if not series.empty:
            rounded = series.round().astype(int)
            counts = rounded.value_counts()
            total = float(len(rounded))
            for lbl, val in zip(score_labels, score_vals):
                row[f'{lbl}數量'] = int(counts.get(val, 0))
            for lbl, val in zip(score_labels, score_vals):
                row[f'{lbl}比例'] = float(counts.get(val, 0)) / total if total > 0 else np.nan
        else:
            for lbl in score_labels:
                row[f'{lbl}數量'] = np.nan
            for lbl in score_labels:
                row[f'{lbl}比例'] = np.nan
        c_rows.append(row)

    c_df = pd.DataFrame(c_rows)
    return s_df, c_df


def generate_comp_year_table(all_df, years):
    """產生核心能力 × 年度 比例彙整表（百分比）。"""
    score_labels = ['極重要', '重要', '普通', '不認同', '極不認同']
    score_vals   = [5, 4, 3, 2, 1]

    def _comp_series(data_df, idx):
        imp_col = core_imp_num[idx] if idx < len(core_imp_num) else None
        fit_col = core_fit_num[idx] if idx < len(core_fit_num) else None
        if imp_col and imp_col in data_df.columns and fit_col and fit_col in data_df.columns:
            return data_df[[imp_col, fit_col]].mean(axis=1).dropna()
        elif imp_col and imp_col in data_df.columns:
            return data_df[imp_col].dropna()
        elif fit_col and fit_col in data_df.columns:
            return data_df[fit_col].dropna()
        return pd.Series(dtype=float)

    rows = []
    for idx, lbl in enumerate(comp_labels):
        # 各年度列
        for yr in years:
            roc_yr = yr - 1911
            yr_df = all_df[all_df['_年度'] == yr]
            s = _comp_series(yr_df, idx)
            row = OrderedDict()
            row['核心能力'] = lbl
            row['年度'] = f'{roc_yr}'
            if s.empty:
                for sl in score_labels:
                    row[sl] = np.nan
            else:
                rounded = s.round().astype(int)
                counts = rounded.value_counts()
                total = float(len(rounded))
                for sl, sv in zip(score_labels, score_vals):
                    row[sl] = (f"{round(counts.get(sv, 0) / total * 100, 1)} %" if total > 0 else np.nan)
            rows.append(row)
        # 合計列
        s_all = _comp_series(all_df, idx)
        row = OrderedDict()
        row['核心能力'] = lbl
        row['年度'] = '合計'
        if s_all.empty:
            for sl in score_labels:
                row[sl] = np.nan
        else:
            rounded = s_all.round().astype(int)
            counts = rounded.value_counts()
            total = float(len(rounded))
            for sl, sv in zip(score_labels, score_vals):
                row[sl] = (f"{round(counts.get(sv, 0) / total * 100, 1)} %" if total > 0 else np.nan)
        rows.append(row)
        # 空白分隔列
        rows.append(OrderedDict([('核心能力', ''), ('年度', '')] + [(sl, np.nan) for sl in score_labels]))

    return pd.DataFrame(rows)


def generate_comp_weight_table(all_df, years):
    """產生核心能力權重表：各年度平均分數與標準差（百分制，1-5分×20），含總分列。"""
    n = len(comp_labels)
    weight_pct = f'{round(100 / n)}%' if n > 0 else '—'

    def _comp_series(data_df, idx):
        imp_col = core_imp_num[idx] if idx < len(core_imp_num) else None
        fit_col = core_fit_num[idx] if idx < len(core_fit_num) else None
        if imp_col and imp_col in data_df.columns and fit_col and fit_col in data_df.columns:
            return data_df[[imp_col, fit_col]].mean(axis=1).dropna()
        elif imp_col and imp_col in data_df.columns:
            return data_df[imp_col].dropna()
        elif fit_col and fit_col in data_df.columns:
            return data_df[fit_col].dropna()
        return pd.Series(dtype=float)

    rows = []
    # 各核心能力 × 各年度
    for idx, lbl in enumerate(comp_labels):
        for yr in years:
            roc_yr = yr - 1911
            s = _comp_series(all_df[all_df['_年度'] == yr], idx)
            row = OrderedDict()
            row['核心能力'] = lbl
            row['核心能力權重'] = weight_pct
            row['年度'] = str(roc_yr)
            if s.empty:
                row['平均分數'] = np.nan
                row['標準差'] = np.nan
            else:
                row['平均分數'] = round(float(s.mean()) * 20, 1)
                row['標準差'] = round(float(s.std(ddof=1)) * 20, 1) if s.count() > 1 else np.nan
            rows.append(row)
        # 合計列
        s_all = _comp_series(all_df, idx)
        row = OrderedDict()
        row['核心能力'] = lbl
        row['核心能力權重'] = weight_pct
        row['年度'] = '合計'
        if s_all.empty:
            row['平均分數'] = np.nan
            row['標準差'] = np.nan
        else:
            row['平均分數'] = round(float(s_all.mean()) * 20, 1)
            row['標準差'] = round(float(s_all.std(ddof=1)) * 20, 1) if s_all.count() > 1 else np.nan
        rows.append(row)

    # 總分列：各年度加權平均（等權重→直接平均所有構面）
    for yr in years:
        roc_yr = yr - 1911
        yr_df = all_df[all_df['_年度'] == yr]
        all_series_list = [_comp_series(yr_df, i) for i in range(n)]
        valid = [s for s in all_series_list if not s.empty]
        row = OrderedDict()
        row['核心能力'] = '總分'
        row['核心能力權重'] = ''
        row['年度'] = str(roc_yr)
        if valid:
            combined = pd.concat(valid)
            row['平均分數'] = round(float(combined.mean()) * 20, 1)
            row['標準差'] = round(float(combined.std(ddof=1)) * 20, 1) if combined.count() > 1 else np.nan
        else:
            row['平均分數'] = np.nan
            row['標準差'] = np.nan
        rows.append(row)
    # 總分合計
    all_series_list = [_comp_series(all_df, i) for i in range(n)]
    valid = [s for s in all_series_list if not s.empty]
    row = OrderedDict()
    row['核心能力'] = '總分'
    row['核心能力權重'] = ''
    row['年度'] = '合計'
    if valid:
        combined = pd.concat(valid)
        row['平均分數'] = round(float(combined.mean()) * 20, 1)
        row['標準差'] = round(float(combined.std(ddof=1)) * 20, 1) if combined.count() > 1 else np.nan
    else:
        row['平均分數'] = np.nan
        row['標準差'] = np.nan
    rows.append(row)

    return pd.DataFrame(rows)


# 7. 匯出成新的 Excel 檔（含多個工作表）
base_name = os.path.splitext(os.path.basename(inp))[0]
out_path = f"{base_name}_Result.xlsx"

# 產生全資料統計
summary_df, comp_df = generate_stats(num_df)
comp_year_df = generate_comp_year_table(num_df, unique_years)
comp_weight_df = generate_comp_weight_table(num_df, unique_years)

with pd.ExcelWriter(out_path, engine='xlsxwriter') as writer:
    orig.to_excel(writer, sheet_name='原始資料', index=False)
    num_df.to_excel(writer, sheet_name='個案含分數', index=False)
    summary_df.to_excel(writer, sheet_name='題目統計摘要', index=False)
    comp_df.to_excel(writer, sheet_name='核心能力', index=False)
    comp_year_df.to_excel(writer, sheet_name='核心能力_年度比例', index=False)
    comp_weight_df.to_excel(writer, sheet_name='核心能力_權重表', index=False)
    # 依年度輸出各年度的統計摘要與核心能力
    for yr in unique_years:
        yr_df = num_df[num_df['_年度'] == yr]
        yr_summary, yr_comp = generate_stats(yr_df)
        yr_summary.to_excel(writer, sheet_name=f'題目統計_{yr}'[:31], index=False)
        yr_comp.to_excel(writer, sheet_name=f'核心能力_{yr}'[:31], index=False)

print(f'已產生：{out_path}')
if unique_years:
    print(f'已依年度分類：{unique_years}')
else:
    print('警告：BA 欄無法解析年度資訊，未產生年度分類工作表。')
