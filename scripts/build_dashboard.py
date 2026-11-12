#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_dashboard.py — 炬森五金「山东/浙江区域销售数据看板」一键生成器
====================================================================
读取销售发货明细 Excel（当年 + 上年）与固化的型号映射资产，自动生成一个
自包含、可离线打开的交互式 HTML 看板（内联 ECharts，含密码门禁、地图、
气泡图、品类/系列同比与下钻）。

核心设计：型号 → 品类/系列的归属来自已验证的资产 model_mapping.json
（由成品看板固化），脚本负责「清洗 + 物料描述解析 + 应用映射 + 汇总」，
不依赖按名称猜测的分类规则，保证口径与成品一致。

用法:
    python3 build_dashboard.py \
        --file26 2026_1-9月.xls 2026_10月.xls \
        --file25 25-26.xls \
        --targets targets.json \
        --out sales_dashboard.html

数据口径（与成品一致）:
    - 客户名 trim；青岛博丽特 → 青岛炬鑫；剔除山东鹏大、德州科硕
    - 业务员按省份强制归属：山东 → 卢力，浙江 → 刘滔，其余 → 其他
    - 金额/数量转数值，缺失填 0；退货负数正常冲减
    - 表头统一 header=4（前 4 行为"发货日期从/至"筛选条件）
    - 型号层同比 = 当年截至最大月 vs 上年同期；客户/品类另算上年全年
"""
import argparse
import json
import os
import re
import sys
from collections import defaultdict

import pandas as pd

# 可选：中文拼音排序（复刻成品 localeCompare('zh') 的客户/型号次序）。
# 未安装时回退到 Python 默认 Unicode 排序，仅极少数同金额客户的相对位置不同。
try:
    from pypinyin import lazy_pinyin

    def pinyin_key(s):
        return ''.join(lazy_pinyin(str(s)))
except ImportError:  # pragma: no cover
    def pinyin_key(s):
        return str(s)

# ============================ 业务配置（参数化） ============================
REMOVE_CUSTOMERS = {'山东鹏大新型材料科技有限公司', '德州科硕建材有限公司'}
MERGE_CUSTOMERS = {'青岛博丽特五金有限公司': '青岛炬鑫五金有限公司'}
PROVINCE_SALES = {'山东省': '卢力', '浙江省': '刘滔'}

# 12 个品类的固定展示顺序
CATEGORY_ORDER = ['拉篮', '基础五金-滑轨', '基础五金-铰链', '衣帽间五金', '升降机',
                  '柜体照明', '水槽', '辅料宣传', '水龙头', '开门器',
                  '基础五金-反弹器', '其他']

MONTHS = list(range(1, 13))

# 规范简称：循环剥离公司组织形式与行业词后缀（源数据"客户简称"字段含人名/旧称，不可靠）
SHORT_SUFFIXES = ['股份有限公司', '有限公司', '公司', '供应链管理', '电子商务',
                  '装饰工程', '装饰材料', '五金贸易', '建材', '贸易', '五金', '商贸']


def make_short(name):
    s = name
    changed = True
    while changed:
        changed = False
        for suf in SHORT_SUFFIXES:
            if s.endswith(suf) and len(s) > len(suf):
                s = s[:-len(suf)]
                changed = True
                break
    return s


# ============================ 型号解析（已用全量源数据验证） ============================
LB_RE = re.compile(r'^L-(BCG|BCW|BCP)-M01 \d+ (.*)$')
OT_RE = re.compile(r'^OT款(.*?)平推系列开门器')
HASH_RE = re.compile(r'#[^\\]*')


def extract_model(desc):
    """物料描述 → 规范型号文本。

    规则（固化，金额可 100% 对齐成品）:
      1. OT 款平推系列开门器 → OT+款式+版本（版本号移到末尾）
      2. K-TSM → 'K-TSM锁扣反弹器'
      3. L-BCG/BCW/BCP 系列 → 去 L- 前缀与尺寸段
      4. "炬森X..."（水槽/龙头）→ 去"炬森"前缀（N/Y 开头的保留）
      5. 去 #颜色/纹理 段（保留反斜杠尾缀）
      6. 逗号配件前缀（下水器/接水盘/导轨等）一律保留，不剥离
    """
    d = str(desc).strip()
    if not d:
        return d
    m = OT_RE.match(d)
    if m:
        core = m.group(1)
        mver = re.match(r'^([0-9.]+)(.*)$', core)
        if mver:
            core = mver.group(2) + mver.group(1)
        d = 'OT' + core
    if 'K-TSM' in d:
        d = 'K-TSM锁扣反弹器'
    m = LB_RE.match(d)
    if m:
        d = f'{m.group(1)}-M01 {m.group(2)}'
    if d.startswith('炬森X'):
        d = d[2:]
    d = HASH_RE.sub('', d)
    return d.strip()


class ModelResolver:
    """把物料描述解析为固化映射中的规范型号，并提供品类/系列查询。"""

    def __init__(self, model_mapping, dl_map=None):
        # model_mapping: list of {c, m, s}
        self.rows = model_mapping
        self.models = set(x['m'] for x in self.rows)
        self.model_info = {x['m']: (x['c'], x['s']) for x in self.rows}
        # 大类|小类 → 品类（兜底，用于 2025 独有/新型号）
        self.dl_map = dl_map or {}
        # 基础名（去反斜杠尾缀）→ 成品型号列表
        self.base_map = defaultdict(list)
        for m in self.models:
            self.base_map[m.split('\\')[0]].append(m)

    def fallback_category(self, dl, xl):
        return self.dl_map.get(f'{dl}|{xl}', '其他')

    def resolve(self, desc):
        """返回规范型号；无法确定时返回 None。"""
        e = extract_model(desc)
        if e in self.models:
            return e
        cands = self.base_map.get(e.split('\\')[0], [])
        if len(cands) == 1:
            return cands[0]
        # 多候选：优先尾缀完全一致
        for c in cands:
            if c == e:
                return c
        return None

    def category_series(self, model):
        return self.model_info.get(model, ('其他', model))


# ============================ 数据读取与清洗 ============================
def pick_sheet(path, prefer):
    xl = pd.ExcelFile(path)
    if len(xl.sheet_names) == 1:
        return xl.sheet_names[0]
    for s in xl.sheet_names:
        if prefer and prefer in s:
            return s
    return xl.sheet_names[0]


def load_frames(paths, prefer, header, explicit_sheet=None):
    frames = []
    for p in paths:
        sheet = explicit_sheet or pick_sheet(p, prefer)
        frames.append(pd.read_excel(p, sheet_name=sheet, header=header))
    return frames


def clean(frames, resolver):
    """清洗并型号解析，返回 (记录列表, 未解析型号金额字典)。

    能 resolve 到固化型号的记录带规范型号与系列；resolve 不到的（2025 独有/
    新型号）仍保留，用大类/小类兜底品类，型号设为 None（不进型号明细）。
    """
    records = []
    unresolved = defaultdict(float)
    for df in frames:
        for _, r in df.iterrows():
            raw_name = r.get('客户名称')
            if pd.isna(raw_name):
                continue
            name = str(raw_name).strip()
            name = MERGE_CUSTOMERS.get(name, name)
            if name in REMOVE_CUSTOMERS:
                continue
            province = str(r.get('省份')).strip() if not pd.isna(r.get('省份')) else ''
            city = str(r.get('地级市')).strip() if not pd.isna(r.get('地级市')) else ''
            sp = PROVINCE_SALES.get(province, '其他')
            qty = pd.to_numeric(pd.Series([r.get('数量')]), errors='coerce').fillna(0).iloc[0]
            amt = pd.to_numeric(pd.Series([r.get('含税本位币金额')]), errors='coerce').fillna(0).iloc[0]
            date = pd.to_datetime(r.get('发/退货日期'), errors='coerce')
            if pd.isna(date):
                continue
            desc = str(r.get('物料描述'))
            model = resolver.resolve(desc)
            if model is not None:
                cat, series = resolver.category_series(model)
            else:
                cat = resolver.fallback_category(r.get('大类名称'), r.get('小类名称'))
                series = None
                unresolved[desc] += float(amt)
            short = make_short(name)
            # 键口径品类：源小类含"开门器"的记录（柜门开门器）在月度品类字典里
            # 保留"开门器"键（金额为 0，金额实际归反弹器），复刻成品历史口径
            key_cat = '开门器' if '开门器' in str(r.get('小类名称')) else cat
            records.append({
                'name': name, 'short': short, 'province': province, 'city': city,
                'sp': sp, 'month': int(date.month), 'model': model, 'cat': cat,
                'key_cat': key_cat,
                'series': series, 'qty': float(qty), 'amt': float(amt),
                'date': date.strftime('%Y-%m-%d'), 'desc': desc,
            })
    return records, unresolved


# ============================ 汇总工具 ============================
def monthly_dict(records, value_key='amt'):
    out = {m: 0.0 for m in MONTHS}
    for r in records:
        out[r['month']] += r[value_key]
    return {str(m): round(float(v), 2) for m, v in out.items()}


def category_dict(records):
    # 键集合 = 源口径(key_cat) ∪ 最终口径(cat)；金额按最终口径，复刻成品开门器 0 键
    keys = set()
    amounts = defaultdict(float)
    for r in records:
        keys.add(r['key_cat'])
        keys.add(r['cat'])
        amounts[r['cat']] += r['amt']
    return {k: round(float(amounts.get(k, 0.0)), 2) for k in keys}


# ============================ 构建 RAW ============================
def build_raw(rec26, rec25_full, city_target_wan, total_target, region_targets, title):
    # ---- 最大月份 / 截止日期 ----
    max_month = max(r['month'] for r in rec26)
    cutoff = max(r['date'] for r in rec26)
    rec25_ytd = [r for r in rec25_full if r['month'] <= max_month]

    # ---- 客户集合（2026 ∪ 2025 并集，按 2026 金额降序，2026 无销量排末尾）----
    cust_amt26 = defaultdict(float)
    for r in rec26:
        cust_amt26[r['name']] += r['amt']
    all_names = set(cust_amt26) | set(r['name'] for r in rec25_full)
    customers = sorted(all_names,
                       key=lambda n: (-cust_amt26.get(n, 0.0), pinyin_key(n)))

    def recs_of(recs, name):
        return [r for r in recs if r['name'] == name]

    # 客户静态属性：优先 2026，2026 无记录则从 2025 取
    cust_meta = {}
    for r in rec26 + rec25_full:
        if r['name'] not in cust_meta:
            cust_meta[r['name']] = r

    # 城市目标只归属该城市 2026 销售额最大的主客户；同城其他客户目标为 0
    city_26 = defaultdict(float)
    for r in rec26:
        city_26[r['city']] += r['amt']
    city_main = {}
    for city in set(r['city'] for r in rec26):
        city_main[city] = max(
            (n for n in customers if cust_meta[n]['city'] == city),
            key=lambda n: sum(r['amt'] for r in recs_of(rec26, n)))

    customers_out = []
    for name in customers:
        meta = cust_meta[name]
        sub26 = recs_of(rec26, name)
        sub25y = recs_of(rec25_ytd, name)
        sub25f = recs_of(rec25_full, name)
        city = meta['city']
        tgt_wan = float(city_target_wan.get(city, 0)) if city_main.get(city) == name else 0.0
        customers_out.append({
            'name': name, 'short': meta['short'], 'province': meta['province'],
            'city': city, 'salesperson': meta['sp'],
            'total_2026': round(sum(r['amt'] for r in sub26), 2),
            'total_2025_ytd': round(sum(r['amt'] for r in sub25y), 2),
            'total_2025_full': round(sum(r['amt'] for r in sub25f), 2),
            'monthly_2026': monthly_dict(sub26),
            'monthly_2025_ytd': monthly_dict(sub25y),
            'monthly_2025_full': monthly_dict(sub25f),
            'category_2026': category_dict(sub26),
            'category_2025_ytd': category_dict(sub25y),
            'category_2025_full': category_dict(sub25f),
            'city_target': round(tgt_wan * 10000, 2),
            'city_target_wan': tgt_wan,
        })

    # ---- 城市汇总（目标城市清单）----
    city_prov = {}
    for r in rec26:
        city_prov.setdefault(r['city'], r['province'])
    city_sales26 = defaultdict(float)
    city_sales25y = defaultdict(float)
    city_sales25f = defaultdict(float)
    for r in rec26:
        city_sales26[r['city']] += r['amt']
    for r in rec25_ytd:
        city_sales25y[r['city']] += r['amt']
    for r in rec25_full:
        city_sales25f[r['city']] += r['amt']

    city_order = list(city_target_wan.keys())
    city_summary = []
    for city in city_order:
        tgt_wan = float(city_target_wan[city])
        tgt = tgt_wan * 10000
        a26 = city_sales26.get(city, 0.0)
        a25y = city_sales25y.get(city, 0.0)
        a25f = city_sales25f.get(city, 0.0)
        rate = a26 / tgt * 100 if tgt else 0.0
        yoy_inc = a26 - a25y
        yoy_rate = yoy_inc / a25y * 100 if a25y else 0.0
        city_summary.append({
            'city': city, 'province': city_prov.get(city, ''),
            'target': round(tgt, 2), 'target_wan': tgt_wan,
            'actual_2026': round(a26, 2), 'actual_2025_ytd': round(a25y, 2),
            'actual_2025_full': round(a25f, 2),
            'completion_rate': round(rate, 2), 'gap': round(tgt - a26, 2),
            'yoy_increment': round(yoy_inc, 2), 'yoy_rate': round(yoy_rate, 2),
        })

    # ---- 全局月度 / 品类 / 总额 ----
    global_monthly = {
        '2026': monthly_dict(rec26),
        '2025_ytd': monthly_dict(rec25_ytd),
        '2025_full': monthly_dict(rec25_full),
    }
    global_category = {
        '2026': category_dict(rec26),
        '2025_ytd': category_dict(rec25_ytd),
        '2025_full': category_dict(rec25_full),
    }
    totals = {
        '2026': round(sum(r['amt'] for r in rec26), 2),
        '2025_ytd': round(sum(r['amt'] for r in rec25_ytd), 2),
        '2025_full': round(sum(r['amt'] for r in rec25_full), 2),
    }

    # ---- 月度品类（含按业务员拆分）----
    monthly_cat = {}
    monthly_cat_sales = {}
    for m in MONTHS:
        sm = [r for r in rec26 if r['month'] == m]
        if not sm:
            continue
        # 键集合 = 源口径(key_cat) ∪ 最终口径(cat)；金额按最终口径
        keys = set()
        mc = defaultdict(float)
        key_sp = defaultdict(set)
        amt_sp = defaultdict(lambda: defaultdict(float))
        for r in sm:
            keys.add(r['key_cat'])
            keys.add(r['cat'])
            mc[r['cat']] += r['amt']
            key_sp[r['key_cat']].add(r['sp'])
            amt_sp[r['cat']][r['sp']] += r['amt']
        monthly_cat[str(m)] = {k: round(mc.get(k, 0.0), 2) for k in keys}
        monthly_cat_sales[str(m)] = {
            k: {sp: round(amt_sp[k].get(sp, 0.0), 2)
                for sp in key_sp[k] | set(amt_sp[k])}
            for k in keys}

    # ---- 客户 × 月度 × 品类 ----
    cust_monthly_cat = {}
    for name in customers:
        cust_recs = recs_of(rec26, name)
        key_sets = defaultdict(set)          # 当月键集合 = 源口径 ∪ 最终口径
        amounts = defaultdict(lambda: defaultdict(float))  # 金额按最终口径
        for r in cust_recs:
            key_sets[r['month']].add(r['key_cat'])
            key_sets[r['month']].add(r['cat'])
            amounts[r['month']][r['cat']] += r['amt']
        out = {}
        for m, ks in key_sets.items():
            out[str(m)] = {k: round(amounts[m].get(k, 0.0), 2) for k in ks}
        if out:
            cust_monthly_cat[name] = out

    # ---- 产品型号（12 品类，含数量/金额同比）----
    def model_agg(recs):
        agg = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0]))
        for r in recs:
            if r['model'] is None:
                continue  # 独有/新型号不进型号明细
            cell = agg[r['cat']][r['model']]
            cell[0] += r['qty']
            cell[1] += r['amt']
        return agg

    agg26 = model_agg(rec26)
    agg25 = model_agg(rec25_ytd)
    product_models = {}
    for cat in CATEGORY_ORDER:
        m26, m25 = agg26.get(cat, {}), agg25.get(cat, {})
        names = list(m26.keys()) + [x for x in m25 if x not in m26]
        if not names:
            continue
        rows = []
        for model in names:
            q26, a26 = m26.get(model, [0.0, 0.0])
            q25, a25 = m25.get(model, [0.0, 0.0])
            rows.append({
                'model': model,
                'qty_2026': round(q26, 4), 'amt_2026': round(a26, 2),
                'qty_2025': round(q25, 4), 'amt_2025': round(a25, 2),
            })
        rows.sort(key=lambda x: -x['amt_2026'])
        product_models[cat] = rows

    # ---- 客户 × 品类 × 型号 + TOP10 ----
    customer_product_models = {}
    customer_top10 = {}
    for name in customers:
        # 客户型号集合 = 2026 ∪ 2025 采购型号（含 2026 金额为 0 的历史型号）
        cm = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0, 0.0, 0.0]))
        for r in recs_of(rec26, name):
            if r['model'] is None:
                continue
            cell = cm[r['cat']][r['model']]
            cell[0] += r['qty']
            cell[1] += r['amt']
        for r in recs_of(rec25_ytd, name):
            if r['model'] is None:
                continue
            cell = cm[r['cat']][r['model']]  # 不存在则自动新建
            cell[2] += r['qty']
            cell[3] += r['amt']
        out = {}
        flat = []
        for cat, mdict in cm.items():
            out[cat] = []
            for model, (q26, a26, q25, a25) in mdict.items():
                item = {'model': model, 'qty_2026': round(q26, 4),
                        'amt_2026': round(a26, 2), 'qty_2025': round(q25, 4),
                        'amt_2025': round(a25, 2)}
                out[cat].append(item)
                flat.append({**item, 'category': cat})
            out[cat].sort(key=lambda x: -x['amt_2026'])
        customer_product_models[name] = out
        flat.sort(key=lambda x: (-x['amt_2026'], -x['qty_2026'], x['model']))
        customer_top10[name] = flat[:10]

    # ---- 辅料（宣传品类）明细 ----
    fuliao_detail = []
    for r in rec26:
        if r['cat'] != '辅料宣传':
            continue
        fuliao_detail.append({
            'customer': r['name'], 'short': r['short'], 'salesperson': r['sp'],
            'material': r['desc'], 'qty': round(r['qty'], 4),
            'amount': round(r['amt'], 2), 'date': r['date'],
        })

    raw = {
        'target': int(total_target),
        'max_month': int(max_month),
        'data_cutoff': cutoff,
        'city_target': {k: float(v) for k, v in city_target_wan.items()},
        'city_summary': city_summary,
        'customers': customers_out,
        'filter_customers': customers,
        'global_monthly': global_monthly,
        'global_category': global_category,
        'totals': totals,
        'monthly_cat_sales': monthly_cat_sales,
        'monthly_cat': monthly_cat,
        'cust_monthly_cat': cust_monthly_cat,
        'fuliao_detail': fuliao_detail,
        'product_models': product_models,
        'customer_product_models': customer_product_models,
        'customer_top10': customer_top10,
    }
    return raw


# ============================ 系列同比 SERIES_YOY ============================
def build_series_yoy(raw, model_mapping):
    """按系列汇总产品型号的金额/数量与型号数（2026 vs 2025 同期）。"""
    # model → 型号数据
    model_data = {}
    for cat, lst in raw['product_models'].items():
        for it in lst:
            model_data[it['model']] = it
    agg = defaultdict(lambda: {'n': 0, 'a26': 0.0, 'a25': 0.0})
    for x in model_mapping:
        d = model_data.get(x['m'])
        if d is None:
            continue
        a = agg[x['s']]
        a['n'] += 1
        a['a26'] += d['amt_2026']
        a['a25'] += d['amt_2025']
    out = []
    for s, a in agg.items():
        out.append({
            's': s, 'n': a['n'],
            'a26': round(a['a26'], 2), 'a25': round(a['a25'], 2),
        })
    return sorted(out, key=lambda x: -x['a26'])


# ============================ 渲染 HTML ============================
def render_html(raw, series_yoy, model_mapping, template_path, out_path,
                title, passwords, region_targets):
    with open(template_path, 'r', encoding='utf-8') as f:
        tpl = f.read()

    data_constants = (
        f'const RAW = {json.dumps(raw, ensure_ascii=False)};\n'
        f'const SERIES_MODEL_MAP = {json.dumps(model_mapping, ensure_ascii=False)};\n'
        f'const SERIES_YOY = {json.dumps(series_yoy, ensure_ascii=False)};\n'
        f'const REGION_TARGETS = {json.dumps(region_targets, ensure_ascii=False)};'
    )
    tpl = tpl.replace('__DATA_CONSTANTS__', data_constants, 1)

    # 密码规则
    lines = []
    for i, (pwd, scope) in enumerate(passwords.items()):
        kw = 'if' if i == 0 else 'else if'
        lines.append(f"  {kw} (pwd === '{pwd}') {{ dataScope = '{scope}'; valid = true; }}")
    tpl = tpl.replace('__PASSWORD_RULES__', '\n'.join(lines), 1)

    tpl = tpl.replace('__TITLE__', title)

    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(tpl)
    return len(tpl)


def verify(raw, series_yoy):
    t = raw['totals']
    print('=' * 64)
    print(f"数据截止 {raw['data_cutoff']}  最大月份 {raw['max_month']}")
    print(f"2026 {t['2026']/1e4:.2f}万 | 2025同期 {t['2025_ytd']/1e4:.2f}万 | "
          f"2025全年 {t['2025_full']/1e4:.2f}万")
    print(f"客户 {len(raw['customers'])} | 城市 {len(raw['city_summary'])} | "
          f"辅料明细 {len(raw['fuliao_detail'])}")
    nsku = sum(len(v) for v in raw['product_models'].values())
    s26 = sum(it['amt_2026'] for v in raw['product_models'].values() for it in v)
    print(f"SKU {nsku} | 型号层合计 {s26/1e4:.2f}万 (差异 "
          f"{(s26-t['2026'])/1e4:.4f}万)")
    print(f"系列数 {len(series_yoy)} | 系列金额合计 "
          f"{sum(d['a26'] for d in series_yoy)/1e4:.2f}万")


def main():
    ap = argparse.ArgumentParser(description='炬森五金销售数据看板一键生成器')
    ap.add_argument('--file26', nargs='+', required=True,
                    help='当年(2026)销售明细，可多个；表头 header=4')
    ap.add_argument('--file25', nargs='+', required=True,
                    help='上年(2025)销售明细，含全年；表头 header=4')
    ap.add_argument('--sheet26', default=None)
    ap.add_argument('--sheet25', default=None)
    ap.add_argument('--targets', default=None,
                    help='地级市目标 JSON（万）：{"济南市":600}')
    ap.add_argument('--target', type=int, default=32900000,
                    help='年度总目标（元），默认 32900000')
    ap.add_argument('--region-targets', default=None,
                    help='区域/业务员目标 JSON，默认山东2290/浙江1000')
    ap.add_argument('--passwords', default=None,
                    help='密码 JSON：{"4030":"all","1357":"卢力","2468":"刘滔"}')
    ap.add_argument('--title', default='炬森五金 · 山东/浙江区域销售数据看板')
    ap.add_argument('--template', default=None)
    ap.add_argument('--out', default='sales_dashboard.html')
    args = ap.parse_args()

    base = os.path.dirname(os.path.abspath(__file__))
    assets = os.path.join(base, '..', 'assets')
    template = args.template or os.path.join(assets, 'template.html')
    if not os.path.exists(template):
        sys.exit(f'模板不存在: {template}')

    model_mapping = json.load(open(os.path.join(assets, 'model_mapping.json'),
                                  encoding='utf-8'))
    dl_map = json.load(open(os.path.join(assets, 'dl_category_map.json'),
                            encoding='utf-8'))
    resolver = ModelResolver(model_mapping, dl_map)

    # ---- 读取并清洗 ----
    f26 = load_frames(args.file26, '26', 4, args.sheet26)
    rec26, un26 = clean(f26, resolver)
    f25 = load_frames(args.file25, '25', 4, args.sheet25)
    rec25, un25 = clean(f25, resolver)

    if un26:
        print('提示：当年出现无法精确匹配的物料（已用大类/小类兜底品类）:')
        for k, v in un26.items():
            print(f'  {k[:40]!r} {v:.2f}元')

    # ---- 目标配置 ----
    if args.targets:
        city_target_wan = json.load(open(args.targets, encoding='utf-8'))
        city_target_wan = {k: (float(v['target_wan']) if isinstance(v, dict)
                               else float(v)) for k, v in city_target_wan.items()}
    else:
        city_target_wan = json.load(open(os.path.join(assets, 'targets_example.json'),
                                        encoding='utf-8'))

    if args.region_targets:
        region_targets = json.load(open(args.region_targets, encoding='utf-8'))
    else:
        region_targets = {
            'all': 32900000,
            'regions': {'山东省': 22900000, '浙江省': 10000000},
            'sales': {'卢力': 22900000, '刘滔': 10000000},
        }
    passwords = (json.load(open(args.passwords, encoding='utf-8')) if args.passwords
                 else {'4030': 'all', '1357': '卢力', '2468': '刘滔'})

    raw = build_raw(rec26, rec25, city_target_wan, args.target, region_targets,
                    args.title)
    series_yoy = build_series_yoy(raw, model_mapping)
    verify(raw, series_yoy)
    size = render_html(raw, series_yoy, model_mapping, template, args.out,
                       args.title, passwords, region_targets)
    print(f'看板已生成: {args.out} ({size/1024:.0f}KB)')


if __name__ == '__main__':
    main()
