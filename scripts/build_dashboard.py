#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
build_dashboard.py — 炬森五金「鲁浙销售数据看板」生成器
====================================================
输入两份原始销售发货明细（.xls/.xlsx），输出一个自包含的交互式 HTML 看板。

用法:
    python3 build_dashboard.py \
        --file26 2026数据文件.xls   # 当年数据: 前4行为筛选条件, 表头在第5行(header=4)
        --file25 2025数据文件.xls   # 上年数据: 表头在第1行(header=0), 包含全年1-12月
        [--out sales_dashboard.html]      # 输出文件, 默认 sales_dashboard.html
        [--target 31900000]               # 年度总目标(元), 默认3190万
        [--template template.html]        # HTML模板, 默认找脚本同目录template.html
        [--skip-25]                       # 跳过上年数据(不生成同比)

数据口径(固化自业务规则):
    - 业务归属: 山东省→卢力, 浙江省→刘滔(原始销售员一律按省份覆盖)
    - 剔除客户: 山东鹏大新型材料科技有限公司, 德州科硕建材有限公司
    - 客户合并: 青岛博丽特五金有限公司 → 青岛炬鑫五金有限公司
    - 客户排名: 删除当年销售额<5万的客户; 筛选器只保留当年≥50万的客户
    - 品类: 铰链/滑轨/反弹器/拉篮/水槽/水龙头/柜体照明/开门器/升降机/衣帽间五金/辅料宣传
    - 功能五金 = 除基础五金(铰链/滑轨/反弹器)和辅料外的全部品类
    - 同比口径: 当年截至最大月份 vs 上年同期(截至同月); 上年全年单独统计
"""
import argparse
import json
import os
import re
import sys
from collections import defaultdict

import pandas as pd

# ============================ 业务配置 ============================
REMOVE_CUSTOMERS = {'山东鹏大新型材料科技有限公司', '德州科硕建材有限公司'}
MERGE_CUSTOMERS = {'青岛博丽特五金有限公司': '青岛炬鑫五金有限公司'}
# 16城市年度目标(万元)
CITY_TARGET_WAN = {
    '济南市': 600, '青岛市': 500, '淄博市': 180, '东营市': 100, '烟台市': 260,
    '潍坊市': 100, '泰安市': 50, '临沂市': 450, '德州市': 450, '聊城市': 100,
    '菏泽市': 100, '杭州市': 200, '宁波市': 100, '温州市': 200, '金华市': 180, '台州市': 350,
}
SALES_TARGET = {'卢力': 23000000, '刘滔': 10000000}
KEEP_MIN_WAN = 5       # 客户排名: 当年销售额低于该值(万)的客户删除
FILTER_MIN_WAN = 50    # 筛选器: 仅当年销售额≥该值(万)的客户进入

CATEGORY_ORDER = ['基础五金-铰链', '基础五金-滑轨', '基础五金-反弹器', '拉篮', '水槽',
                  '水龙头', '柜体照明', '开门器', '升降机', '衣帽间五金', '辅料宣传', '其他']


def normalize_province(p):
    return '山东省' if '山东' in str(p) else '浙江省'


def salesperson_of(province):
    return '卢力' if normalize_province(province) == '山东省' else '刘滔'


# ============================ 品类分类 ============================
def classify_category(dl, xl):
    """按大类名称/小类名称分类到11个品类。规则优先级不可随意调整。"""
    dl = str(dl) if not pd.isna(dl) else ''
    xl = str(xl) if not pd.isna(xl) else ''
    if '辅料' in dl:
        return '辅料宣传'
    if '衣柜' in dl:
        return '衣帽间五金'
    if xl == '抽屉系列':           # 抽屉系列(含OT款抽屉开门器)一律归拉篮
        return '拉篮'
    if '拉篮' in dl:               # 零部件_拉篮(主包材类_纸箱等)归拉篮
        return '拉篮'
    if '铰链' in dl or '铰链' in xl:
        return '基础五金-铰链'
    if '滑轨' in dl or '滑轨' in xl:
        return '基础五金-滑轨'
    if '反弹' in xl:
        return '基础五金-反弹器'
    if '升降机' in dl:
        return '升降机'
    if '开门器' in dl or '上翻门' in dl or '电动推拉' in dl or xl == '开门器':
        return '开门器'
    if '照明' in xl or '智能控制' in dl:
        return '柜体照明'
    if xl == '拉篮':
        return '拉篮'
    if xl == '水槽':
        return '水槽'
    if xl == '水龙头':
        return '水龙头'
    if xl in ('衣杆', '衣柜五金'):
        return '衣帽间五金'
    if xl == '宣传类':
        return '辅料宣传'
    if xl == '平推系列':           # OT款平推开门器
        return '开门器'
    return '其他'


# ============================ 型号提取 ============================
# 定制映射: 少数系列/描述需要归一化为用户认可的关键型号名
MODEL_OVERRIDES = {
    '成品配件同步杆': '配件同步杆',
    '成品配件DS-A辅件包': '配件DS-A辅件包',
}
# 数据归纳的配件前缀: 逗号前的配件名剥离(仅限出现在X产品行前的固定配件前缀)
STRIP_X_PREFIXES = ('110台控下水器,', '140下水器,', '140台控下水器,', '185下水器,', '90台控下水器,', '170下水器,')
# L-B系列: 'L-BCG-M01 800 mini款抽屉系列锅具板篮' → 'BCG-M01 mini款抽屉系列锅具板篮'(去尺寸段)
LB_RE = re.compile(r'^L-(BCG|BCW|BCP)-M01 \d+ (.*)$')
OT_RE = re.compile(r'^OT款(.*?)平推系列开门器')
KTSM_RE = re.compile(r'K-TSM')


def extract_model(desc):
    """从物料描述提取产品型号名(中文关键信息)。

    规则(固化):
      1. OT款平推系列开门器 → OT+款式+版本(版本号移到末尾, 如 'OT款1.5小角度加重型平推系列开门器'→'OT小角度加重型1.5')
      2. K-TSM → 'K-TSM锁扣反弹器'
      3. L-BCG/L-BCW/L-BCP 系列 → 去掉 L- 前缀
      4. 炬森X开头(水槽/水龙头) → 去掉'炬森'前缀 与 '#颜色' 后缀
      5. 少量定制映射(成品配件同步杆等)
      6. 其余: 物料描述原样(含 '产品退市,已废弃' 等尾缀)
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
        return 'OT' + core
    if KTSM_RE.search(d):
        return 'K-TSM锁扣反弹器'
    m = LB_RE.match(d)
    if m:
        return f'{m.group(1)}-M01 {m.group(2)}'
    # 水槽/水龙头配件: 剥离固定配件前缀
    for p in STRIP_X_PREFIXES:
        if d.startswith(p):
            d = d[len(p):]
            break
    if d.startswith('炬森X'):
        d = d[2:]
        if '#' in d:
            d = d[:d.index('#')]
    if '#' in d:
        d = d[:d.index('#')]
    d = MODEL_OVERRIDES.get(d, d)
    return d


# ============================ 数据读取 ============================
def read_data(file26, file25=None):
    def pick_sheet(path, want26):
        xl = pd.ExcelFile(path)
        if len(xl.sheet_names) == 1:
            return xl.sheet_names[0]
        for s in xl.sheet_names:
            if want26 and ('26' in s or '2_6' in s):
                return s
            if (not want26) and '25' in s:
                return s
        return xl.sheet_names[0]

    df26 = pd.read_excel(file26, sheet_name=pick_sheet(file26, True), header=4)
    df25 = pd.read_excel(file25, sheet_name=pick_sheet(file25, False), header=0) if file25 else None
    return df26, df25


def clean_df(df, year_tag):
    """清洗: 剔除客户、合并客户、业务员按省份覆盖。返回新DataFrame副本。"""
    out = df.copy()
    out['__cust'] = out['客户名称'].apply(
        lambda c: None if (pd.isna(c) or c in REMOVE_CUSTOMERS) else MERGE_CUSTOMERS.get(c, c))
    out = out[out['__cust'].notna()]
    out['__sp'] = out['省份'].apply(salesperson_of)
    out['__date'] = pd.to_datetime(out['发/退货日期'], errors='coerce')
    out['__month'] = out['__date'].dt.month
    return out


# ============================ 汇总 ============================
def build_raw(df26, df25, target, max_month_override=None, data_cutoff_override=None):
    """由清洗后的当年/上年明细生成看板RAW JSON。"""
    # ---- 当年(2026) ----
    cur = clean_df(df26, '26')
    max_month = int(cur['__month'].max()) if max_month_override is None else max_month_override
    cutoff = pd.to_datetime(cur['__date'].max()).strftime('%Y-%m-%d') if data_cutoff_override is None else data_cutoff_override

    # ---- 上年(2025): 同期(1..max_month) 与 全年 ----
    prev_full = clean_df(df25, '25') if df25 is not None else None
    if prev_full is not None:
        prev_ytd = prev_full[prev_full['__month'] <= max_month]
    else:
        prev_ytd = None

    # ---- 客户列表(当年排序) ----
    cust_sales26 = cur.groupby('__cust')['含税本位币金额'].sum()
    keep = cust_sales26[cust_sales26 >= KEEP_MIN_WAN * 10000].sort_values(ascending=False)
    customers = list(keep.index)

    def filter_cust(df_, custs):
        return df_[df_['__cust'].isin(custs)] if df_ is not None else None

    cur = filter_cust(cur, customers)
    prev_full = filter_cust(prev_full, customers)
    prev_ytd = filter_cust(prev_ytd, customers)

    # ---- 品类字段 ----
    for df_ in (cur, prev_full, prev_ytd):
        if df_ is not None and '__cat' not in df_.columns:
            df_['__cat'] = df_.apply(lambda r: classify_category(r['大类名称'], r['小类名称']), axis=1)
    if prev_full is not None and '__cat' not in prev_full.columns:
        prev_full['__cat'] = prev_full.apply(lambda r: classify_category(r['大类名称'], r['小类名称']), axis=1)
    if prev_ytd is not None and '__cat' not in prev_ytd.columns:
        prev_ytd['__cat'] = prev_ytd.apply(lambda r: classify_category(r['大类名称'], r['小类名称']), axis=1)

    months = [str(m) for m in range(1, 13)]

    def monthly_sum(df_):
        if df_ is None or len(df_) == 0:
            return {m: 0.0 for m in months}
        s = df_.groupby('__month')['含税本位币金额'].sum()
        return {str(m): float(s.get(m, 0.0)) for m in range(1, 13)}

    global_monthly = {
        '2026': monthly_sum(cur),
        '2025_ytd': monthly_sum(prev_ytd),
        '2025_full': monthly_sum(prev_full),
    }

    def category_sum(df_):
        if df_ is None or len(df_) == 0:
            return {}
        s = df_.groupby('__cat')['含税本位币金额'].sum()
        # 保留全部出现过的品类键(含金额为0的退货抵消情形), 与业务看板结构一致
        return {k: float(v) for k, v in s.items()}

    global_category = {
        '2026': category_sum(cur),
        '2025_ytd': category_sum(prev_ytd),
        '2025_full': category_sum(prev_full),
    }

    def total_of(d):
        return float(sum(d.values()))

    totals = {
        '2026': float(cur['含税本位币金额'].sum()),
        '2025_ytd': float(prev_ytd['含税本位币金额'].sum()) if prev_ytd is not None and len(prev_ytd) else 0.0,
        '2025_full': float(prev_full['含税本位币金额'].sum()) if prev_full is not None and len(prev_full) else 0.0,
    }

    # ---- 城市汇总 ----
    city_summary = []
    for city, wan in CITY_TARGET_WAN.items():
        province = '山东省' if city in ('济南市', '青岛市', '淄博市', '东营市', '烟台市', '潍坊市',
                                        '泰安市', '临沂市', '德州市', '聊城市', '菏泽市') else '浙江省'
        tgt = wan * 10000
        a26 = float(cur[cur['地级市'] == city]['含税本位币金额'].sum())
        a25y = float(prev_ytd[prev_ytd['地级市'] == city]['含税本位币金额'].sum()) if prev_ytd is not None and len(prev_ytd) else 0.0
        a25f = float(prev_full[prev_full['地级市'] == city]['含税本位币金额'].sum()) if prev_full is not None and len(prev_full) else 0.0
        rate = a26 / tgt * 100 if tgt else 0.0
        yoy_inc = a26 - a25y
        yoy_rate = yoy_inc / a25y * 100 if a25y else 0.0
        city_summary.append({
            'city': city, 'province': province, 'target': tgt, 'target_wan': wan,
            'actual_2026': round(a26, 2), 'actual_2025_ytd': round(a25y, 2), 'actual_2025_full': round(a25f, 2),
            'completion_rate': round(rate, 2), 'gap': round(tgt - a26, 2),
            'yoy_increment': round(yoy_inc, 2), 'yoy_rate': round(yoy_rate, 2),
        })

    # ---- 客户明细 ----
    city_of = cur.drop_duplicates('__cust').set_index('__cust')['地级市'].to_dict()
    customers_out = []
    for i, c in enumerate(customers, 1):
        short = cur[cur['__cust'] == c]['客户简称']
        short = short.iloc[0] if len(short) and not pd.isna(short.iloc[0]) else c[:10]
        city = city_of.get(c, '')
        prov = '山东省' if normalize_province(cur[cur['__cust'] == c]['省份'].iloc[0]) == '山东省' else '浙江省'
        sub = cur[cur['__cust'] == c]
        sub25y = prev_ytd[prev_ytd['__cust'] == c] if prev_ytd is not None else None
        sub25f = prev_full[prev_full['__cust'] == c] if prev_full is not None else None
        tgt_wan = CITY_TARGET_WAN.get(city, 0)
        customers_out.append({
            'name': c, 'short': str(short), 'province': prov, 'city': city,
            'salesperson': salesperson_of(prov),
            'total_2026': round(float(sub['含税本位币金额'].sum()), 2),
            'total_2025_ytd': round(float(sub25y['含税本位币金额'].sum()), 2) if sub25y is not None else 0.0,
            'total_2025_full': round(float(sub25f['含税本位币金额'].sum()), 2) if sub25f is not None else 0.0,
            'monthly_2026': monthly_sum(sub),
            'monthly_2025_ytd': monthly_sum(sub25y),
            'monthly_2025_full': monthly_sum(sub25f),
            'category_2026': category_sum(sub),
            'category_2025_ytd': category_sum(sub25y),
            'category_2025_full': category_sum(sub25f),
            'city_target': tgt_wan * 10000, 'city_target_wan': float(tgt_wan),
        })

    filter_customers = [c for c in customers if cust_sales26[c] >= FILTER_MIN_WAN * 10000]

    # ---- 月度×品类 ----
    monthly_cat = {}
    monthly_cat_sales = {}
    for m in range(1, 13):
        sm = cur[cur['__month'] == m]
        if len(sm):
            mc = {}
            ms = {}
            for (cat, sp), v in sm.groupby(['__cat', '__sp'])['含税本位币金额'].sum().items():
                mc[cat] = mc.get(cat, 0.0) + float(v)
                ms.setdefault(cat, {})[sp] = float(v)
            monthly_cat[str(m)] = mc
            monthly_cat_sales[str(m)] = ms

    # ---- 客户×月度×品类 ----
    cust_monthly_cat = {}
    for c in customers:
        cmc = {}
        sub = cur[cur['__cust'] == c]
        for m in range(1, 13):
            sm = sub[sub['__month'] == m]
            if len(sm):
                cs = sm.groupby('__cat')['含税本位币金额'].sum()
                cmc[str(m)] = {k: float(v) for k, v in cs.items() if abs(v) > 0.005}
        if cmc:
            cust_monthly_cat[c] = cmc

    # ---- 产品型号 ----
    def model_agg(df_, prefix):
        if df_ is None or len(df_) == 0:
            return {}
        df_ = df_.copy()
        df_['__model'] = df_['物料描述'].apply(extract_model)
        agg = {}
        for (cat, model), g in df_.groupby(['__cat', '__model']):
            agg.setdefault(cat, []).append({
                'model': model,
                'qty_' + prefix: float(g['数量'].sum()),
                'amt_' + prefix: float(g['含税本位币金额'].sum()),
            })
        return agg

    models26 = model_agg(cur, '2026')
    models25 = model_agg(prev_ytd, '2025')  # 型号层只与同期比较

    product_models = {}
    for cat in CATEGORY_ORDER:
        m26 = {m['model']: m for m in models26.get(cat, [])}
        m25 = {m['model']: m for m in models25.get(cat, [])}
        merged = []
        for model in list(m26) + [x for x in m25 if x not in m26]:
            merged.append({
                'model': model,
                'qty_2026': m26.get(model, {}).get('qty_2026', 0.0),
                'amt_2026': m26.get(model, {}).get('amt_2026', 0.0),
                'qty_2025': m25.get(model, {}).get('qty_2025', 0.0),
                'amt_2025': m25.get(model, {}).get('amt_2025', 0.0),
            })
        if merged:
            product_models[cat] = merged

    # ---- 客户×品类×型号 + TOP10 ----
    customer_product_models = {}
    customer_top10 = {}
    for c in customers:
        sub = cur[cur['__cust'] == c]
        sub25 = prev_ytd[prev_ytd['__cust'] == c] if prev_ytd is not None else None
        sub = sub.copy(); sub['__model'] = sub['物料描述'].apply(extract_model)
        cm = {}
        for (cat, model), g in sub.groupby(['__cat', '__model']):
            cm.setdefault(cat, []).append({
                'model': model,
                'qty_2026': float(g['数量'].sum()),
                'amt_2026': float(g['含税本位币金额'].sum()),
                'qty_2025': 0.0, 'amt_2025': 0.0,
            })
        if sub25 is not None:
            sub25 = sub25.copy(); sub25['__model'] = sub25['物料描述'].apply(extract_model)
            for (cat, model), g in sub25.groupby(['__cat', '__model']):
                for it in cm.get(cat, []):
                    if it['model'] == model:
                        it['qty_2025'] = float(g['数量'].sum())
                        it['amt_2025'] = float(g['含税本位币金额'].sum())
                        break
        customer_product_models[c] = cm
        top = []
        for cat, items in cm.items():
            for it in items:
                top.append({**it, 'category': cat})
        top.sort(key=lambda x: (-x['amt_2026'], -x['qty_2026'], x['model']))
        customer_top10[c] = top[:10]

    # ---- 辅料(物料)明细 ----
    fuliao_detail = []
    fl = cur[cur['__cat'] == '辅料宣传']
    for _, r in fl.iterrows():
        fuliao_detail.append({
            'customer': r['__cust'], 'short': str(r['客户简称']) if not pd.isna(r['客户简称']) else str(r['__cust'])[:10],
            'salesperson': r['__sp'], 'material': str(r['物料描述']),
            'qty': float(r['数量']), 'amount': float(r['含税本位币金额']),
            'date': r['__date'].strftime('%Y-%m-%d'),
        })

    raw = {
        'target': int(target),
        'max_month': int(max_month),
        'data_cutoff': cutoff,
        'city_target': CITY_TARGET_WAN,
        'city_summary': city_summary,
        'customers': customers_out,
        'filter_customers': filter_customers,
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


def render_html(raw, template_path, out_path):
    with open(template_path, 'r', encoding='utf-8') as f:
        template = f.read()
    if '__RAW_JSON__' not in template:
        raise RuntimeError('模板中未找到 __RAW_JSON__ 占位符')
    payload = json.dumps(raw, ensure_ascii=False)
    html = template.replace('__RAW_JSON__', payload, 1)
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(html)
    return len(html)


def verify(raw):
    """自校验: 打印关键指标, 供人工核对。"""
    t = raw['totals']
    print('=' * 60)
    print(f"数据截止: {raw['data_cutoff']}  最大月份: {raw['max_month']}")
    print(f"当年总额: {t['2026']/10000:.2f}万  上年同期: {t['2025_ytd']/10000:.2f}万  上年全年: {t['2025_full']/10000:.2f}万")
    print(f"同比增量: {(t['2026']-t['2025_ytd'])/10000:.2f}万")
    print(f"客户数: {len(raw['customers'])}  筛选器客户: {len(raw['filter_customers'])}")
    nc = sum(len(v) for v in raw['product_models'].values())
    print(f"产品型号数: {nc}  辅料明细: {len(raw['fuliao_detail'])}")
    g = raw['global_category']['2026']
    for k in sorted(g, key=lambda x: -g[x]):
        print(f"  品类 {k}: {g[k]/10000:.2f}万")
    # 型号层金额应等于客户总额
    s26 = sum(it['amt_2026'] for v in raw['product_models'].values() for it in v)
    print(f"型号层当年合计: {s26/10000:.2f}万 (差异 {(s26-t['2026'])/10000:.2f}万)")


def main():
    ap = argparse.ArgumentParser(description='炬森五金鲁浙销售数据看板生成器')
    ap.add_argument('--file26', required=True, help='当年(2026)销售明细, 前4行筛选条件, 表头在第5行')
    ap.add_argument('--file25', help='上年(2025)销售明细, 表头在第1行, 含全年数据')
    ap.add_argument('--skip-25', action='store_true', help='跳过上年数据(不生成同比)')
    ap.add_argument('--out', default='sales_dashboard.html')
    ap.add_argument('--target', type=int, default=31900000, help='年度总目标(元)')
    ap.add_argument('--template', default=None, help='HTML模板路径(默认脚本同目录template.html)')
    args = ap.parse_args()

    base = os.path.dirname(os.path.abspath(__file__))
    if args.template:
        template = args.template
    else:
        candidates = [os.path.join(base, 'template.html'),
                      os.path.join(base, '..', 'assets', 'template.html')]
        template = next((p for p in candidates if os.path.exists(p)), candidates[0])
    if not os.path.exists(template):
        sys.exit(f'模板不存在: {template}')

    df26, df25 = read_data(args.file26, args.file25 if not args.skip_25 else None)
    raw = build_raw(df26, df25, args.target)
    verify(raw)
    size = render_html(raw, template, args.out)
    print(f'看板已生成: {args.out} ({size/1024:.0f}KB)')


if __name__ == '__main__':
    main()
