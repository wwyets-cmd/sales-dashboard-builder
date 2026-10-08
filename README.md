# sales-dashboard-builder

将炬森五金（或同类）销售发货明细 Excel 一键生成为「销售数据看板」自包含 HTML 的豆包 Skill（**全国通用版**）。

输入当年 + 上年两份发货明细，输出一个**完全离线可用**（内联 ECharts）、**无密码直接访问**的交互式数据看板：
KPI 总览 → 月度同比折线 → 城市目标进度 → 品类分析（基础五金/功能五金/辅料物料三大看板）→ 客户排名（可展开月度同比/品类同比/TOP10 型号）→ 业务员卡片 → 产品型号统计 → 月度分析表，全部模块联动筛选器（年度/月份/省份/业务员/客户，选项由数据动态生成）。

## 快速开始

```bash
python3 scripts/build_dashboard.py \
    --file26 2026年数据.xls \                       # 当年数据（表头在第5行）
    --file25 2025年数据.xls \                       # 上年全年数据（表头在第1行）
    --targets assets/targets_example.json \         # 地级市目标配置（万）
    --out sales_dashboard.html
```

脚本会自动：
- 按 sheet 名识别当年/上年文件；
- 业务员按数据 `销售员` 字段归属（可用 `--province-owner` 做省份覆盖，如销售员字段错位时）；
- 执行数据清洗（剔除/合并客户、客户排名与筛选器阈值，全部参数化）；
- 品类自动分类（基础五金/功能五金/辅料三大看板）；
- 产品型号提取（OT 款开门器、K-TSM 反弹器、炬森X 去前缀/去尺寸段等规则）；
- 目标体系：地级市目标配置 → 自动计算总目标/省份目标/业务员目标，随筛选器联动；
- 校验关键指标（型号层金额守恒、客户数、辅料条数）后注入模板输出 HTML。

## 参数

| 参数 | 默认 | 说明 |
| --- | --- | --- |
| `--targets file.json` | 无 | 地级市目标(万)：`{"济南市":600}` 或 `{"济南市":{"province":"山东省","target_wan":600}}` |
| `--salesman-col` | `销售员` | 业务员字段列名 |
| `--province-owner` | 无 | 省份→业务员覆盖 JSON（销售员字段错位时用） |
| `--exclude-customers` | 无 | 剔除客户数组 |
| `--merge-customers` | 无 | 客户合并对象 |
| `--keep-min-wan` / `--filter-min-wan` | `5` / `50` | 客户排名 / 筛选器阈值（万） |
| `--title` | `炬森五金 · 销售数据看板` | 看板标题 |
| `--target` | 城市目标之和 | 年度总目标（元） |

## 目录结构

```
sales-dashboard-builder/
├── SKILL.md                  # 技能说明：触发条件、使用指南、业务规则、故障排查
├── scripts/
│   └── build_dashboard.py    # 一键生成器（读取→清洗→分类→型号→汇总→注入HTML）
├── assets/
│   ├── template.html         # 看板HTML模板（RAW为占位符，不含真实数据；内联ECharts；无密码门禁）
│   └── targets_example.json  # 地级市目标示例配置（按需替换为全国目标）
└── install.sh                # 一键安装到豆包 workspace/.user_skills
```

## 安装

```bash
git clone --depth 1 https://github.com/wwyets-cmd/sales-dashboard-builder /tmp/sdb
/tmp/sdb/install.sh            # 自动安装到 ~/.doubao/agent_mode/workspace/.user_skills
```

## 数据安全说明

- `assets/template.html` 中 `const RAW = __RAW_JSON__` 为占位符，**模板本身不含任何真实销售数据**；真实数据仅在本地生成时注入生成的 HTML 中。
- 看板为**无密码**直接访问，生成后的 HTML 若对外分发请注意数据可见范围。
