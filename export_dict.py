# -*- coding: utf-8 -*-
"""
字段字典导出工具
================================================================
把数据字典（Excel / CSV / MySQL 的 information_schema）统一导出成
一份 dict.json，供后续做「一句话查数」的字段检索。

同时输出一份质量报告，指出哪些字段的注释没法用来做检索。

用法：
    1. 改下面「配置区」里的几行
    2. 双击或在终端运行： python export_dict.py
    3. 同目录下得到 dict.json 和 dict_quality.md

依赖：
    Excel/CSV 模式：openpyxl（你已有）
    MySQL   模式：pymysql（pip install pymysql）
================================================================
"""

import csv
import json
import os
import re
import sys
from collections import Counter, defaultdict
from datetime import datetime

# Windows 控制台中文输出
if sys.platform == "win32":
    try:
        sys.stdout.reconfigure(encoding="utf-8")
    except Exception:
        pass


# ================================================================
# 配置区 —— 只需要改这里
# ================================================================

数据源类型 = "excel"          # "excel" / "csv" / "mysql"

字典文件 = "字典.xlsx"         # excel / csv 模式用这个
输出文件 = "dict.json"
质量报告文件 = "dict_quality.md"

数据库配置 = {
    "host": "127.0.0.1",
    "port": 3306,
    "user": "readonly",               # 强烈建议只读账号
    "password": "",
    "database": "your_database",
    "charset": "utf8mb4",
}

# ================================================================


# ---------- 列名智能识别：不同人的字典表列名写法五花八门 ----------
候选列名 = {
    "表名":   ["表名", "表英文名", "表名称", "table", "table_name", "tablename"],
    "表注释": ["表注释", "表中文名", "表说明", "表描述", "table_comment"],
    "字段名": ["字段名", "字段", "字段英文名", "字段名称", "column", "column_name", "field", "field_name"],
    "类型":   ["类型", "数据类型", "字段类型", "type", "data_type", "字段数据类型"],
    "注释":   ["注释", "中文名", "字段注释", "说明", "描述", "comment", "字段说明"],
}


def 猜列名(表头):
    """把真实表头映射到标准字段名。先精确匹配，再包含匹配。"""
    映射 = {}
    已用小写 = {str(h).strip().lower(): i for i, h in enumerate(表头) if h is not None}

    for 标准名, 候选们 in 候选列名.items():
        for 候选 in 候选们:
            if 候选.lower() in 已用小写:
                映射[标准名] = 已用小写[候选.lower()]
                break
        else:
            # 精确没中，退化为包含匹配
            for 候选 in 候选们:
                for 小写名, 下标 in 已用小写.items():
                    if 候选.lower() in 小写名:
                        映射[标准名] = 下标
                        break
                if 标准名 in 映射:
                    break
    return 映射


# ---------- 字段类型归类：决定这个字段能不能 SUM、能不能按时间筛 ----------
def 归类类型(类型字符串):
    t = str(类型字符串 or "").lower()
    if re.search(r"int|decimal|numeric|float|double|real|money", t):
        return "number"
    if re.search(r"date|time|timestamp", t):
        return "date"
    if re.search(r"bool|bit", t):
        return "bool"
    return "text"


# ---------- 三种数据源的读取 ----------
def 从表格读取(路径):
    """读 Excel。两种排版都支持：
       A. 一个 sheet 装全部字段，靠「表名」列区分
       B. 每个 sheet 一张表，sheet 名就是表名
    """
    from openpyxl import load_workbook

    工作簿 = load_workbook(路径, data_only=True)
    结果 = []

    for 工作表 in 工作簿.worksheets:
        行们 = list(工作表.iter_rows(values_only=True))
        if not 行们:
            continue

        # 找表头行：前 5 行里第一行非空单元格 >= 2 的
        表头下标 = 0
        for i, 行 in enumerate(行们[:5]):
            非空 = [c for c in 行 if c not in (None, "")]
            if len(非空) >= 2:
                表头下标 = i
                break

        表头 = 行们[表头下标]
        列映射 = 猜列名(表头)

        if "字段名" not in 列映射:
            continue

        有表名列 = "表名" in 列映射
        当前表名 = None if 有表名列 else 工作表.title

        for 行 in 行们[表头下标 + 1:]:
            if not 行 or all(c in (None, "") for c in 行):
                continue

            def 取(键):
                idx = 列映射.get(键)
                if idx is None or idx >= len(行):
                    return ""
                return str(行[idx]).strip() if 行[idx] is not None else ""

            字段名 = 取("字段名")
            if not 字段名:
                continue

            表名 = 取("表名") if 有表名列 else 当前表名
            if not 表名:
                表名 = "未命名表"

            结果.append({
                "表名": 表名,
                "表注释": 取("表注释"),
                "字段名": 字段名,
                "类型": 取("类型"),
                "注释": 取("注释"),
            })

    return 结果


def 从CSV读取(路径):
    结果 = []
    with open(路径, "r", encoding="utf-8-sig", newline="") as f:
        读取器 = csv.reader(f)
        行们 = list(读取器)
    if not 行们:
        return 结果

    表头 = 行们[0]
    列映射 = 猜列名(表头)
    if "字段名" not in 列映射:
        raise ValueError("CSV 里没找到「字段名」这一列，检查一下表头")

    有表名列 = "表名" in 列映射
    当前表名 = os.path.splitext(os.path.basename(路径))[0]

    for 行 in 行们[1:]:
        if not 行 or all(c.strip() == "" for c in 行):
            continue

        def 取(键):
            idx = 列映射.get(键)
            if idx is None or idx >= len(行):
                return ""
            return 行[idx].strip()

        字段名 = 取("字段名")
        if not 字段名:
            continue

        结果.append({
            "表名": 取("表名") if 有表名列 else 当前表名,
            "表注释": 取("表注释"),
            "字段名": 字段名,
            "类型": 取("类型"),
            "注释": 取("注释"),
        })

    return 结果


def 从数据库读取(配置):
    try:
        import pymysql
    except ImportError:
        print("缺少 pymysql，先装：pip install pymysql")
        sys.exit(1)

    SQL = """
        SELECT c.TABLE_NAME,
               t.TABLE_COMMENT,
               c.COLUMN_NAME,
               c.COLUMN_TYPE,
               c.COLUMN_COMMENT
        FROM information_schema.COLUMNS c
        JOIN information_schema.TABLES t
          ON t.TABLE_SCHEMA = c.TABLE_SCHEMA AND t.TABLE_NAME = c.TABLE_NAME
        WHERE c.TABLE_SCHEMA = %s
        ORDER BY c.TABLE_NAME, c.ORDINAL_POSITION
    """

    连接 = pymysql.connect(**配置)
    try:
        with 连接.cursor() as 游标:
            游标.execute(SQL, (配置["database"],))
            行们 = 游标.fetchall()
    finally:
        连接.close()

    return [{
        "表名": r[0] or "",
        "表注释": r[1] or "",
        "字段名": r[2] or "",
        "类型": r[3] or "",
        "注释": r[4] or "",
    } for r in 行们]


# ---------- 质量检查：判断这条注释能不能给检索用 ----------
占位注释 = {"", "-", "—", "/", "无", "暂无", "空", "待补充", "待定", "todo",
            "null", "none", "n/a", "0", "1", "字段", "字段名"}

枚举线索_字段名 = ("status", "state", "type", "kind", "flag", "is_", "level", "channel")
枚举线索_注释 = ("状态", "类型", "是否", "等级", "渠道", "来源")


def 检查质量(全部字段):
    # 同表内重复注释：这才是真问题。跨表重复（create_time 都叫创建时间）属正常
    表内注释计数 = defaultdict(Counter)
    for f in 全部字段:
        表内注释计数[f["表名"]][f["注释"]] += 1

    问题 = []
    for f in 全部字段:
        注释 = (f["注释"] or "").strip()
        问题类型 = []

        if 注释 == "":
            问题类型.append("无注释")
        elif 注释.lower() in 占位注释:
            问题类型.append("占位符")
        else:
            # 这两条互不排斥：一句"金额"既太短，又和同表别的字段撞了
            if len(注释) <= 2:
                问题类型.append("注释过短")
            if f["表名"] and 表内注释计数[f["表名"]][注释] >= 2:
                问题类型.append("同表内重复")

        字段名小写 = f["字段名"].lower()
        像枚举 = (any(k in 字段名小写 for k in 枚举线索_字段名)
                  or any(k in 注释 for k in 枚举线索_注释))
        if 像枚举 and "枚举无取值说明" not in 问题类型:
            问题类型.append("枚举无取值说明")

        f["像枚举"] = 像枚举
        f["问题"] = 问题类型
        if 问题类型:
            问题.append(f)

    return 问题


# ---------- 组装输出 ----------
def 组装(全部字段):
    按表 = defaultdict(lambda: {"表名": "", "表注释": "", "字段": []})

    for f in 全部字段:
        表 = 按表[f["表名"]]
        表["表名"] = f["表名"]
        if f["表注释"]:
            表["表注释"] = f["表注释"]

        注释 = (f["注释"] or "").strip()
        字段小写 = f["字段名"].lower()

        # 检索文本：把能帮模型理解这个字段的语料全拼进去
        检索文本 = " ".join(x for x in [
            表["表注释"], f["表名"], f["字段名"], 注释, f["类型"],
        ] if x)

        表["字段"].append({
            "字段名": f["字段名"],
            "类型": f["类型"],
            "归类": 归类类型(f["类型"]),
            "注释": 注释,
            "检索文本": 检索文本,
            "问题": f["问题"],
            # 给 NL2SQL 用的两个提示：能不能聚合、是不是状态字段
            "可聚合": 归类类型(f["类型"]) == "number" and not any(
                k in 字段小写 for k in ("id", "code", "no")),
            "像枚举": f.get("像枚举", False),
        })

    return {
        "生成时间": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "数据源类型": 数据源类型,
        "表数": len(按表),
        "字段数": len(全部字段),
        "表": sorted(按表.values(), key=lambda x: x["表名"]),
    }


def 写质量报告(数据, 问题列表, 路径):
    总数 = 数据["字段数"]
    有注释 = sum(1 for t in 数据["表"] for f in t["字段"]
               if f["注释"] and f["注释"].lower() not in 占位注释)
    可用 = sum(1 for t in 数据["表"] for f in t["字段"]
               if f["注释"] and "问题" in f and not f["问题"])

    分类计数 = Counter()
    for f in 问题列表:
        for p in f["问题"]:
            分类计数[p] += 1

    行 = []
    行.append("# 字段字典质量报告\n")
    行.append(f"生成时间：{数据['生成时间']}\n")
    行.append("## 总览\n")
    行.append("| 指标 | 数值 |")
    行.append("|---|---|")
    行.append(f"| 表数 | {数据['表数']} |")
    行.append(f"| 字段总数 | {总数} |")
    行.append(f"| 有注释 | {有注释}（{有注释 / 总数 * 100:.1f}%） |" if 总数 else "| 有注释 | 0 |")
    行.append(f"| **可直接用于检索** | **{总数 - len(问题列表)}（{(总数 - len(问题列表)) / 总数 * 100:.1f}%）** |" if 总数 else "")
    行.append(f"| 有问题 | {len(问题列表)} |")
    行.append("")

    行.append("## 问题分布\n")
    行.append("| 问题类型 | 数量 | 影响 |")
    行.append("|---|---|---|")
    说明 = {
        "无注释": "检索完全找不到，模型只能靠字段名硬猜",
        "占位符": "等同于无注释，但更容易让人误以为有",
        "注释过短": "歧义大，多个字段会同时命中",
        "同表内重复": "同表内多个字段注释一样，检索必选错",
        "枚举无取值说明": "模型不知道有哪些值，会编造筛选条件",
    }
    for 类型, 数量 in 分类计数.most_common():
        行.append(f"| {类型} | {数量} | {说明.get(类型, '')} |")
    行.append("")

    行.append("## 优先要改的 30 个字段\n")
    行.append("| 表 | 字段 | 类型 | 现有注释 | 问题 |")
    行.append("|---|---|---|---|---|")

    # 排序：问题越多越靠前，同等情况下优先枚举字段（危害最大）
    def 权重(f):
        return (-len(f["问题"]), 0 if "枚举无取值说明" in f["问题"] else 1, f["表名"])

    for f in sorted(问题列表, key=权重)[:30]:
        注释 = f["注释"] or "（空）"
        行.append(f"| {f['表名']} | `{f['字段名']}` | {f['类型']} | {注释} | {'、'.join(f['问题'])} |")

    行.append("")
    行.append("---\n")
    行.append("**怎么用这份报告**：先改「枚举无取值说明」和「同表内重复」——这两类会直接让检索选错字段，")
    行.append("而且改法是去库里跑 `SELECT DISTINCT 字段` 把真实取值补进注释，工作量小、收益最大。")
    行.append("「无注释」的字段数量通常最多，按使用频率从高到低补，不用一次补完。")

    with open(路径, "w", encoding="utf-8") as f:
        f.write("\n".join(行))


def main():
    print("=" * 56)
    print("字段字典导出")
    print("=" * 56)

    if 数据源类型 == "mysql":
        print(f"数据源：MySQL / {数据库配置['database']}")
        全部字段 = 从数据库读取(数据库配置)
    elif 数据源类型 == "csv":
        print(f"数据源：CSV / {字典文件}")
        全部字段 = 从CSV读取(字典文件)
    else:
        print(f"数据源：Excel / {字典文件}")
        全部字段 = 从表格读取(字典文件)

    if not 全部字段:
        print("\n没读到任何字段。检查一下：列名里有没有「字段名」这一列？")
        return

    print(f"读到 {len(全部字段)} 个字段")

    问题列表 = 检查质量(全部字段)
    数据 = 组装(全部字段)

    with open(输出文件, "w", encoding="utf-8") as f:
        json.dump(数据, f, ensure_ascii=False, indent=2)

    写质量报告(数据, 问题列表, 质量报告文件)

    print(f"表数：{数据['表数']}")
    可直接用 = len(全部字段) - len(问题列表)
    print(f"可直接用于检索：{可直接用} / {len(全部字段)} "
          f"（{可直接用 / len(全部字段) * 100:.1f}%）")

    if 问题列表:
        分类 = Counter()
        for f in 问题列表:
            for p in f["问题"]:
                分类[p] += 1
        print("\n问题分布：")
        for 类型, 数量 in 分类.most_common():
            print(f"  {类型:<12} {数量}")

    print(f"\n已输出：")
    print(f"  {os.path.abspath(输出文件)}")
    print(f"  {os.path.abspath(质量报告文件)}")
    print("=" * 56)


if __name__ == "__main__":
    main()
