# -*- coding: utf-8 -*-
"""
SQL 校准器核心逻辑
==================
从 sql_calibrator.py 提取的核心功能，适配 MCP 工具和 CLI 双模式。

核心命令均返回 dict 结构化结果，由调用方决定如何输出。
"""

import os
import re
import sqlite3
import time
from pathlib import Path

import pandas as pd

# ============================================================
# 常量配置
# ============================================================
DEFAULT_DB_DIR = "/tmp/calibrator_data/"


# ============================================================
# 工具函数
# ============================================================

def get_db_path(db_name: str) -> str:
    """获取数据库文件的完整路径"""
    os.makedirs(DEFAULT_DB_DIR, exist_ok=True)
    safe_name = re.sub(r'[\\/*?:"<>|]', '_', db_name)
    if not safe_name.endswith('.db'):
        safe_name += '.db'
    return os.path.join(DEFAULT_DB_DIR, safe_name)


def sanitize_table_name(name: str) -> str:
    """清理表名：去掉非法字符，确保SQLite合法"""
    name = re.sub(r'^[0-9]+', '', name)
    name = re.sub(r'[\\/*?:"<>|\s]', '_', name)
    name = name.strip('_')
    if not name:
        name = 'unnamed_table'
    return name


def sanitize_col_name(name: str) -> str:
    """
    清理列名中的特殊字符，使其在 SQLite 方括号引用中安全
    保留中文、英文、数字、下划线、点号、括号、井号、减号等常见字符
    """
    if not name or name.strip() == '':
        return 'unknown_column'
    name = str(name).strip()
    # 去掉可能导致问题的控制字符
    name = re.sub(r'[\x00-\x1f]', '', name)
    if not name:
        return 'unknown_column'
    return name


def flatten_multiindex(columns) -> list:
    """
    智能扁平化 pandas MultiIndex 列名

    处理规则：
    - ('宝箱ID', '')          → '宝箱ID'           （第二层为空时取第一层）
    - ('道具信息', 'item_id') → '道具信息.item_id'  （两层都有值则用点连接）
    - ('备注', 'Unnamed: 7_level_1') → '备注'       （跳过 Unnamed 自动列）
    - 三层/四层表头同理递归处理

    Args:
        columns: pandas MultiIndex 对象

    Returns:
        扁平化的列名列表
    """
    result = []

    if not isinstance(columns, pd.MultiIndex):
        # 不是 MultiIndex，直接返回字符串列表
        return [sanitize_col_name(str(c)) for c in columns]

    for col in columns:
        if isinstance(col, tuple):
            parts = []
            for level in col:
                s = str(level).strip() if pd.notna(level) else ''
                # 跳过空值和 pandas 自动生成的 Unnamed 列
                if s and s.lower() != 'nan' and not s.lower().startswith('unnamed'):
                    parts.append(s)

            if len(parts) == 0:
                result.append('unknown_column')
            elif len(parts) == 1:
                result.append(sanitize_col_name(parts[0]))
            else:
                result.append(sanitize_col_name('.'.join(parts)))
        else:
            result.append(sanitize_col_name(str(col).strip()))

    return result


def is_likely_dual_header(df_raw_row0, df_raw_row1):
    """
    智能判断是否为双表头格式

    双表头的特征（游戏配表典型模式）：
    - 第0行是中文分类名（如"开启条件"、"宝箱道具ID"）
    - 第1行是实际的英文/具体列名（如"UnlockCondition.Items#1.Type"、"ChestPropID"）
    - 第1行看起来像列名（字符串为主），不像数据（数字为主）

    单表头的特征：
    - 第0行已经是完整的列名
    - 第1行是数据行（包含大量数字）

    Args:
        df_raw_row0: 第0行的值列表
        df_raw_row1: 第1行的值列表

    Returns:
        bool: 是否判断为双表头
    """
    # 统计第1行中字符串类型 vs 数值类型的比例
    str_count = 0
    num_count = 0
    total = 0

    for val in df_raw_row1:
        total += 1
        if pd.isna(val):
            continue
        s = str(val).strip()
        # 尝试判断是否为数值
        try:
            float(s)
            num_count += 1
        except (ValueError, TypeError):
            str_count += 1

    if total == 0:
        return False

    # 如果第1行大部分是字符串而非数字，更像是双表头
    str_ratio = str_count / total

    # 额外检查：第0行和第1行是否有明显不同的模式
    # 双表头通常第0行较短（分类名），第1行较长且带点号（如 UnlockCondition.Items#1.Type）
    has_dotted_names = False
    for val in df_raw_row1:
        if pd.notna(val):
            s = str(val).strip()
            if '.' in s and len(s) > 10:
                has_dotted_names = True
                break

    # 判定逻辑：
    # - 第1行字符串占比 > 60% 且包含带点号的长列名 → 双表头
    # - 或者第1行全是字符串 → 可能是双表头
    if str_ratio > 0.6 and has_dotted_names:
        return True
    if str_ratio > 0.85:
        return True

    return False


def infer_sqlite_type(series: pd.Series) -> str:
    """
    根据 pandas 列数据推断 SQLite 类型
    - 全部为空 → TEXT
    - 可以转整数 → INTEGER
    - 可以转浮点数 → REAL
    - 其他 → TEXT
    """
    non_null = series.dropna()
    if len(non_null) == 0:
        return 'TEXT'

    # 尝试转为整数
    try:
        non_null.astype(int)
        return 'INTEGER'
    except (ValueError, TypeError):
        pass

    # 尝试转为浮点数
    try:
        non_null.astype(float)
        return 'REAL'
    except (ValueError, TypeError):
        pass

    return 'TEXT'


def format_table(rows, headers):
    """
    格式化表格输出，优先使用 tabulate，否则手动对齐
    """
    # 转换 None 为 NULL 显示
    display_rows = []
    for row in rows:
        display_rows.append(tuple('NULL' if v is None else str(v) for v in row))

    try:
        from tabulate import tabulate
        return tabulate(display_rows, headers=headers, tablefmt='grid',
                       showindex=False, missingval='NULL')
    except ImportError:
        pass

    # 手动实现表格对齐
    if not display_rows:
        return "(空结果集)"

    num_cols = len(headers)
    col_widths = [len(str(h)) for h in headers]

    for row in display_rows:
        for i, val in enumerate(row):
            if i < num_cols:
                col_widths[i] = max(col_widths[i], len(val))

    padding = 2
    col_widths = [w + padding for w in col_widths]

    sep = '+' + '+'.join('-' * w for w in col_widths) + '+'
    hdr = '|' + '|'.join(
        str(h).ljust(col_widths[i]) for i, h in enumerate(headers)
    ) + '|'

    lines = [sep, hdr, sep]
    for row in display_rows:
        cells = []
        for i in range(num_cols):
            val = row[i] if i < len(row) else ''
            cells.append(val.ljust(col_widths[i]))
        lines.append('|' + '|'.join(cells) + '|')
    lines.append(sep)

    return '\n'.join(lines)


# ============================================================
# 核心命令实现（返回结构化 dict，不直接 print）
# ============================================================

def cmd_import(xlsx_path: str, db_name: str = "default") -> dict:
    """
    导入 Excel 文件到 SQLite 数据库

    支持功能：
    - 自动检测双表头（MultiIndex）并智能拍平
    - 多 Sheet 分别建表
    - 自动推断列类型
    - 添加自增主键 _rowid_

    Returns:
        dict: {success, message, db_path, tables: [{name, rows, columns}], total_tables, total_rows}
    """
    xlsx_path = os.path.abspath(xlsx_path)
    if not os.path.exists(xlsx_path):
        return {
            "success": False,
            "message": f"文件不存在: {xlsx_path}",
        }

    db_path = get_db_path(db_name)

    try:
        xl = pd.ExcelFile(xlsx_path)
    except Exception as e:
        return {
            "success": False,
            "message": f"无法读取 Excel 文件: {e}",
        }

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    total_tables = 0
    total_rows = 0
    table_details = []
    logs = []

    for sheet_name in xl.sheet_names:
        table_name = sanitize_table_name(sheet_name)
        logs.append(f"处理 Sheet: '{sheet_name}' -> 表 '{table_name}'")

        # ===== 第一步：读取原始前两行来判断表头类型 =====
        df_raw = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=None, nrows=3)

        if df_raw.empty or len(df_raw) < 1:
            logs.append(f"  [警告] 跳过空 Sheet")
            continue

        row0 = df_raw.iloc[0].tolist()   # 可能的表头行1
        row1 = df_raw.iloc[1].tolist() if len(df_raw) > 1 else []  # 可能的表头行2 或 数据行1

        # ===== 第二步：智能判断单/双表头 =====
        use_multiheader = False
        final_columns = []
        df_data = None

        if len(df_raw) >= 2 and is_likely_dual_header(row0, row1):
            # 判定为双表头
            use_multiheader = True
            df_data = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=[0, 1])
            final_columns = flatten_multiindex(df_data.columns)
            logs.append(f"  ★ 检测到双表头格式，已智能拍平 ({len(final_columns)} 列)")
        else:
            # 单表头
            df_data = pd.read_excel(xlsx_path, sheet_name=sheet_name, header=0)
            final_columns = flatten_multiindex(df_data.columns)
            logs.append(f"  使用单表头格式 ({len(final_columns)} 列)")

        # ===== 第三步：清理数据 =====
        df_data = df_data.dropna(how='all')    # 去掉全空行
        df_data = df_data.dropna(how='all', axis=1)  # 去掉全空列

        if len(df_data) == 0:
            logs.append(f"  [警告] 跳过空表")
            continue

        # 更新最终列名（重新扁平化，因为 dropna 后可能变化）
        if use_multiheader:
            final_columns = flatten_multiindex(df_data.columns)
        else:
            final_columns = [sanitize_col_name(str(c)) for c in df_data.columns]

        # 确保列名唯一
        seen = {}
        unique_cols = []
        for col in final_columns:
            base = col
            if col in seen:
                seen[col] += 1
                col = f"{col}_{seen[col]}"
            else:
                seen[col] = 0
            unique_cols.append(col)

        df_data.columns = unique_cols
        final_columns = unique_cols

        # 显示列名预览
        preview = ', '.join(final_columns[:10])
        if len(final_columns) > 10:
            preview += f"... (共{len(final_columns)}列)"
        logs.append(f"  列名: {preview}")

        # ===== 第四步：建表 =====
        cursor.execute(f"DROP TABLE IF EXISTS [{table_name}]")

        col_defs = ["_rowid_ INTEGER PRIMARY KEY AUTOINCREMENT"]
        for col_name in final_columns:
            col_type = infer_sqlite_type(df_data[col_name])
            col_defs.append(f"[{sanitize_col_name(col_name)}] {col_type}")

        create_sql = f"CREATE TABLE [{table_name}] ({', '.join(col_defs)})"

        try:
            cursor.execute(create_sql)
        except sqlite3.Error as e:
            logs.append(f"  [错误] 建表失败: {e}")
            continue

        # ===== 第五步：插入数据 =====
        placeholders = ', '.join(['?' for _ in final_columns])
        insert_sql = f"INSERT INTO [{table_name}] VALUES (NULL, {placeholders})"

        inserted = 0
        errors = 0
        for idx, row in df_data.iterrows():
            values = []
            for col in final_columns:
                val = row.get(col)
                if pd.isna(val):
                    values.append(None)
                else:
                    values.append(val)

            try:
                cursor.execute(insert_sql, values)
                inserted += 1
            except Exception as e:
                errors += 1
                if errors <= 3:
                    logs.append(f"  [警告] 第{idx+1}行插入失败: {e}")

        if errors > 3:
            logs.append(f"  ... 共 {errors} 行插入失败")

        total_tables += 1
        total_rows += inserted
        table_details.append({
            "sheet_name": sheet_name,
            "table_name": table_name,
            "rows": inserted,
            "columns": final_columns,
            "errors": errors,
        })
        logs.append(f"  ✓ 插入 {inserted} 行数据" + (f" ({errors} 行失败)" if errors else ""))

    conn.commit()
    conn.close()

    return {
        "success": True,
        "message": f"导入完成！共 {total_tables} 张表, {total_rows} 行数据",
        "db_path": db_path,
        "tables": table_details,
        "total_tables": total_tables,
        "total_rows": total_rows,
        "logs": logs,
    }


def cmd_query(db_name: str, sql: str) -> dict:
    """
    执行 SQL 查询并以表格形式返回结果

    Returns:
        dict: {success, message, headers, rows, row_count, elapsed_ms, formatted}
    """
    db_path = get_db_path(db_name)

    if not os.path.exists(db_path):
        return {
            "success": False,
            "message": f"数据库不存在: {db_path}\n请先使用 import 命令导入数据",
        }

    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    start_time = time.perf_counter()

    try:
        cursor.execute(sql)
    except sqlite3.Error as e:
        conn.close()
        return {
            "success": False,
            "message": f"SQL 错误: {e}",
        }

    rows = cursor.fetchall()
    elapsed_ms = (time.perf_counter() - start_time) * 1000

    if cursor.description:
        headers = [desc[0] for desc in cursor.description]
    else:
        headers = []

    row_data = [tuple(row) for row in rows]

    return {
        "success": True,
        "headers": headers,
        "rows": row_data,
        "row_count": len(row_data),
        "elapsed_ms": round(elapsed_ms, 2),
        "formatted": format_table(row_data, headers),
    }


def cmd_tables(db_name: str) -> dict:
    """
    列出数据库中的所有表及其行数

    Returns:
        dict: {success, message, tables: [{name, count}], formatted}
    """
    db_path = get_db_path(db_name)

    if not os.path.exists(db_path):
        return {
            "success": False,
            "message": f"数据库不存在: {db_path}",
        }

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute(
        "SELECT name FROM sqlite_master "
        "WHERE type='table' AND name NOT LIKE 'sqlite_%' "
        "ORDER BY name"
    )
    tables = [row[0] for row in cursor.fetchall()]

    if not tables:
        conn.close()
        return {
            "success": True,
            "tables": [],
            "formatted": "(空数据库，没有任何表)",
        }

    table_info = []
    for t in tables:
        cursor.execute(f"SELECT COUNT(*) FROM [{t}]")
        count = cursor.fetchone()[0]
        table_info.append((t, count))
    conn.close()

    return {
        "success": True,
        "tables": [{"name": t, "count": c} for t, c in table_info],
        "formatted": format_table(table_info, ['表名', '行数']),
    }


def cmd_schema(db_name: str, table_name: str) -> dict:
    """
    显示指定表的结构信息（列名、类型、约束等）

    Returns:
        dict: {success, message, table_name, columns: [{cid, name, type, notnull, pk, default}], formatted}
    """
    db_path = get_db_path(db_name)

    if not os.path.exists(db_path):
        return {
            "success": False,
            "message": f"数据库不存在: {db_path}",
        }

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()

    cursor.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name=?",
        (table_name,)
    )
    if not cursor.fetchone():
        conn.close()
        return {
            "success": False,
            "message": f"表 '{table_name}' 不存在",
        }

    cursor.execute(f"PRAGMA table_info([{table_name}])")
    columns = cursor.fetchall()
    conn.close()

    rows = []
    for col in columns:
        cid, name, dtype, notnull, default, pk = col
        rows.append({
            "cid": cid,
            "name": name,
            "type": dtype,
            "notnull": 'NOT NULL' if notnull else 'NULL',
            "pk": 'PK' if pk else '',
            "default": str(default) if default is not None else '',
        })

    return {
        "success": True,
        "table_name": table_name,
        "columns": rows,
        "formatted": format_table(
            [(r['cid'], r['name'], r['type'], r['notnull'], r['pk'], r['default']) for r in rows],
            ['CID', '列名', '类型', '约束', '主键', '默认值']
        ),
    }
