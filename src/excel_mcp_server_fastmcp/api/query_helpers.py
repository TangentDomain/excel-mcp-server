"""
SQL 查询错误处理工具函数 — 从 advanced_sql_query.py 提取。

提供结构化错误提示、分类、修复建议，独立于查询引擎。
"""

import math
import re
from typing import Any

import numpy as np


class StructuredSQLError(Exception):
    """结构化SQL错误,为AI提供可自动修复的错误信息."""

    def __init__(self, error_code: str, message: str, hint: str = "", context: dict = None):
        self.error_code = error_code
        self.message = message
        self.hint = hint
        self.context = context or {}
        super().__init__(message)


def unsupported_error_hint(err_detail: str) -> str:
    """为UnsupportedError提供替代建议。"""
    err_upper = err_detail.upper()
    if "INSERT" in err_upper or "DELETE" in err_upper or "DROP" in err_upper or "ALTER" in err_upper or "CREATE" in err_upper:
        return "此工具仅支持SELECT查询.数据修改请使用excel_update_query(UPDATE语句)."
    if "NATURAL JOIN" in err_upper:
        return "不支持NATURAL JOIN,请改用显式ON条件:JOIN 表2 ON 表1.列 = 表2.列"
    if "FETCH" in err_upper or "NEXT" in err_upper:
        return "不支持FETCH/NEXT语法,请用LIMIT:SELECT ... LIMIT 10"
    if "RECURSIVE" in err_upper:
        return "不支持递归CTE(WITH RECURSIVE).请改用普通CTE或子查询."
    if "LATERAL" in err_upper:
        return "不支持LATERAL JOIN.请改用子查询或CTE."
    if "WINDOW" in err_upper and "OVER" not in err_upper:
        return "WINDOW子句请改为直接在窗口函数后写OVER:ROW_NUMBER() OVER (PARTITION BY ... ORDER BY ...)"
    return '请参考工具描述中"SQL已支持功能"列表,使用支持的操作.对于复杂计算,可考虑分步查询.'


def parse_error_hint(err_str: str, sql: str) -> str:
    """根据SQLGlot ParseError和原始SQL,生成AI可自动修复的提示."""
    sql_upper = sql.strip().upper()
    hint = ""
    typos = [
        ("SELEC ", "SELECT"),
        ("SELEC$", "SELECT"),
        ("FORM ", "FROM"),
        ("FORM$", "FROM"),
        ("WHER ", "WHERE"),
        ("WHER$", "WHERE"),
        ("GROUPBY", "GROUP BY"),
        ("GROUP  BY", "GROUP BY"),
        ("ORDERBY", "ORDER BY"),
        ("ORDER  BY", "ORDER BY"),
        ("HAVNIG", "HAVING"),
        ("HAVIN", "HAVING"),
        ("INNTER", "INNER"),
        ("LEFR", "LEFT"),
        ("RIGTH", "RIGHT"),
        ("JOINT", "JOIN"),
        ("OUDER", "OUTER"),
        ("DISTIN T", "DISTINCT"),
        ("DISTNCT", "DISTINCT"),
        ("BETWEE N", "BETWEEN"),
        ("BETWEN", "BETWEEN"),
        ("NOTNULL", "NOT NULL"),
        ("ISNUL", "IS NULL"),
        ("LIK E", "LIKE"),
        ("LIEK", "LIKE"),
        ("EXIS TS", "EXISTS"),
        ("EXIST ", "EXISTS"),
        ("LIMITT", "LIMIT"),
        ("OFFEST", "OFFSET"),
        ("ASCEND", "ASC"),
        ("DSCEND", "DESC"),
        ("CROS", "CROSS"),
        ("FUL L", "FULL"),
        ("UNIO N", "UNION"),
        ("UNON", "UNION"),
        ("INTERSE CT", "INTERSECT"),
        ("EXCEP T", "EXCEPT"),
        ("CONCATENATE", "CONCAT"),
        ("SUBSTITUE", "REPLACE"),
    ]
    for typo, correct in typos:
        if typo.rstrip("$") in sql_upper:
            if typo.endswith("$") and not sql_upper.rstrip(";").endswith(typo.rstrip("$")):
                continue
            return f'可能是拼写错误,"{typo.rstrip().rstrip("$")}" 应为 "{correct}"'
    order_keywords = ["SELECT", "FROM", "WHERE", "GROUP BY", "HAVING", "ORDER BY", "LIMIT"]
    found_positions = []
    for kw in order_keywords:
        if " " in kw:
            parts = kw.split()
            pos = sql_upper.find(parts[0])
            if pos != -1:
                after = sql_upper[pos + len(parts[0]) :].lstrip()
                if after.startswith(parts[1]):
                    found_positions.append((pos, kw))
        else:
            pos = sql_upper.find(kw)
            if pos != -1 and (pos == 0 or not sql_upper[pos - 1].isalpha()):
                found_positions.append((pos, kw))
    found_positions.sort(key=lambda x: x[0])
    for i in range(len(found_positions) - 1):
        pos1, kw1 = found_positions[i]
        pos2, kw2 = found_positions[i + 1]
        idx1 = order_keywords.index(kw1)
        idx2 = order_keywords.index(kw2)
        if idx1 > idx2:
            return f'SQL关键字顺序错误:"{kw1}"出现在"{kw2}"之前,但标准顺序要求"{kw1}"在"{kw2}"之后.正确顺序: {" -> ".join(order_keywords)}'
    if "GROUP BY" in sql_upper:
        agg_funcs = ["COUNT(", "SUM(", "AVG(", "MIN(", "MAX(", "COUNT (", "SUM (", "AVG (", "MIN (", "MAX ("]
        if not any(af in sql_upper for af in agg_funcs):
            return "GROUP BY通常与聚合函数一起使用(如COUNT/SUM/AVG/MIN/MAX).如果只是去重,请用SELECT DISTINCT."
    if re.search(r"\bJOIN\b", sql_upper) and " ON " not in sql_upper and not re.search(r"\bCROSS\s+JOIN\b", sql_upper):
        return "JOIN缺少ON条件.例如:... JOIN 表2 ON 表1.id = 表2.id.如果是笛卡尔积,请用CROSS JOIN."
    if "UPDATE" in sql_upper and "SET" in sql_upper and "SELECT" in sql_upper:
        return "不能在SELECT查询中使用UPDATE.批量修改请使用excel_update_query工具."
    select_match = re.search(r"\bSELECT\s+(.+?)\bFROM\b", sql_upper, re.DOTALL)
    if select_match:
        select_raw = sql[select_match.start(1) : select_match.end(1)]
        keywords_in_select = {
            "AS",
            "DISTINCT",
            "CASE",
            "WHEN",
            "THEN",
            "ELSE",
            "END",
            "AND",
            "OR",
            "NOT",
            "IN",
            "BETWEEN",
            "LIKE",
            "IS",
            "NULL",
            "TRUE",
            "FALSE",
            "COUNT",
            "SUM",
            "AVG",
            "MIN",
            "MAX",
            "UPPER",
            "LOWER",
            "TRIM",
            "LENGTH",
            "CONCAT",
            "REPLACE",
            "SUBSTRING",
            "LEFT",
            "RIGHT",
            "COALESCE",
            "IFNULL",
            "CAST",
            "ROW_NUMBER",
            "RANK",
            "DENSE_RANK",
            "LAG",
            "LEAD",
            "FIRST_VALUE",
            "LAST_VALUE",
            "OVER",
            "PARTITION",
            "ASC",
            "DESC",
            "ON",
        }
        adjacent_pairs = re.finditer(r"([A-Za-z_]\w*)\s+([A-Za-z_]\w*)", select_raw)
        for m in adjacent_pairs:
            t1, t2 = m.group(1), m.group(2)
            if t1.upper() not in keywords_in_select and t2.upper() not in keywords_in_select:
                return f'SELECT子句中"{t1}"和"{t2}"之间可能缺少逗号.列之间用逗号分隔:SELECT {t1}, {t2}'
    paren_count = sql.count("(") - sql.count(")")
    if paren_count > 0:
        return f'SQL中有{paren_count}个未闭合的括号.请检查每个左括号"("都有对应的右括号")".'
    if paren_count < 0:
        return f'SQL中有多余的{abs(paren_count)}个右括号")".请删除多余的括号.'
    single_quotes = len(re.findall(r"(?<!')'(?!')", sql))
    if single_quotes % 2 != 0:
        return "SQL中的单引号数量为奇数,可能有未闭合的引号.字符串值需要用单引号包裹,如 '值'."
    cn_punctuation = {"\uff0c": ",", "\uff08": "(", "\uff09": ")", "\uff1a": ":", "\uff1b": ";"}
    for cn, en in cn_punctuation.items():
        if cn in sql:
            return f'SQL中使用了中文标点"{cn}",应改为英文标点"{en}".'
    cross_file_bracket = re.search(r"\[[^\]]+\.xlsx?\]\.\w+", sql, re.IGNORECASE)
    if cross_file_bracket:
        matched = cross_file_bracket.group(0)
        return f"检测到跨文件引用语法 \"{matched}\",当前版本不支持 SQL Server 风格的 [文件名.xlsx].表名 语法。请使用 @'path' 语法进行跨文件查询，例如: FROM 表名@'/path/to/file.xlsx' alias"
    excel_funcs = {
        "SUMIF": "请用 CASE WHEN ... THEN ... END 替代 SUMIF",
        "COUNTIF": "请用 COUNT(CASE WHEN ... THEN 1 END) 替代 COUNTIF",
        "VLOOKUP": "请用 JOIN 替代 VLOOKUP",
        "IF": "请用 CASE WHEN ... THEN ... ELSE ... END 替代 IF 函数",
        "IFS": "请用 CASE WHEN ... THEN ... ELSE ... END 替代 IFS",
    }
    for func, suggestion in excel_funcs.items():
        if re.search(r"\b" + func + r"\s*\(", sql_upper):
            return f'Excel函数"{func}"不是SQL语法.{suggestion}.'
    subquery_pattern = re.search(r"\(\s*SELECT\b.+?\)\s*$", sql.strip(), re.IGNORECASE | re.DOTALL)
    if subquery_pattern:
        end_part = sql.strip()[subquery_pattern.end() :].strip()
        if not end_part or (not re.match(r"^AS\b", end_part, re.IGNORECASE) and not re.match(r"^[A-Za-z_]\w*$", end_part)):
            return "FROM子查询或UNION结果需要别名.例如:FROM (SELECT ...) AS subquery"
    if "SUBSTRING" in sql_upper and "(" in sql:
        substr_match = re.search(r"SUBSTRING\s*\((.+?)\)", sql, re.IGNORECASE)
        if substr_match:
            args = [a.strip() for a in substr_match.group(1).split(",")]
            if len(args) == 2:
                return "SUBSTRING需要3个参数:SUBSTRING(列, 起始位置, 长度).如果要从位置N取到末尾,请用SUBSTRING(列, N, LENGTH(列)-N+1)."
    return hint


def classify_value_error(err_str: str) -> str:
    """将ValueError分类为标准错误码。"""
    err_upper = err_str.upper()
    if "列 '" in err_str or "COLUMN" in err_upper:
        return "column_not_found"
    if "表 '" in err_str or "TABLE" in err_upper:
        return "table_not_found"
    if "窗口函数" in err_str or "WINDOW" in err_upper:
        return "window_function_error"
    if "JOIN" in err_upper:
        return "join_error"
    if "子查询" in err_str or "SUBQUERY" in err_upper:
        return "subquery_error"
    if "UNION" in err_upper:
        return "union_error"
    if "CTE" in err_upper or "WITH" in err_upper:
        return "cte_error"
    if "GROUP BY" in err_upper:
        return "group_by_error"
    if "ORDER BY" in err_upper:
        return "order_by_error"
    if "不支持" in err_str or "UNSUPPORTED" in err_upper:
        return "unsupported_feature"
    if "表达式" in err_str:
        return "expression_error"
    return "execution_error"


def generate_value_error_hint(err_str: str) -> str:
    """根据ValueError内容生成AI修复建议。"""
    if "列 '" in err_str and "可用列" in err_str:
        return "请检查列名拼写,或先用excel_get_headers查看可用列名."
    if "表 '" in err_str and "可用表" in err_str:
        return "请检查表名拼写,或先用excel_list_sheets查看可用工作表."
    if "FROM子查询" in err_str:
        return "请检查FROM子查询中的SQL语法和表名.FROM子查询需要别名:FROM (SELECT ...) AS alias."
    if "JOIN表" in err_str and "不存在" in err_str:
        return "请检查JOIN的表名是否正确,先用excel_list_sheets确认可用工作表."
    if "JOIN缺少ON条件" in err_str:
        return "JOIN必须包含ON条件,例如:... JOIN 表2 ON 表1.id = 表2.id."
    if "没有列 '" in err_str:
        return "请检查ON条件中的列名,确认列属于哪个表."
    if "不支持的窗口函数" in err_str:
        return "仅支持 ROW_NUMBER,RANK,DENSE_RANK,LAG,LEAD,FIRST_VALUE,LAST_VALUE 窗口函数."
    if "需要 ORDER BY" in err_str:
        return "该窗口函数必须包含 ORDER BY 子句."
    if "UNION" in err_str and "SELECT" in err_str:
        return "请确保UNION两侧的SELECT列数一致."
    if "数学表达式" in err_str:
        return "请检查数学运算符和操作数是否正确."
    if "字符串函数" in err_str:
        return "请检查函数名和参数.支持的字符串函数:UPPER/LOWER/TRIM/LENGTH/CONCAT/REPLACE/SUBSTRING/LEFT/RIGHT."
    return ""


def generate_value_error_suggested_fix(err_str: str, sql: str) -> str:
    """根据ValueError内容生成具体的修复SQL建议。"""
    if "列 '" in err_str and "你是否想用" in err_str:
        suggestion_match = re.search(r"你是否想用:\s*(.+?)\?", err_str)
        if suggestion_match:
            suggested_col = suggestion_match.group(1).strip().split(",")[0].strip()
            col_match = re.search(r"列 '(.+?)'", err_str)
            if col_match:
                return sql.replace(col_match.group(1), suggested_col)
    if "表 '" in err_str and "你是否想用" in err_str:
        suggestion_match = re.search(r"你是否想用:\s*(.+?)\?", err_str)
        if suggestion_match:
            suggested_table = suggestion_match.group(1).strip().split(",")[0].strip()
            col_match = re.search(r"表 '(.+?)'", err_str)
            if col_match:
                return sql.replace(col_match.group(1), suggested_table)
    return ""


def safe_float_comparison(left, right, op):
    """安全比较函数,处理None值。"""
    if left is None or right is None:
        return False
    try:
        left_float = float(left)
        right_float = float(right)
        if op == ">":
            return left_float > right_float
        elif op == ">=":
            return left_float >= right_float
        elif op == "<":
            return left_float < right_float
        elif op == "<=":
            return left_float <= right_float
        elif op == "==" or op == "=":
            max_val = max(abs(left_float), abs(right_float))
            epsilon = max(max_val * 1e-9, 1e-10)
            return abs(left_float - right_float) <= epsilon
        return False
    except (TypeError, ValueError):
        return False


def sanitize_float_for_excel(value: Any) -> Any:
    """清理浮点值,防止NaN/Inf/超范围值导致xlsx文件损坏。"""
    if value is None:
        return None
    if isinstance(value, (float, np.floating)):
        try:
            f_val = float(value)
            if np.isnan(f_val) or math.isinf(f_val):
                return None
            if abs(f_val) > 1e308:
                return 1e308 if f_val > 0 else -1e308
        except (ValueError, TypeError, OverflowError):
            pass
    return value
