"""Excel MCP Server 性能基准工作负载。

测量 SQL-over-Excel 引擎的关键性能指标，为 autoresearch 自动迭代优化提供可量化的基线。

主指标 (lower is better):
    cold_query_ms — 全新引擎实例首次执行聚合查询的端到端延迟。
    覆盖完整热路径: HeaderAnalyzer 双重 I/O + calamine 读取 + dtype 优化 + sqlglot 解析 + pandas 执行。

次指标:
    cold_load_ms  — 纯数据加载耗时 (load_data_with_cache 首次)。
    warm_query_ms — 缓存命中时第二次聚合查询耗时 (基线下限)。
    orderby_ms    — ORDER BY DESC LIMIT 排序查询耗时。

工作负载: 10K 行 × 6 列双行表头游戏配置表 (固定随机种子, 可复现)。

设计原则:
    1. 每个指标用全新子进程测量, 隔离缓存污染。
    2. 主进程只负责生成测试文件 + 汇总子进程结果。
    3. 输出格式严格遵守 autoresearch 约定:
       METRIC <name>=<value>  (主指标, 供 harness 解析)
       ASI <name>=<value>     (次指标, 记录到 run metadata)
"""

from __future__ import annotations

import os
import random
import statistics
import subprocess
import sys
from pathlib import Path

REPO_ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(REPO_ROOT / "src"))

ROWS = 10000
COLS = 6
REPEAT = 3  # 每个指标重复测量取中位数, 降低噪声
SEED = 20260629

TEST_FILE = REPO_ROOT / "tools" / "perf-benchmark" / "fixture_10k.xlsx"

# 子进程测量脚本: 每个 metric 一个独立进程, 避免缓存共享
# 通过环境变量 _BENCH_METRIC 选择要测量的指标
_PROBE = """
import os, sys, time, gc
sys.path.insert(0, os.path.join(os.environ["_BENCH_ROOT"], "src"))
fp = os.environ["_BENCH_FILE"]
metric = os.environ["_BENCH_METRIC"]

from excel_mcp_server_fastmcp.api.advanced_sql_query import AdvancedSQLQueryEngine
engine = AdvancedSQLQueryEngine()
gc.collect()

if metric == "cold_load":
    t0 = time.perf_counter()
    data = engine._load_data_with_cache(fp)
    t1 = time.perf_counter()
    assert data and "Sheet1" in data, "加载失败"
    print(f"VALUE={t1 - t0}")

elif metric == "cold_query":
    sql = "SELECT 稀有度, COUNT(*) AS cnt, AVG(攻击力) AS avg_atk FROM Sheet1 GROUP BY 稀有度"
    t0 = time.perf_counter()
    r = engine.execute_sql_query(file_path=fp, sql=sql)
    t1 = time.perf_counter()
    assert r.get("success"), f"查询失败: {r.get('message','')[:120]}"
    print(f"VALUE={t1 - t0}")

elif metric == "warm_query":
    # 预热: 先加载一次填充缓存
    engine._load_data_with_cache(fp)
    sql = "SELECT 稀有度, COUNT(*) AS cnt, AVG(攻击力) AS avg_atk FROM Sheet1 GROUP BY 稀有度"
    t0 = time.perf_counter()
    r = engine.execute_sql_query(file_path=fp, sql=sql)
    t1 = time.perf_counter()
    assert r.get("success"), f"查询失败: {r.get('message','')[:120]}"
    print(f"VALUE={t1 - t0}")

elif metric == "orderby":
    sql = "SELECT * FROM Sheet1 ORDER BY 攻击力 DESC LIMIT 10"
    t0 = time.perf_counter()
    r = engine.execute_sql_query(file_path=fp, sql=sql)
    t1 = time.perf_counter()
    assert r.get("success"), f"查询失败: {r.get('message','')[:120]}"
    print(f"VALUE={t1 - t0}")

else:
    raise SystemExit(f"unknown metric: {metric}")
"""


def build_fixture() -> None:
    """构建 10K 行 × 6 列双行表头测试文件 (固定种子, 可复现)."""
    if TEST_FILE.exists():
        # 已存在则跳过, 保持每次迭代文件一致 (mtime 也固定, 命中缓存逻辑可复现)
        return
    from openpyxl import Workbook

    rng = random.Random(SEED)
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"

    # 双行表头: 第1行中文描述, 第2行英文字段名
    ws.append(["ID", "名称", "稀有度", "攻击力", "生命值", "价格"])
    ws.append(["item_id", "item_name", "rarity", "attack", "hp", "price"])

    rarities = ["Common", "Rare", "Epic", "Legendary", "Mythic"]
    for i in range(1, ROWS + 1):
        ws.append(
            [
                i,
                f"Item-{i:05d}",
                rng.choice(rarities),
                rng.randint(1, 1000),
                rng.randint(50, 9999),
                round(rng.uniform(0.5, 500.0), 2),
            ]
        )

    wb.save(TEST_FILE)
    wb.close()


def measure_once(metric: str) -> float:
    """在独立子进程中测量一次指标, 返回耗时(秒)."""
    env = os.environ.copy()
    env["_BENCH_ROOT"] = str(REPO_ROOT)
    env["_BENCH_FILE"] = str(TEST_FILE)
    env["_BENCH_METRIC"] = metric

    proc = subprocess.run(
        [sys.executable, "-c", _PROBE],
        env=env,
        capture_output=True,
        text=True,
        cwd=str(REPO_ROOT),
        timeout=120,
    )
    if proc.returncode != 0:
        sys.stderr.write(f"[benchmark] metric={metric} 子进程失败:\n{proc.stderr}\n")
        raise SystemExit(f"子进程测量失败: {metric}")

    for line in proc.stdout.strip().splitlines():
        line = line.strip()
        if line.startswith("VALUE="):
            return float(line.split("=", 1)[1])
    raise SystemExit(f"子进程未返回 VALUE: {metric}\nstdout={proc.stdout}\nstderr={proc.stderr}")


def measure_median(metric: str, repeat: int = REPEAT) -> float:
    """重复测量取中位数, 降低单次噪声."""
    samples = [measure_once(metric) for _ in range(repeat)]
    return statistics.median(samples)


def main() -> int:
    """运行全部性能指标测量并输出 autoresearch 约定格式的 METRIC/ASI 行。"""
    build_fixture()

    print(f"[benchmark] fixture: {TEST_FILE.name} ({ROWS} rows × {COLS} cols)", file=sys.stderr)
    print(f"[benchmark] repeat per metric: {REPEAT}", file=sys.stderr)

    # 各指标独立测量, 互不污染
    cold_query = measure_median("cold_query")
    cold_load = measure_median("cold_load")
    warm_query = measure_median("warm_query")
    orderby = measure_median("orderby")

    # 输出 autoresearch 约定格式 (单位: 毫秒)
    # METRIC 行被 run_experiment 自动解析为主指标
    # ASI 行记录为次指标 metadata
    print(f"METRIC cold_query_ms={cold_query * 1000:.2f}")
    print(f"ASI cold_load_ms={cold_load * 1000:.2f}")
    print(f"ASI warm_query_ms={warm_query * 1000:.2f}")
    print(f"ASI orderby_ms={orderby * 1000:.2f}")
    print(f"ASI cold_warm_ratio={(cold_query / warm_query if warm_query > 0 else 0):.1f}")

    print(
        f"[benchmark] cold_query={cold_query * 1000:.1f}ms  cold_load={cold_load * 1000:.1f}ms  warm_query={warm_query * 1000:.1f}ms  orderby={orderby * 1000:.1f}ms",
        file=sys.stderr,
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
