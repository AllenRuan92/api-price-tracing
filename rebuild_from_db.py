# -*- coding: utf-8 -*-
"""
按当前口径重算 prices.xlsx（数据源：本地 prices_full.db 全量快照）

用途：改了 fetch_prices.WATCHLIST（增删产品线、调正则）之后，历史行仍是旧口径，
      跑一次本脚本，把 prices_full.db 里**所有日期**的快照按新口径重新挑选、
      重写整张 prices.xlsx，保证历史与当下口径一致、趋势线不出现假变动。

注意：prices_full.db 在 .gitignore 里，只存在于本地（GitHub Actions 每次跑都是空库）。
      所以本脚本只能覆盖本机已抓过的日期；要补更长的历史用 backfill_history.py --rebuild
      （从 tokenpricing 仓库的 history 目录拉，需要联网）。

用法：python rebuild_from_db.py
"""
import os
import sqlite3
import shutil
from datetime import datetime

import pandas as pd

import fetch_prices as fp


def load_snapshot(conn, date_str):
    """把某一天的全量快照还原成 fetch_prices 期望的 models dict 结构"""
    sql = """
        SELECT 模型ID, 模型名称, 输入价格_per_million, 输出价格_per_million,
               上下文长度, 数据来源
        FROM price_snapshots WHERE 抓取日期 = ?
    """
    models = {}
    for mid, name, ip, op, ctx, src in conn.execute(sql, (date_str,)):
        models[mid] = {
            "display_name": name,
            "pricing": {"input_per_million": ip, "output_per_million": op},
            "context_length": ctx,
            "sources": {s: 1 for s in (src or "").split(",") if s},
        }
    return models


def main():
    if not os.path.exists(fp.DB_PATH):
        print(f"找不到全量库 {fp.DB_PATH}，无法重算。")
        return

    # 先备份现有 Excel，出问题可回滚
    if os.path.exists(fp.EXCEL_PATH):
        backup = os.path.join(
            fp.SCRIPT_DIR, f"prices_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx")
        shutil.copy2(fp.EXCEL_PATH, backup)
        print(f"已备份原 Excel → {os.path.basename(backup)}")

    conn = sqlite3.connect(fp.DB_PATH)
    try:
        dates = [r[0] for r in conn.execute(
            "SELECT DISTINCT 抓取日期 FROM price_snapshots ORDER BY 抓取日期")]
        print(f"全量库共 {len(dates)} 个日期：{dates}")

        all_rows = []
        for d in dates:
            models = load_snapshot(conn, d)
            rows = fp.build_watchlist_rows(models, d)
            for r in rows:
                r["数据源"] = "prices_full.db"
            all_rows.extend(rows)
            print(f"  {d}: 全量 {len(models):>5} 个模型 → 命中主推产品线 {len(rows)} 个")
    finally:
        conn.close()

    if not all_rows:
        print("没有重算出任何行，已中止（未改动 Excel）。")
        return

    df = pd.DataFrame(all_rows, columns=fp.EXCEL_COLUMNS)
    df = df.sort_values(["抓取日期", "厂商", "产品线"]).reset_index(drop=True)
    with pd.ExcelWriter(fp.EXCEL_PATH, engine="openpyxl") as w:
        df.to_excel(w, sheet_name="价格数据", index=False)

    print(f"\n重算完成：{len(df)} 行写入 {os.path.basename(fp.EXCEL_PATH)}")
    print("各厂商产品线数：")
    for (prov, s), n in df[df["抓取日期"] == dates[-1]].groupby(["厂商", "产品线"]).size().items():
        print(f"  {prov:16s} {s}")
    print("\n下一步：python visualize.py 重新生成仪表盘")


if __name__ == "__main__":
    main()
