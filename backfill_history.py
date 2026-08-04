# -*- coding: utf-8 -*-
"""
回填 / 重算历史数据（数据源：tokenpricing 仓库的 database/history/ 快照目录）

两种模式：
  python backfill_history.py             只补 prices.xlsx 里缺失的日期，已有日期不动
  python backfill_history.py --rebuild   **全量重算**：无视已有内容，把 history 目录
                                         能拿到的每一天都按当前 WATCHLIST 重新挑选，
                                         整张表重写

什么时候用 --rebuild：改了 fetch_prices.WATCHLIST（增删产品线、调正则）之后。
历史行若仍是旧口径，趋势图会显示出实际并不存在的价格跳变。
"""
import os, sys, time, json, re
from collections import defaultdict
import requests
import pandas as pd

import fetch_prices as fp  # 复用配置、白名单规则与列定义

H = {"User-Agent": "backfill", "Accept": "application/vnd.github+json"}
HIST_API = "https://api.github.com/repos/Atena-IT/tokenpricing/contents/database/history"
RAW_BASE = "https://raw.githubusercontent.com/Atena-IT/tokenpricing/main/database/history/"

EXCEL_PATH = fp.EXCEL_PATH


def get_with_retry(url, tries=4, delay=4, is_json=True):
    for i in range(1, tries + 1):
        try:
            r = requests.get(url, headers=H, timeout=40)
            r.raise_for_status()
            return r.json() if is_json else r.text
        except Exception as e:
            print(f"  请求失败({i}/{tries}) {type(e).__name__}: {str(e)[:80]}")
            if i < tries:
                time.sleep(delay)
    return None


def list_history_files():
    """列出 history 目录所有快照文件名"""
    items = get_with_retry(HIST_API)
    if not items:
        print("无法列出 history 目录")
        return []
    files = [it["name"] for it in items if it["type"] == "file" and it["name"].endswith(".json")]
    return sorted(files)


def pick_one_per_day(files):
    """每天只取第一个时间点的文件。文件名形如 prices-20260614T184401Z.json"""
    by_day = {}
    for f in files:
        m = re.search(r"prices-(\d{8})T(\d{6})Z", f)
        if not m:
            continue
        day = m.group(1)  # YYYYMMDD
        if day not in by_day or f < by_day[day]:
            by_day[day] = f
    return dict(sorted(by_day.items()))


def extract_rows(models, date_str):
    """从一份快照的 models dict 按白名单规则抽出主推产品线行（复用 fetch_prices 逻辑）"""
    rows = fp.build_watchlist_rows(models, date_str)
    for r in rows:
        r["数据源"] = "tokenpricing-history"
    return rows


def main():
    rebuild = "--rebuild" in sys.argv
    print("=== 历史全量重算（按当前 WATCHLIST 口径重写整张表）===" if rebuild
          else "=== 回填缺失日期（已有日期保持不动）===")

    # 已有数据 / 已有日期
    existing_dates = set()
    cur = pd.DataFrame(columns=fp.EXCEL_COLUMNS)
    if os.path.exists(EXCEL_PATH):
        try:
            old = pd.read_excel(EXCEL_PATH, sheet_name="价格数据")
            old["抓取日期"] = pd.to_datetime(old["抓取日期"]).dt.strftime("%Y-%m-%d")
            if rebuild:
                print(f"现有 {len(old)} 行将被整表替换（旧口径数据不保留）")
            elif "产品线" not in old.columns:
                print("现有 Excel 是旧口径（无「产品线」列），与新口径不可比 → 自动切换为全量重算")
                rebuild = True
            else:
                cur = old
                existing_dates = set(old["抓取日期"].unique())
                print(f"prices.xlsx 已有日期：{sorted(existing_dates)}")
        except Exception as e:
            print(f"读取现有 Excel 失败（将全量重建）：{e}")
            rebuild = True

    files = list_history_files()
    print(f"history 目录共 {len(files)} 个快照文件")
    if not files:
        print("拿不到 history 目录，中止（未改动 Excel）。")
        return
    per_day = pick_one_per_day(files)
    print(f"覆盖 {len(per_day)} 个不同日期：{list(per_day.keys())}")

    all_new = []
    for day, fname in per_day.items():
        date_str = f"{day[:4]}-{day[4:6]}-{day[6:8]}"
        if not rebuild and date_str in existing_dates:
            print(f"  跳过 {date_str}（已存在）")
            continue
        print(f"  拉取 {date_str} ← {fname}")
        data = get_with_retry(RAW_BASE + fname)
        if not data:
            print(f"    ✗ {date_str} 拉取失败，跳过")
            continue
        models = data.get("models", {})
        rows = extract_rows(models, date_str)
        all_new.extend(rows)
        print(f"    ✓ {date_str}: 全量 {len(models)} 个模型 → 命中主推产品线 {len(rows)} 个")
        time.sleep(0.5)  # 轻微限速，避免触发 GitHub 限流

    if not all_new:
        print("没有需要写入的数据（未改动 Excel）。")
        return

    new_df = pd.DataFrame(all_new, columns=fp.EXCEL_COLUMNS)
    combined = new_df if rebuild else pd.concat([cur, new_df], ignore_index=True)
    combined = combined.sort_values(["抓取日期", "厂商", "产品线"]).reset_index(drop=True)
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl") as w:
        combined.to_excel(w, sheet_name="价格数据", index=False)
    action = "重算" if rebuild else "回填"
    print(f"\n{action}完成：写入 {len(all_new)} 行，prices.xlsx 共 {len(combined)} 行")
    print(f"现有日期：{sorted(set(combined['抓取日期']))}")


if __name__ == "__main__":
    main()
