# -*- coding: utf-8 -*-
"""探查 tokenpricing 数据源是否有历史快照目录（用于回填六月历史）"""
import requests, json

BASE = "https://api.github.com/repos/Atena-IT/tokenpricing/contents/database"
H = {"User-Agent": "probe", "Accept": "application/vnd.github+json"}

def ls(path):
    url = f"https://api.github.com/repos/Atena-IT/tokenpricing/contents/{path}"
    r = requests.get(url, headers=H, timeout=30)
    if r.status_code != 200:
        print(f"  [{r.status_code}] {path} 访问失败")
        return []
    items = r.json()
    return items if isinstance(items, list) else []

print("=== database/ 目录结构 ===")
for it in ls("database"):
    print(f"  {it['type']:5}  {it['name']}")

# 常见的历史目录名探测
for cand in ["database/history", "database/snapshots", "database/archive", "database/daily", "database/historical"]:
    items = ls(cand)
    if items:
        print(f"\n=== 发现历史目录 {cand}（前20项）===")
        for it in items[:20]:
            print(f"  {it['type']:5}  {it['name']}")

# 看 current 目录里有什么
print("\n=== database/current/ 目录 ===")
for it in ls("database/current"):
    print(f"  {it['type']:5}  {it['name']}  ({it.get('size','?')} bytes)")

# 用 git commits API 看 prices.json 的历史提交（能借此回填）
print("\n=== prices.json 最近提交历史（可借 commit 回填）===")
url = "https://api.github.com/repos/Atena-IT/tokenpricing/commits?path=database/current/prices.json&per_page=30"
r = requests.get(url, headers=H, timeout=30)
if r.status_code == 200:
    commits = r.json()
    print(f"  共找到 {len(commits)} 条提交（取最近30）：")
    seen_dates = set()
    for c in commits:
        d = c["commit"]["committer"]["date"][:10]
        seen_dates.add(d)
        print(f"    {c['commit']['committer']['date']}  {c['sha'][:8]}")
    print(f"\n  覆盖的不同日期：{sorted(seen_dates)}")
else:
    print(f"  [{r.status_code}] 提交历史访问失败")
