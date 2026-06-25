# -*- coding: utf-8 -*-
"""临时探查脚本：从 prices_full.db 里找出豆包(Doubao)相关的真实模型 ID"""
import os, sqlite3

DB = os.path.join(os.path.dirname(os.path.abspath(__file__)), "prices_full.db")
conn = sqlite3.connect(DB)
today = conn.execute("SELECT MAX(抓取日期) FROM price_snapshots").fetchone()[0]
print(f"最新抓取日期: {today}\n")

# 关键词匹配：doubao / 豆包 / bytedance / seed / volc
keys = ["doubao", "豆包", "bytedance", "byte", "seed", "volc", "ark"]
rows = conn.execute(
    "SELECT 模型ID, 厂商, 模型名称, 输入价格_per_million, 输出价格_per_million "
    "FROM price_snapshots WHERE 抓取日期=?", (today,)
).fetchall()

hits = []
for mid, prov, name, inp, outp in rows:
    blob = f"{mid} {prov} {name}".lower()
    if any(k in blob for k in keys):
        hits.append((mid, prov, name, inp, outp))

print(f"匹配到 {len(hits)} 个疑似豆包/字节模型：\n")
for mid, prov, name, inp, outp in sorted(hits):
    print(f"  {mid}")
    print(f"     厂商={prov} 名称={name} 输入=${inp}/M 输出=${outp}/M")
conn.close()
