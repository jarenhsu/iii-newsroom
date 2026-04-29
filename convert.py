#!/usr/bin/env python3
"""
資策會新聞月報轉換工具
用法：python convert.py Opview_Insight_Result_臨時新聞事件_2026-04-29.xlsx
"""
import sys, json, re
from pathlib import Path
from collections import defaultdict
import pandas as pd

DEPTS = [
    {"id":"digital",   "patterns":["數位轉型研究院","數轉院"]},
    {"id":"software",  "patterns":["軟體技術研究院","軟體院"]},
    {"id":"ai",        "patterns":["人工智慧研究院","AI院","ai院","AI研究院"]},
    {"id":"cybersec",  "patterns":["資安科技研究所","資安所"]},
    {"id":"education", "patterns":["數位教育研究所","教研所","數位教育所"]},
    {"id":"mic",       "patterns":["產業情報研究所","MIC","IEK"]},
    {"id":"law",       "patterns":["科技法律研究所","科法所"]},
]
DEPT_META = [
    {"id":"digital",  "name":"數位轉型研究院","short":"數轉院","color":"#34d399"},
    {"id":"software", "name":"軟體技術研究院","short":"軟體院","color":"#60a5fa"},
    {"id":"ai",       "name":"人工智慧研究院","short":"AI院",  "color":"#a78bfa"},
    {"id":"cybersec", "name":"資安科技研究所","short":"資安所","color":"#f87171"},
    {"id":"education","name":"數位教育研究所","short":"教研所","color":"#fbbf24"},
    {"id":"mic",      "name":"產業情報研究所","short":"MIC",   "color":"#fb923c"},
    {"id":"law",      "name":"科技法律研究所","short":"科法所","color":"#e879f9"},
    {"id":"other",    "name":"其他/機構整體", "short":"其他",  "color":"#64748b"},
]

def classify(title, content):
    text = str(title) + str(content)
    for d in DEPTS:
        for p in d["patterns"]:
            if p in text: return d["id"]
    return "other"

def ngrams(text, n=4):
    t = str(text)
    return set(t[i:i+n] for i in range(len(t)-n+1))

def similarity(a, b):
    sa, sb = ngrams(a), ngrams(b)
    if not sa or not sb: return 0
    return len(sa & sb) / len(sa | sb)

def cluster_news(news_list, threshold=0.2):
    titles = [n['title'] for n in news_list]
    n = len(titles)
    parent = list(range(n))
    def find(x):
        while parent[x] != x: parent[x] = parent[parent[x]]; x = parent[x]
        return x
    def union(x, y): parent[find(x)] = find(y)
    for i in range(n):
        for j in range(i+1, n):
            if similarity(titles[i], titles[j]) >= threshold:
                union(i, j)
    groups = defaultdict(list)
    for i in range(n): groups[find(i)].append(i)
    clusters = []
    for idxs in sorted(groups.values(), key=lambda c: -len(c)):
        if len(idxs) < 2: continue
        rep_titles = sorted([titles[i] for i in idxs if titles[i]], key=len)
        rep = rep_titles[0] if rep_titles else ''
        clusters.append({"title": rep, "count": len(idxs)})
    return clusters

def convert(xlsx_path: str):
    path = Path(xlsx_path)
    if not path.exists():
        print(f"❌ 找不到檔案：{xlsx_path}"); sys.exit(1)

    print(f"📖 讀取 {path.name} ...")
    df = pd.read_excel(path, sheet_name=0)
    df['_date'] = pd.to_datetime(df['發布時間'], errors='coerce').dt.strftime('%Y-%m-%d')

    m = re.search(r'(\d{4})-(\d{2})', path.name)
    if m:
        month = f"{m.group(1)}-{m.group(2)}"
    else:
        valid = df['_date'].dropna()
        month = sorted(valid)[0][:7] if len(valid) else __import__('datetime').date.today().strftime('%Y-%m')

    year, mon = month.split('-')
    month_label = f"{year}年{int(mon)}月"
    print(f"📅 月份：{month_label}")

    dsent = {d["id"]: {"正面":0,"中立":0,"負面":0} for d in DEPT_META}
    news_list = []

    for _, row in df.iterrows():
        title   = str(row.get('標題',''))     if pd.notna(row.get('標題',''))     else ''
        content = str(row.get('內容',''))     if pd.notna(row.get('內容',''))     else ''
        sent    = str(row.get('情緒標記','')) if pd.notna(row.get('情緒標記','')) else ''
        dept_id = classify(title, content)
        if sent in dsent[dept_id]: dsent[dept_id][sent] += 1
        news_list.append({
            'title': title,
            'website': str(row.get('網站',''))     if pd.notna(row.get('網站',''))     else '',
            'channel': str(row.get('頻道',''))     if pd.notna(row.get('頻道',''))     else '',
            'date':    row['_date'] if pd.notna(row['_date']) else '',
            'sentiment': sent,
            'url':     str(row.get('原始連結','')) if pd.notna(row.get('原始連結','')) else '',
            'dept':    dept_id,
        })

    # Cluster by dept
    dept_clusters = {}
    for meta in DEPT_META:
        dept_news = [n for n in news_list if n['dept'] == meta['id']]
        if dept_news:
            clusters = cluster_news(dept_news)
            if clusters:
                dept_clusters[meta['id']] = clusters[:10]

    daily = df.groupby('_date').size().reset_index(name='count')
    daily = [{"date": r['_date'], "count": int(r['count'])} for _, r in daily.iterrows() if r['_date'] and str(r['_date']) != 'NaT']
    top_media = df['網站'].value_counts().head(15).reset_index()
    top_media.columns = ['name','count']
    depts_out = []
    for meta in DEPT_META:
        s = dsent[meta["id"]]
        total = s["正面"]+s["中立"]+s["負面"]
        depts_out.append({"id":meta["id"],"name":meta["name"],"short":meta["short"],"color":meta["color"],
                          "total":total,"positive":s["正面"],"neutral":s["中立"],"negative":s["負面"]})

    data = {
        'month': month, 'monthLabel': month_label,
        'total': len(df),
        'positive': int((df['情緒標記']=='正面').sum()),
        'neutral':  int((df['情緒標記']=='中立').sum()),
        'negative': int((df['情緒標記']=='負面').sum()),
        'daily': daily, 'topMedia': top_media.to_dict('records'),
        'depts': depts_out, 'news': news_list, 'deptClusters': dept_clusters,
    }

    out_dir = Path(__file__).parent / 'data'
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / f"{month}.json"
    with open(out_path, 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',',':'))

    manifest_path = out_dir / 'manifest.json'
    manifest = json.loads(manifest_path.read_text()) if manifest_path.exists() else {"months": []}
    if month not in manifest["months"]:
        manifest["months"].append(month)
        manifest["months"].sort()
    manifest_path.write_text(json.dumps(manifest))

    print(f"\n✅ 完成！輸出：data/{month}.json")
    print(f"   總則數：{data['total']}  正面：{data['positive']}  中立：{data['neutral']}  負面：{data['negative']}")
    print("\n📊 各部門露出：")
    for d in depts_out:
        bar = '█' * min(d['total'], 30)
        print(f"   {d['short']:8s} {bar} {d['total']}")
    print("\n🔥 熱門事件群組（Top 3）：")
    for dept_id, clusters in dept_clusters.items():
        meta = next(m for m in DEPT_META if m['id']==dept_id)
        print(f"   [{meta['short']}] {clusters[0]['title'][:30]} ({clusters[0]['count']}則)")
    print("\n🚀 接著執行：")
    print("   git add data/")
    print(f"   git commit -m 'Add {month_label} news data'")
    print("   git push")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法：python convert.py <xlsx檔案路徑>"); sys.exit(1)
    convert(sys.argv[1])
