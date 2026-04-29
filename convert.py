#!/usr/bin/env python3
"""
資策會新聞月報轉換工具
用法：python convert.py Opview_Insight_Result_*.xlsx
需要安裝：pip install pandas openpyxl anthropic
"""
import sys, json, re, time
from pathlib import Path
from collections import defaultdict
import pandas as pd

# ── 部門定義 ────────────────────────────────────────────────────────────────
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
    {"id":"digital",  "name":"數位轉型研究院","short":"數轉院","color":"#5ec49a"},
    {"id":"software", "name":"軟體技術研究院","short":"軟體院","color":"#6098c8"},
    {"id":"ai",       "name":"人工智慧研究院","short":"AI院",  "color":"#8870c8"},
    {"id":"cybersec", "name":"資安科技研究所","short":"資安所","color":"#d07070"},
    {"id":"education","name":"數位教育研究所","short":"教研所","color":"#d4a830"},
    {"id":"mic",      "name":"產業情報研究所","short":"MIC",   "color":"#d07848"},
    {"id":"law",      "name":"科技法律研究所","short":"科法所","color":"#b868c0"},
    {"id":"hq",       "name":"總部/機構整體", "short":"總部",  "color":"#7a90a8"},
]

DEPT_DESC = """
部門清單與業務範疇（用於判斷新聞歸屬）：
- digital（數位轉型研究院）：數位轉型、智慧城市、智慧城鄉、數位治理、AIoT、智慧製造、數位發展政策、智慧交通、ITS世界大會、資策會助攻智慧交通
- software（軟體技術研究院）：軟體開發、雲端服務、軟體架構、系統整合、開放網路、RAN、電信軟體、資策會軟體院
- ai（人工智慧研究院）：AI模型、機器學習、LLM、生成式AI、AI Agent、自然語言處理、電腦視覺、多模態AI、AI應用開發
- cybersec（資安科技研究所）：資訊安全、網路攻擊、資安防護、零信任、資安署、中小企業資安、資安輔導、個資保護
- education（數位教育研究所）：數位教育、人才培訓、AI課程、數位學習、就業博覽會、AI人才培育、職能發展
- mic（產業情報研究所）：產業分析、市場研究、科技趨勢、MWC報告、半導體分析、產業報告、投資情報、IEK
- law（科技法律研究所）：科技法律、資料治理、個資法、AI法規、數位法規、商標著作權、合規、Disinformation、事實查核
- hq（總部/機構整體）：資策會整體形象、執行長活動、整體合作備忘錄、無法明確對應特定部門、純引用資策會名稱而無具體業務
"""

# ── 精確關鍵字分類 ───────────────────────────────────────────────────────────
def classify_exact(title, content):
    text = str(title) + str(content)
    for d in DEPTS:
        for p in d["patterns"]:
            if p in text:
                return d["id"]
    return None  # 需要 AI 判斷

# ── 標題相似度（用於 AI 結果傳播） ─────────────────────────────────────────
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

# ── AI 批次分類 ──────────────────────────────────────────────────────────────
def ai_classify_batch(items, client):
    """用 Claude API 分類一批新聞，回傳 {idx: dept_id} dict"""
    import anthropic

    system = f"""你是資策會新聞分類專家。根據新聞標題和內文摘要，判斷每則新聞最可能歸屬哪個部門。

{DEPT_DESC}

分類規則：
1. 優先看主題和技術領域，而非表面關鍵字
2. 同事件（如ITS世界大會）歸同一部門
3. 確實無法判斷才用 hq
4. 只回傳純 JSON 陣列，格式：[{{"idx":0,"dept":"cybersec"}}]
5. 不要包含任何解釋或 markdown"""

    user_msg = "請分類以下新聞（回傳純 JSON）：\n" + json.dumps(items, ensure_ascii=False)

    resp = client.messages.create(
        model="claude-sonnet-4-20250514",
        max_tokens=1000,
        system=system,
        messages=[{"role": "user", "content": user_msg}]
    )
    text = resp.content[0].text.strip()
    if '```' in text:
        text = text.split('```')[1]
        if text.startswith('json'): text = text[4:]
    parsed = json.loads(text.strip())
    return {item['idx']: item['dept'] for item in parsed}

def ai_classify_all(unclassified_rows, client):
    """分批呼叫 AI 分類所有未分類新聞"""
    batch = [
        {"idx": i, "title": str(r['title']), "content": str(r['content'])[:250]}
        for i, r in enumerate(unclassified_rows)
    ]

    results = {}
    BATCH_SIZE = 40
    total_batches = (len(batch) - 1) // BATCH_SIZE + 1

    for b_idx in range(0, len(batch), BATCH_SIZE):
        chunk = batch[b_idx:b_idx + BATCH_SIZE]
        print(f"  🤖 AI 分類 batch {b_idx//BATCH_SIZE + 1}/{total_batches} ({len(chunk)} 則)...", end=' ', flush=True)
        try:
            classified = ai_classify_batch(chunk, client)
            results.update(classified)
            print(f"✓")
        except Exception as e:
            print(f"✗ ({e})")
            for item in chunk:
                results[item['idx']] = 'hq'
        time.sleep(0.3)

    return results

# ── 主程式 ───────────────────────────────────────────────────────────────────
def convert(xlsx_path: str, use_ai: bool = True):
    path = Path(xlsx_path)
    if not path.exists():
        print(f"❌ 找不到檔案：{xlsx_path}"); sys.exit(1)

    print(f"📖 讀取 {path.name} ...")
    df = pd.read_excel(path, sheet_name=0)
    df['_date'] = pd.to_datetime(df['發布時間'], errors='coerce').dt.strftime('%Y-%m-%d')

    # 判斷月份
    m = re.search(r'(\d{4})-(\d{2})', path.name)
    if m:
        month = f"{m.group(1)}-{m.group(2)}"
    else:
        valid = df['_date'].dropna()
        month = sorted(valid)[0][:7] if len(valid) else __import__('datetime').date.today().strftime('%Y-%m')

    year, mon = month.split('-')
    month_label = f"{year}年{int(mon)}月"
    print(f"📅 月份：{month_label}")

    # ── 步驟1：精確關鍵字分類 ──
    rows = []
    for _, row in df.iterrows():
        title   = str(row.get('標題',''))     if pd.notna(row.get('標題',''))     else ''
        content = str(row.get('內容',''))     if pd.notna(row.get('內容',''))     else ''
        sent    = str(row.get('情緒標記','')) if pd.notna(row.get('情緒標記','')) else ''
        date    = row['_date'] if pd.notna(row['_date']) else ''
        dept_id = classify_exact(title, content)
        rows.append({
            'title': title, 'content': content,
            'website': str(row.get('網站',''))     if pd.notna(row.get('網站',''))     else '',
            'channel': str(row.get('頻道',''))     if pd.notna(row.get('頻道',''))     else '',
            'date': date, 'sentiment': sent,
            'url':  str(row.get('原始連結',''))    if pd.notna(row.get('原始連結','')) else '',
            'dept': dept_id,  # None = 待 AI 分類
        })

    exact_count = sum(1 for r in rows if r['dept'] is not None)
    unclassified = [r for r in rows if r['dept'] is None]
    print(f"✅ 精確分類：{exact_count} 則  ⏳ 待 AI 分類：{len(unclassified)} 則")

    # ── 步驟2：AI 分類（標題相似度傳播） ──
    if unclassified and use_ai:
        try:
            import anthropic
            client = anthropic.Anthropic()  # 讀取環境變數 ANTHROPIC_API_KEY
            print("🤖 啟動 AI 智慧分類...")

            # 先用標題相似度：若未分類的新聞和已分類的非常相似，直接繼承
            propagated = 0
            for r in unclassified:
                best_sim, best_dept = 0, None
                for cr in rows:
                    if cr['dept'] is None: continue
                    s = similarity(r['title'], cr['title'])
                    if s > best_sim:
                        best_sim, best_dept = s, cr['dept']
                if best_sim >= 0.35:
                    r['dept'] = best_dept
                    propagated += 1

            still_unclassified = [r for r in rows if r['dept'] is None]
            print(f"  📎 相似度傳播：{propagated} 則  ⏳ 剩餘 AI 分類：{len(still_unclassified)} 則")

            # 剩餘的送 Claude API
            if still_unclassified:
                ai_results = ai_classify_all(still_unclassified, client)
                for i, r in enumerate(still_unclassified):
                    r['dept'] = ai_results.get(i, 'hq')

        except ImportError:
            print("⚠️  未安裝 anthropic 套件，跳過 AI 分類（pip install anthropic）")
            for r in unclassified:
                r['dept'] = 'hq'
        except Exception as e:
            print(f"⚠️  AI 分類失敗：{e}，改用總部分類")
            for r in unclassified:
                if r['dept'] is None:
                    r['dept'] = 'hq'
    else:
        for r in rows:
            if r['dept'] is None:
                r['dept'] = 'hq'

    # ── 步驟3：統計 ──
    dsent = {d["id"]: {"正面":0,"中立":0,"負面":0} for d in DEPT_META}
    news_list = []
    for r in rows:
        dept_id = r['dept'] or 'hq'
        sent = r['sentiment']
        if sent in dsent.get(dept_id, {}):
            dsent[dept_id][sent] += 1
        news_list.append({
            'title': r['title'], 'website': r['website'], 'channel': r['channel'],
            'date': r['date'], 'sentiment': sent, 'url': r['url'], 'dept': dept_id,
        })

    # 部門群組
    dept_clusters = {}
    for meta in DEPT_META:
        dept_news = [n for n in news_list if n['dept'] == meta['id']]
        if dept_news:
            clusters = cluster_news(dept_news)
            if clusters:
                dept_clusters[meta['id']] = clusters[:10]

    daily = df.groupby('_date').size().reset_index(name='count')
    daily = [{"date": r['_date'], "count": int(r['count'])} for _, r in daily.iterrows()
             if r['_date'] and str(r['_date']) != 'NaT']
    top_media = df['網站'].value_counts().head(15).reset_index()
    top_media.columns = ['name','count']

    depts_out = []
    for meta in DEPT_META:
        s = dsent[meta["id"]]
        total = s["正面"]+s["中立"]+s["負面"]
        depts_out.append({"id":meta["id"],"name":meta["name"],"short":meta["short"],"color":meta["color"],
                          "total":total,"positive":s["正面"],"neutral":s["中立"],"negative":s["負面"]})

    data = {
        'month': month, 'monthLabel': month_label, 'total': len(df),
        'positive': int((df['情緒標記']=='正面').sum()),
        'neutral':  int((df['情緒標記']=='中立').sum()),
        'negative': int((df['情緒標記']=='負面').sum()),
        'daily': daily, 'topMedia': top_media.to_dict('records'),
        'depts': depts_out, 'news': news_list, 'deptClusters': dept_clusters,
    }

    # ── 輸出 ──
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
        bar = '█' * min(d['total']//2, 30)
        print(f"   {d['short']:8s} {bar} {d['total']}")
    print("\n🚀 接著執行：")
    print("   git add data/")
    print(f"   git commit -m 'Add {month_label} news data'")
    print("   git push")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法：python convert.py <xlsx檔案路徑>")
        print("      設定環境變數 ANTHROPIC_API_KEY 以啟用 AI 分類")
        sys.exit(1)
    convert(sys.argv[1])
