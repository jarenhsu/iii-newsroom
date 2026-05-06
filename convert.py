#!/usr/bin/env python3
"""
資策會新聞月報轉換工具
用法：python convert.py <xlsx檔案1> [檔案2 ...]
支援多檔合併：python convert.py 5月_1.xlsx 5月_2.xlsx 5月_3.xlsx

篩選邏輯學習自人工分類（1-4月），自動過濾無關新聞並分類部門。
"""
import sys, json, re
from pathlib import Path
from collections import defaultdict, Counter
import pandas as pd

# ── 部門定義 ─────────────────────────────────────────────────────────────────
DEPT_META = [
    {"id":"digital",  "name":"數位轉型研究院","short":"數轉院","color":"#5ab878"},
    {"id":"software", "name":"軟體技術研究院","short":"軟體院","color":"#6aaa8a"},
    {"id":"ai",       "name":"人工智慧研究院","short":"AI院",  "color":"#a8c860"},
    {"id":"cybersec", "name":"資安科技研究所","short":"資安所","color":"#c87a50"},
    {"id":"education","name":"數位教育研究所","short":"教研所","color":"#d4b84a"},
    {"id":"mic",      "name":"產業情報研究所","short":"MIC",   "color":"#b87a30"},
    {"id":"law",      "name":"科技法律研究所","short":"科法所","color":"#8a6ab8"},
    {"id":"hq",       "name":"其他/機構整體", "short":"其他",  "color":"#6a8060"},
]

DEPTS_KW = [
    {"id":"digital",   "patterns":["數位轉型研究院","數轉院"]},
    {"id":"software",  "patterns":["軟體技術研究院","軟體院"]},
    {"id":"ai",        "patterns":["人工智慧研究院","AI院","ai院","AI研究院"]},
    {"id":"cybersec",  "patterns":["資安科技研究所","資安所"]},
    {"id":"education", "patterns":["數位教育研究所","教研所","數位教育所"]},
    {"id":"mic",       "patterns":["產業情報研究所","MIC","IEK"]},
    {"id":"law",       "patterns":["科技法律研究所","科法所"]},
]

TOPIC_RULES = [
    (["digital"],   ["ITS世界大會","ITS World","智慧交通","智慧運輸","智慧城市","智慧城鄉",
                     "AI賦能在地","數位發展部","智慧示範城市","數位治理","AIoT",
                     "創新匯流","創業歸故里","智慧城市展","跨品牌設備",
                     "FinTechSpace","金融科技創新園區","水靈科技","AI主動節能","黑晶"]),
    (["software"],  ["軟體開發","開放原始碼","雲端服務","軟體架構","Spring Boot",
                     "AI.*RAN","開放網路","電信軟體","Open RAN"]),
    (["ai"],        ["AI Agent","多模態","LLM","生成式AI","AI.*算力","矽光子","CPO.*AI",
                     "製造業AI","全球電信商大會","MWC.*AI","宏達電.*AI眼鏡","AI眼鏡.*資策會",
                     "資策會.*AI眼鏡","HTC.*資策會","智慧眼鏡.*生態系","AI能量登錄",
                     "群電.*資策會","無人機.*資策會","張育誠"]),
    (["cybersec"],  ["資安署","資通安全署","中小企業資安","資安輔導","資安健檢","資安攻擊",
                     "零信任","資安治理","上市.*資安","網路攻擊"]),
    (["education"], ["就業博覽會","AI人才","數位人才","先鋒教育","致理科大","人力轉型",
                     "AI Coding","提示工程","課程.*資策會","資策會.*課程","AI.*碩士",
                     "數位學習","職能培育","iPAS","產業新尖兵","勞動部.*資策會",
                     "資策會.*勞動","TIPS.*頒證","文化大學.*資策會","長榮大學.*資策會"]),
    (["mic"],       ["產業報告","科技趨勢","MWC.*六大","AI眼鏡.*出貨","智慧眼鏡.*出貨",
                     "十大重點科技","長照.*產業","健康照護.*產業","量子科技.*市場",
                     "AI.*競爭版圖","聯電.*晶圓","半導體.*分析","台達.*光寶","側寫黃欽勇"]),
    (["law"],       ["資料治理","個人資料","個資","Grab.*foodpanda","中資.*審查",
                     "闢謠","事實查核","Disinformation","不實資訊","商標.*品牌",
                     "雀巢.*撞名","顏慧欣","科技法律","連鎖加盟","品牌.*合規",
                     "智慧財產.*行銷","歐盟.*經濟安全","中資.*疑慮"]),
]

# ── 刪除規則（學習自 1-4 月人工篩選）────────────────────────────────────────
DELETE_RULES = [
    # 股市/投資平台格式
    (r'^\d{4}[A-Za-z\u4e00-\u9fff]{2,6}[-\s]', '股市格式標題'),
    (r'MoneyDJ新聞\d{4}', '股市格式'),
    (r'CMoney|PChome股市|富聯網|聚財網|旺得富|Cmoney 投資', '股市平台'),
    # 求職/面試
    (r'精選面試分享|面試趣', '求職面試'),
    # 政治/社會無關
    (r'高虹安|柯文哲|黃國昌|李貞秀|陳昭姿|許忠信|陳智菡|韓國瑜|侯友宜', '政治人物'),
    (r'立委.*宣誓就職|新科立委|兩年條款|直言最難溝通|對綠白合作表態', '政治事件'),
    # 其他機構主角（非資策會）
    (r'^(?=.*勤業眾信)(?!.*資策會)', '勤業眾信'),
    (r'^(?=.*商研院)(?!.*資策會)', '商研院'),
    (r'^(?=.*工研院)(?!.*資策會)', '工研院'),
    (r'台灣新聞通訊社-中聯資源', '無關廠商'),
    # 資策會只是小字帶到
    (r'跟著中科院.*就對了', '無關引用'),
    (r'FinTechSpace.*(?:尾牙|迎新春|餐飲|慶祝)', '內部活動'),
    (r'即時新聞】攜手資策會建構AI眼鏡生態系$', '無內容'),
    # 無關科技話題
    (r'月球.*(?:飯店|核反應器)|核能登月', '無關話題'),
    # 低品質媒體轉載（無新聞價值）
    (r'束褲3C團|SOGI手機王', '低品質媒體'),
    # 其他課程廣告（非資策會）
    (r'【AI必學】Google Sheets Gemini', '他機構課程廣告'),
]

KEEP_EXCEPTIONS = [
    r'資策會.*資安署|資安署.*資策會',
    r'資策會.*工研院|工研院.*資策會',
    r'資策會.*商研院|商研院.*資策會',
]

# ── 核心函式 ─────────────────────────────────────────────────────────────────
def should_delete(title, content):
    text = title + content
    for ex in KEEP_EXCEPTIONS:
        if re.search(ex, text): return False
    if '資策會' not in text:
        has_kw = any(any(p in text for p in d["patterns"]) for d in DEPTS_KW)
        has_topic = any(any(re.search(kw, text) for kw in kws) for _, kws in TOPIC_RULES)
        if not has_kw and not has_topic: return True
    for pattern, _ in DELETE_RULES:
        if re.search(pattern, title): return True
    return False

def classify(title, content):
    text = str(title) + str(content)
    for d in DEPTS_KW:
        for p in d["patterns"]:
            if p in text: return d["id"]
    for dept_ids, keywords in TOPIC_RULES:
        for kw in keywords:
            if re.search(kw, text): return dept_ids[0]
    return "hq"

def ngrams(text, n=4):
    t = str(text); return set(t[i:i+n] for i in range(len(t)-n+1))

def sim(a, b):
    sa, sb = ngrams(a), ngrams(b)
    if not sa or not sb: return 0
    return len(sa & sb) / len(sa | sb)

def dedup(rows, threshold=0.85):
    kept = []
    for row in rows:
        if not any(sim(row['title'], k['title']) >= threshold for k in kept):
            kept.append(row)
    return kept

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
            if sim(titles[i], titles[j]) >= threshold: union(i, j)
    groups = defaultdict(list)
    for i in range(n): groups[find(i)].append(i)
    clusters = []
    for idxs in sorted(groups.values(), key=lambda c: -len(c)):
        if len(idxs) < 2: continue
        rep = sorted([titles[i] for i in idxs if titles[i]], key=len)[0]
        clusters.append({"title": rep, "count": len(idxs)})
    return clusters

# ── 主程式 ───────────────────────────────────────────────────────────────────
def convert(*xlsx_paths):
    # 合併所有檔案
    dfs = []
    for p in xlsx_paths:
        path = Path(p)
        if not path.exists():
            print(f"❌ 找不到：{p}"); continue
        dfs.append(pd.read_excel(path, sheet_name=0))
        print(f"📖 讀取 {path.name}（{len(dfs[-1])} 則）")
    if not dfs: sys.exit(1)

    df = pd.concat(dfs, ignore_index=True)
    df['date'] = pd.to_datetime(df['發布時間'], errors='coerce').dt.strftime('%Y-%m-%d')
    df = df.sort_values('date').reset_index(drop=True)

    # 判斷月份
    first_path = Path(xlsx_paths[0])
    m = re.search(r'(\d{4})-(\d{2})', first_path.name)
    if m:
        month = f"{m.group(1)}-{m.group(2)}"
    else:
        valid = df['date'].dropna()
        month = sorted(valid)[0][:7] if len(valid) else __import__('datetime').date.today().strftime('%Y-%m')

    year, mon = month.split('-')
    month_label = f"{year}年{int(mon)}月"
    print(f"\n📅 月份：{month_label}  原始：{len(df)} 則")

    # 過濾
    rows_pass = []
    del_stats = Counter()
    for _, row in df.iterrows():
        title   = str(row.get('標題',''))     if pd.notna(row.get('標題',''))     else ''
        content = str(row.get('內容',''))     if pd.notna(row.get('內容',''))     else ''
        sent    = str(row.get('情緒標記','')) if pd.notna(row.get('情緒標記','')) else ''
        url     = str(row.get('原始連結','')) if pd.notna(row.get('原始連結','')) else ''
        website = str(row.get('網站',''))     if pd.notna(row.get('網站',''))     else ''
        date    = row['date'] if pd.notna(row['date']) else ''
        if should_delete(title, content):
            del_stats['deleted'] += 1
            continue
        rows_pass.append({'title':title,'content':content,'website':website,
                          'date':date,'sentiment':sent,'url':url,'dept':classify(title,content)})

    rows_final = dedup(rows_pass, threshold=0.85)
    print(f"✅ 過濾後：{len(rows_pass)} 則  去重後：{len(rows_final)} 則（刪除 {del_stats['deleted']} 則）")

    # 組裝 JSON
    dsent = {d["id"]: {"正面":0,"中立":0,"負面":0} for d in DEPT_META}
    news_list = []
    for r in rows_final:
        dept_id = r['dept']; sent = r['sentiment']
        if sent in dsent[dept_id]: dsent[dept_id][sent] += 1
        news_list.append({'title':r['title'],'website':r['website'],'channel':'',
                          'date':r['date'],'sentiment':sent,'url':r['url'],'dept':dept_id})

    dept_clusters = {}
    for meta in DEPT_META:
        dn = [n for n in news_list if n['dept']==meta['id']]
        if dn:
            cl = cluster_news(dn)
            if cl: dept_clusters[meta['id']] = cl[:10]

    daily = [{"date":d,"count":c} for d,c in sorted(Counter(r['date'] for r in rows_final).items())]
    top_media = [{"name":k,"count":v} for k,v in Counter(r['website'] for r in rows_final).most_common(15)]
    depts_out = []
    pos = sum(1 for r in rows_final if r['sentiment']=='正面')
    neu = sum(1 for r in rows_final if r['sentiment']=='中立')
    neg = sum(1 for r in rows_final if r['sentiment']=='負面')
    for meta in DEPT_META:
        s = dsent[meta["id"]]; total = s["正面"]+s["中立"]+s["負面"]
        depts_out.append({"id":meta["id"],"name":meta["name"],"short":meta["short"],"color":meta["color"],
                          "total":total,"positive":s["正面"],"neutral":s["中立"],"negative":s["負面"]})

    data = {
        'month':month,'monthLabel':month_label,'total':len(rows_final),
        'positive':pos,'neutral':neu,'negative':neg,
        'daily':daily,'topMedia':top_media,'depts':depts_out,
        'news':news_list,'deptClusters':dept_clusters,
    }

    out_dir = Path(__file__).parent / 'data'
    out_dir.mkdir(exist_ok=True)
    out_path = out_dir / f"{month}.json"
    with open(out_path,'w',encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, separators=(',',':'))

    manifest_path = out_dir / 'manifest.json'
    manifest = json.loads(manifest_path.read_text()) if manifest_path.exists() else {"months":[]}
    if month not in manifest["months"]:
        manifest["months"].append(month); manifest["months"].sort()
    manifest_path.write_text(json.dumps(manifest))

    print(f"\n📊 各部門露出：")
    for d in depts_out:
        if d['total']>0:
            bar = '█' * min(d['total']//3, 25)
            print(f"   {d['short']:6s} {bar} {d['total']}")
    print(f"\n   正面:{pos}  中立:{neu}  負面:{neg}")
    print(f"\n🚀 接著執行：")
    print(f"   git add data/")
    print(f"   git commit -m 'Add {month_label} news data'")
    print(f"   git push")

if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("用法：python convert.py <xlsx1> [xlsx2 ...]")
        print("範例：python convert.py 5月_1.xlsx 5月_2.xlsx 5月_3.xlsx")
        sys.exit(1)
    convert(*sys.argv[1:])
