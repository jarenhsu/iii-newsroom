#!/usr/bin/env python3
"""
資策會新聞月報轉換工具
用法：python convert.py <xlsx檔案1> [檔案2 ...]
支援多檔合併：python convert.py 5月_1.xlsx 5月_2.xlsx 5月_3.xlsx

篩選邏輯學習自人工分類（1-6月），自動過濾無關新聞並分類部門。
關鍵字無法判斷時，自動呼叫 Claude API + 網路搜尋輔助分類。

需要 AI 分類時，請先設定環境變數：
  export ANTHROPIC_API_KEY=你的金鑰
並安裝套件：
  pip install anthropic
"""
import sys, json, re, os
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
    # ── 數轉院：數位轉型、智慧城市、AIoT、無人機、供應鏈平台 ──
    (["digital"],   ["ITS世界大會","ITS World","智慧交通","智慧運輸","智慧城市","智慧城鄉",
                     "AI賦能在地","數位發展部","智慧示範城市","數位治理","AIoT",
                     "創新匯流","創業歸故里","智慧城市展","跨品牌設備",
                     "FinTechSpace","金融科技創新園區","水靈科技","AI主動節能","黑晶",
                     # 新增（學習自5-6月）
                     "無人機.*聯盟","台灣無人機聯盟","智慧次系統.*供應鏈","4SA平台",
                     "CoWiSH","共願開市","愛迪生獎.*資策會","資策會.*愛迪生獎",
                     "G REIGNS.*資策會","資策會.*G REIGNS",
                     "Open RAN.*資策會","資策會.*Open RAN",
                     "Interop Tokyo.*資策會","資策會.*Interop",
                     "大南方.*AI","南部.*AI.*健康","智慧健康展",
                     "AI Infrastructure.*資策會","資策會.*AI Infrastructure",
                     "COMPUTEX.*資策會","資策會.*COMPUTEX",
                     "Edge AI.*聯盟.*資策會","資策會.*Edge AI.*聯盟"]),

    # ── 軟體院：軟體技術、通訊標準驗證、車輛AI驗證 ──
    (["software"],  ["軟體開發","開放原始碼","雲端服務","軟體架構","Spring Boot",
                     "電信軟體",
                     # Open RAN 若沒有 資策會 資訊歸軟體院；若有資策會則上面數轉院先判
                     "DEKRA.*資策會","資策會.*DEKRA","智慧車輛.*AI.*驗證",
                     "STI十大關鍵技術","Agentic AI.*通訊","新通訊.*資策會"]),

    # ── AI院：AI應用開發、智慧系統落地 ──
    (["ai"],        ["人工智慧研究院","AI院",
                     "AI Agent","多模態","LLM","矽光子","CPO.*AI",
                     "全球電信商大會","MWC.*AI",
                     "宏達電.*AI眼鏡","AI眼鏡.*資策會","資策會.*AI眼鏡","HTC.*資策會",
                     "智慧眼鏡.*生態系","AI能量登錄","群電.*資策會","張育誠",
                     # 新增（學習自5-6月）
                     "中科院.*資策會.*無人機","資策會.*中科院.*AI",
                     "拉菲爾.*AI","AI.*手術室","童.*醫院.*AI.*資策會",
                     "智慧瓦斯.*資策會","資策會.*智慧瓦斯",
                     "AI.*認養.*資策會","新北.*AI.*資策會",
                     "技職黑客松.*資策會","資策會.*技職黑客松",
                     "高餐大.*資策會","資策會.*高餐大",
                     "陽明交大.*資策會","資策會.*共鑑慧眼",
                     "永擎.*AI.*資策會","台灣.*AI.*治理.*合規",
                     "宜蘭大學.*AI.*資策會","AWS.*資策會.*AI",
                     "Edge AI.*聯盟.*資策會","資策會.*Edge AI.*聯盟",
                     "青年局.*資策會.*AI院","資策會.*AI院.*MOU",
                     "南台灣.*媒體人才.*資策會","新媒體人才.*資策會"]),

    # ── 資安所：資安防護、CMMC、資安認證 ──
    (["cybersec"],  ["資安署","資通安全署","中小企業資安","資安輔導","資安健檢","資安攻擊",
                     "零信任","資安治理","上市.*資安","網路攻擊",
                     # 新增
                     "CMMC.*資策會","資策會.*CMMC","資策會.*中科院.*資安",
                     "F5.*資策會","NetApp.*Cisco.*資策會","CRA.*資策會"]),

    # ── 教研所：人才培育、AI課程、數位學習 ──
    (["education"], ["就業博覽會","數位人才","先鋒教育","致理科大","人力轉型",
                     "AI Coding","提示工程","課程.*資策會","資策會.*課程","AI.*碩士",
                     "數位學習","職能培育","iPAS","產業新尖兵","勞動部.*資策會",
                     "資策會.*勞動","TIPS.*頒證","文化大學.*資策會","長榮大學.*資策會",
                     # 新增（學習自5-6月）
                     "AI人才.*資策會","資策會.*AI人才",
                     "AI.*認證.*資策會","資策會.*認證.*AI",
                     "育達科大.*資策會","資策會.*育達",
                     "未來人才.*資策會","資策會.*白皮書.*人才",
                     "資策會.*生成式AI.*認證","生成式AI.*人才.*資策會",
                     "1111.*資策會","資策會.*1111"]),

    # ── MIC：產業分析、市場報告、COMPUTEX 產業觀察 ──
    (["mic"],       ["產業情報研究所","MIC","IEK",
                     "產業報告","科技趨勢","MWC.*六大","AI眼鏡.*出貨","智慧眼鏡.*出貨",
                     "十大重點科技","長照.*產業","健康照護.*產業","量子科技.*市場",
                     "AI.*競爭版圖","聯電.*晶圓","半導體.*分析","台達.*光寶","側寫黃欽勇",
                     # 新增（學習自5-6月）
                     "COMPUTEX.*產業.*分析","COMPUTEX.*Best Choice.*資策會",
                     "Grab.*資策會.*市場","外送.*資策會.*分析",
                     "供應鏈.*效率.*資策會","電商.*AI.*資策會",
                     "證交所.*資策會.*創新板","資策會.*看旺.*創新板",
                     "伺服器.*產業.*資策會","AIDC.*資策會",
                     "AI晶片.*資策會","資策會.*AI.*算力.*市場"]),

    # ── 科法所：法律、個資、治理 ──
    (["law"],       ["科技法律研究所","科法所",
                     "資料治理","個人資料","個資","中資.*審查",
                     "闢謠","事實查核","Disinformation","不實資訊","商標.*品牌",
                     "雀巢.*撞名","顏慧欣","科技法律","連鎖加盟","品牌.*合規",
                     "智慧財產.*行銷","歐盟.*經濟安全","中資.*疑慮",
                     # 注意：Grab 相關歸 MIC，不歸科法所
                     "Global CBPR","個資.*驗證.*資策會","資策會.*個資.*驗證",
                     "CGT.*資策會","資策會.*CGT"]),
]

# ── 刪除規則（學習自 1-6 月人工篩選）─────────────────────────────────────────
DELETE_RULES = [
    # 股市/投資平台格式
    (r'^\d{4}[A-Za-z\u4e00-\u9fff]{2,6}[-\s]', '股市格式標題'),
    (r'MoneyDJ新聞\d{4}', '股市格式'),
    (r'CMoney|PChome股市|富聯網|聚財網|旺得富|Cmoney 投資|聚財網-財經新聞', '股市平台'),
    # 求職/面試
    (r'精選面試分享|面試趣', '求職面試'),
    # 政治/社會無關
    (r'柯文哲|黃國昌(?!.*資策會)|李貞秀|陳昭姿|許忠信|陳智菡|韓國瑜|侯友宜', '政治人物'),
    (r'高虹安(?!.*資策會)', '高虹安無關資策會'),
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
    # 低品質媒體
    (r'束褲3C團|SOGI手機王', '低品質媒體'),
    # 其他課程廣告（非資策會）
    (r'【AI必學】Google Sheets Gemini', '他機構課程廣告'),
    # 新增（學習自5-6月）
    (r'大台灣新聞網.*洪毓祥', '低品質轉載'),
    (r'洪毓祥.*兼職.*資策會拔官', '人事敏感'),
    (r'洪毓祥曾因兼職未申報被資策會拔官', '人事敏感'),
    (r'翻爆.*《生醫股》|豐雲學堂.*《生醫股》', '股市轉載格式'),
    (r'蕃新聞.*創新驅動.*數位紅利', '無關轉載'),
    (r'台灣人壽.*Curves.*外溢保單', '無關保險廣告'),
    (r'情傷族.*醫美|失戀醫美', '無關話題'),
    (r'使用「這軟體」聊天注意', '無關話題'),
    (r'公共衛生高峰論壇', '無關活動'),
    (r'愛迪生獎.*資策會|資策會.*愛迪生獎', '獎項不計入露出'),
    (r'《生醫股》.*盛弘|盛弘.*《生醫股》', '股市格式'),
    (r'盛弘醫藥.*資策會.*智慧日照', '企業合作非主要露出'),
    (r'高教薪資結構調整', '無關話題'),
    (r'Edge AI.*聯盟.*資策會.*群聯.*宏碁', '重複轉載格式'),
    (r'XR技術奪愛迪生獎', '獎項不計入'),
]

KEEP_EXCEPTIONS = [
    r'資策會.*資安署|資安署.*資策會',
    r'資策會.*工研院|工研院.*資策會',
    r'資策會.*商研院|商研院.*資策會',
    r'資策會.*中科院|中科院.*資策會',
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
    return None  # 無法判斷，交給 AI

def classify_with_ai(unclassified_rows, api_key=None):
    """
    對關鍵字無法判斷部門的新聞，呼叫 Claude API + 網路搜尋來分類。
    unclassified_rows: list of {'idx': int, 'title': str, 'content': str}
    回傳: {idx: dept_id}
    """
    try:
        import anthropic
    except ImportError:
        print("⚠️  請先安裝：pip install anthropic")
        return {r['idx']: 'hq' for r in unclassified_rows}

    key = api_key or os.environ.get('ANTHROPIC_API_KEY', '')
    if not key:
        print("⚠️  未設定 ANTHROPIC_API_KEY，跳過 AI 分類")
        return {r['idx']: 'hq' for r in unclassified_rows}

    client = anthropic.Anthropic(api_key=key)
    results = {}
    BATCH = 15

    SYSTEM = """你是資策會新聞分類專家。以下是資策會的部門及其業務範疇：

- digital（數位轉型研究院/數轉院）：數位轉型、智慧城市、AIoT、無人機、供應鏈平台、智慧交通、ITS、COMPUTEX展覽活動、國際合作展覽
- software（軟體技術研究院/軟體院）：軟體技術、通訊標準、車輛AI驗證、DEKRA合作、ITS專案辦公室、Open RAN（若資策會主導）
- ai（人工智慧研究院/AI院）：AI應用開發與落地、智慧系統整合、Edge AI聯盟、AI認證合作、產學MOU
- cybersec（資安科技研究所/資安所）：資安防護、CMMC認證、資安輔導、資安治理
- education（數位教育研究所/教研所）：人才培育、AI課程認證、數位學習、產學合作教育
- mic（產業情報研究所/MIC）：產業分析報告、市場研究、科技趨勢預測、COMPUTEX產業觀察、證交所合作分析
- law（科技法律研究所/科法所）：科技法律、資料治理、個資保護、事實查核、商標品牌
- hq（其他/機構整體）：資策會整體形象、跨部門活動、無法明確歸屬特定部門

請利用網路搜尋查詢新聞相關資訊，判斷每則新聞最可能屬於哪個部門。

回傳純 JSON 陣列，格式：[{"idx": 0, "dept": "mic", "reason": "MIC產業分析"}]
不要包含任何說明文字或 markdown。"""

    for i in range(0, len(unclassified_rows), BATCH):
        chunk = unclassified_rows[i:i+BATCH]
        batch_num = i//BATCH + 1
        total_batches = (len(unclassified_rows)-1)//BATCH + 1
        print(f"  🤖 AI 分類 batch {batch_num}/{total_batches} ({len(chunk)} 則)...", end=' ', flush=True)

        items = [{"idx": r['idx'], "title": r['title'], "content": r['content'][:200]} for r in chunk]
        user_msg = f"請搜尋並分類以下新聞（回傳純 JSON）：\n{json.dumps(items, ensure_ascii=False)}"

        try:
            resp = client.messages.create(
                model="claude-sonnet-4-6",
                max_tokens=1000,
                system=SYSTEM,
                tools=[{"type": "web_search_20250305", "name": "web_search"}],
                messages=[{"role": "user", "content": user_msg}]
            )
            # 取出所有 text 內容
            text = ' '.join(b.text for b in resp.content if hasattr(b, 'text') and b.text)
            # 清理 JSON
            text = text.strip()
            if '```' in text:
                text = text.split('```')[1]
                if text.startswith('json'): text = text[4:]
            parsed = json.loads(text.strip())
            for item in parsed:
                results[item['idx']] = item.get('dept', 'hq')
            print(f"✓ ({len(parsed)} 則)")
        except Exception as e:
            print(f"✗ ({e})")
            for r in chunk:
                results[r['idx']] = 'hq'

        import time; time.sleep(0.5)

    return results

def ngrams(text, n=4):
    t = str(text); return set(t[i:i+n] for i in range(len(t)-n+1))

def sim(a, b):
    sa, sb = ngrams(a), ngrams(b)
    if not sa or not sb: return 0
    return len(sa & sb) / len(sa | sb)

# 明確轉載標記（內文中有這些 → 表示該篇是轉載自別家媒體，直接刪除）
REPOST_MARKERS = [
    r'(?:本文|全文)?轉載自\s*[《〔\[]?[\u4e00-\u9fff\w]+[》〕\]]?',
    r'資料來源[：:]\s*[\u4e00-\u9fff\w]{2,15}(?:新聞|報|網|社|時報)',
    r'原文(?:刊載|出處|來源)[：:於]?\s*[\u4e00-\u9fff\w]{2,15}',
]

def is_repost(content):
    """
    判斷該篇新聞是否為轉載。
    只有內文中明確標示來源為別家媒體，才視為轉載並刪除。
    若內文和其他媒體一樣但沒有標明轉載，則視為獨立發布，保留。
    """
    if not content:
        return False
    text = content[:400]  # 只看前400字，來源標記通常在文首或文尾
    for pat in REPOST_MARKERS:
        if re.search(pat, text):
            return True
    return False

def dedup(rows, threshold=0.85):
    """
    去重邏輯：
    - 只刪除內文有明確標示「轉載自某媒體」的新聞
    - 相同內容但各自獨立發布（新聞稿）的，全部保留
    - 沒有內文時用標題相似度（>85%）作為 fallback
    """
    kept = []
    for row in rows:
        content = row.get('content', '')
        if is_repost(content):
            continue  # 明確標示轉載 → 刪除
        # 無明確轉載標記：保留（即使內容和別篇一樣，視為獨立發布）
        # 但若完全沒有內文，用標題相似度避免完全空白的重複
        if not content.strip():
            if any(sim(row['title'], k['title']) >= threshold for k in kept):
                continue
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
                          'date':date,'sentiment':sent,'url':url,
                          'dept':classify(title,content)})  # None = 待 AI 分類

    rows_final = dedup(rows_pass, threshold=0.85)

    # AI 輔助分類：對關鍵字無法判斷（dept=None）的新聞，呼叫 Claude API + 網路搜尋
    unclassified = [{'idx': i, 'title': r['title'], 'content': r['content']}
                    for i, r in enumerate(rows_final) if r['dept'] is None]

    if unclassified:
        api_key = os.environ.get('ANTHROPIC_API_KEY', '')
        if api_key:
            print(f"🔍 關鍵字無法判斷：{len(unclassified)} 則，啟動 AI 網路搜尋分類...")
            ai_results = classify_with_ai(unclassified, api_key)
            for i, r in enumerate(rows_final):
                if r['dept'] is None:
                    r['dept'] = ai_results.get(i, 'hq')
            ai_count = sum(1 for r in rows_final if r['dept'] != 'hq')
            print(f"   AI 分類完成，{len(unclassified)} 則中成功判斷部門")
        else:
            print(f"⚠️  {len(unclassified)} 則無法判斷部門（設定 ANTHROPIC_API_KEY 以啟用 AI 分類）")
            for r in rows_final:
                if r['dept'] is None: r['dept'] = 'hq'
    else:
        # 確保沒有 None
        for r in rows_final:
            if r['dept'] is None: r['dept'] = 'hq'

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
