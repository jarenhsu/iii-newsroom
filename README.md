# 資策會新聞露出月報

**網址：** https://jarenhsu.github.io/iii-newsroom

每月新聞媒體露出數據分析，依部門分類呈現。

---

## 首次設定（只需做一次）

### 1. 建立 GitHub Repo

1. 登入 GitHub，點右上角 **+** → **New repository**
2. Repository name 填：`iii-newsroom`
3. 設為 **Public**
4. 不要勾選任何初始化選項，直接點 **Create repository**

### 2. 把這個資料夾推上去

在這個資料夾內開終端機，執行：

```bash
git init
git add .
git commit -m "Initial commit"
git branch -M main
git remote add origin https://github.com/jarenhsu/iii-newsroom.git
git push -u origin main
```

### 3. 開啟 GitHub Pages

1. 進入 repo 頁面 → **Settings** → 左側 **Pages**
2. Source 選 **Deploy from a branch**
3. Branch 選 **main**，資料夾選 **/ (root)**
4. 點 **Save**

等 1–2 分鐘，網站就會在 https://jarenhsu.github.io/iii-newsroom 上線。

---

## 每月更新流程（兩步驟）

### 步驟一：執行轉換腳本

```bash
python convert.py Opview_Insight_Result_臨時新聞事件_2026-05-31.xlsx
```

腳本會自動產生 `data/2026-05.json` 並更新 `data/manifest.json`。

### 步驟二：推上 GitHub

```bash
git add data/
git commit -m "Add 2026年5月 news data"
git push
```

推完後約 1 分鐘，網站自動更新。

---

## 檔案結構

```
iii-newsroom/
├── index.html          # 網站主體（不需修改）
├── convert.py          # 每月轉換腳本
├── README.md
└── data/
    ├── manifest.json   # 月份清單（自動維護）
    ├── 2026-03.json
    └── ...
```

---

## 部門分類關鍵字

| 部門 | 識別關鍵字 |
|------|-----------|
| 數位轉型研究院 | 數位轉型研究院、數轉院 |
| 軟體技術研究院 | 軟體技術研究院、軟體院 |
| 人工智慧研究院 | 人工智慧研究院、AI院、AI研究院 |
| 資安科技研究所 | 資安科技研究所、資安所 |
| 數位教育研究所 | 數位教育研究所、教研所、數位教育所 |
| 產業情報研究所 | 產業情報研究所、MIC、IEK |
| 科技法律研究所 | 科技法律研究所、科法所 |

---

## 環境需求

```bash
pip install pandas openpyxl
```
