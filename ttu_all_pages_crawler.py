import os
import re
import time
import random
import json
from datetime import datetime
from urllib.parse import urljoin, urlparse, unquote
from pathlib import Path

import requests
from bs4 import BeautifulSoup

# ====== 附件讀取套件 ======
try:
    from pypdf import PdfReader
except:
    PdfReader = None
try:
    from docx import Document
except:
    Document = None
try:
    from odf.opendocument import load as odf_load
    from odf import text as odf_text, teletype
    HAS_ODF = True
except:
    HAS_ODF = False
try:
    from openpyxl import load_workbook
except:
    load_workbook = None
try:
    from pptx import Presentation
except:
    Presentation = None

# ==================== 設定 ====================
SAVE_ROOT = Path("ttu_all_pages")
ATTACH_ROOT = SAVE_ROOT / "attachments"
HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0.0.0 Safari/537.36",
    "Accept-Language": "zh-TW,zh;q=0.9",
}
SESSION = requests.Session()
SESSION.headers.update(HEADERS)

# 從 2026 年 8 月 10 日到今天
DATE_START = datetime(2026, 8, 10)
DATE_END = datetime.now()

# 固定頁（沒日期也保留）
KEEP_ALWAYS_KEYWORDS = ["職掌", "執掌", "法規", "規定", "申請", "須知", "說明", "簡介", "流程", "Q&A", "地圖", "統計"]

# ==================== 目標清單 ====================
TARGETS = {
    "處本部": {
        "人員職掌": "https://dean.ttu.edu.tw/p/412-1002-87.php",
        "會議記錄": "https://dean.ttu.edu.tw/p/412-1002-88.php",
        "學分學程": "https://dean.ttu.edu.tw/p/412-1002-90.php",
        "第三週期校務評鑑": "https://dean.ttu.edu.tw/p/412-1002-2911.php",
        "彈性教學週實施說明": "https://curri.ttu.edu.tw/var/file/33/1033/img/651621143.pdf",
        "教務相關法規": "https://tturule.ttu.edu.tw/rulelist/index.php?unit=A1100",
    },
    "註冊課務組_註冊事務": {
        "人員職掌": "https://reg.ttu.edu.tw/p/412-1032-905.php",
        "休退復學申請": "https://reg.ttu.edu.tw/p/412-1032-907.php",
        "教務處各項證明文件申請": "https://reg.ttu.edu.tw/p/412-1032-908.php",
        "行事曆": "https://reg.ttu.edu.tw/p/412-1032-910.php",
        "碩博士班論文口試申請": "https://reg.ttu.edu.tw/p/412-1032-911.php",
        "學位論文專業符合機制規定": "https://reg.ttu.edu.tw/p/412-1032-1816.php",
        "學位授予專區_相關法規": "https://reg.ttu.edu.tw/p/412-1032-922.php",
        "學術倫理_相關法規": "https://reg.ttu.edu.tw/p/412-1032-924.php",
        "學術倫理_學生修習學術研究倫理教育課": "https://reg.ttu.edu.tw/p/412-1032-925.php",
        "學術倫理_公告": "https://reg.ttu.edu.tw/p/412-1032-926.php",
        "悠遊卡學生證使用說明": "https://reg.ttu.edu.tw/p/412-1032-914.php",
        "轉系輔系雙主修申請": "https://reg.ttu.edu.tw/p/412-1032-920.php",
        "檔案下載": "https://reg.ttu.edu.tw/p/412-1032-921.php",
        "碩士班逕修讀博士學位": "https://reg.ttu.edu.tw/p/412-1032-2515.php",
        "數位學位證書": "https://reg.ttu.edu.tw/p/412-1032-2800.php",
    },
    "註冊課務組_課務事務": {
        "最新消息": "https://curri.ttu.edu.tw/p/412-1033-436.php",
        "法規_選課修課": "https://curri.ttu.edu.tw/p/412-1033-1235.php",
        "法規_考試": "https://curri.ttu.edu.tw/p/412-1033-1236.php",
        "法規_教學助理": "https://curri.ttu.edu.tw/p/412-1033-1237.php",
        "法規_課程": "https://curri.ttu.edu.tw/p/412-1033-1238.php",
        "法規_鐘點費": "https://curri.ttu.edu.tw/p/412-1033-1239.php",
        "選課Q&A": "https://curri.ttu.edu.tw/p/412-1033-482.php",
        "檔案下載": "https://curri.ttu.edu.tw/p/412-1033-1187.php",
        "校課程委員會會議紀錄": "https://curri.ttu.edu.tw/p/412-1033-1191.php",
        "暑修須知": "https://curri.ttu.edu.tw/p/412-1033-1192.php",
        "課程地圖": "https://curri.ttu.edu.tw/p/412-1033-1193.php",
        "就學服役彈性修業專區": "https://curri.ttu.edu.tw/p/412-1033-3056.php",
        "校外實習流程": "https://curri.ttu.edu.tw/p/412-1033-3410.php",
    },
    "招生組": {
        "人員執掌": "https://admission.ttu.edu.tw/p/412-1034-451.php",
        "高中專區_申請入學專區": "https://admission.ttu.edu.tw/p/412-1034-2095.php",
        "高中專區_學群介紹模擬面試": "https://admission.ttu.edu.tw/p/412-1034-2096.php",
        "獎學金_尚志獎學金": "https://tturule.ttu.edu.tw/rulelist/showcontent.php?rid=192",
        "獎學金_原住民籍助學金": "https://tturule.ttu.edu.tw/rulelist/showcontent.php?rid=826",
        "獎學金_五年一貫助學金": "https://tturule.ttu.edu.tw/rulelist/showcontent.php?rid=384",
        "獎學金_合作機構碩士在職專班": "https://tturule.ttu.edu.tw/rulelist/showcontent.php?rid=551",
        "獎學金_博士班研究生": "https://tturule.ttu.edu.tw/rulelist/showcontent.php?rid=552",
        "獎學金_外國學生": "https://tturule.ttu.edu.tw/rulelist/showcontent.php?rid=907",
        "獎學金_僑生港澳學生": "https://tturule.ttu.edu.tw/rulelist/showcontent.php?rid=908",
        "獎學金_大陸地區研究生": "https://tturule.ttu.edu.tw/rulelist/showcontent.php?rid=829",
    },
    "教學發展中心": {
        "中心簡介": "https://tldc.ttu.edu.tw/p/412-1003-505.php",
        "教學實踐研究_計畫介紹": "https://tldc.ttu.edu.tw/p/412-1003-2713.php",
        "教學實踐研究_歷年成果": "https://tldc.ttu.edu.tw/p/412-1003-2501.php",
        "教學實踐研究_相關法規": "https://tldc.ttu.edu.tw/p/406-1003-44227,r629.php",
        "教學創新與精進計畫": "https://tldc.ttu.edu.tw/p/412-1003-525.php",
        "教師教學卓越": "https://tldc.ttu.edu.tw/p/412-1003-534.php",
        "遠距教學與數位教學課程": "https://tldc.ttu.edu.tw/p/412-1003-3087.php",
        "自主學習課程": "https://tldc.ttu.edu.tw/p/412-1003-3419.php",
    },
}

# ==================== 工具函式 ====================
def clean_filename(name: str, max_len=80) -> str:
    name = re.sub(r'[\\/:*?"<>|\n\r\t]', "_", name)
    return re.sub(r"\s+", " ", name).strip()[:max_len]

def parse_date(text: str):
    if not text:
        return None
    text = str(text)
    m = re.search(r"(\d{4})[年\-/\.](\d{1,2})[月\-/\.](\d{1,2})", text)
    if m:
        y, mo, d = map(int, m.groups())
        if y < 1912:
            y += 1911
        try:
            return datetime(y, mo, d)
        except:
            pass
    m = re.search(r"(\d{2,3})年(\d{1,2})月(\d{1,2})日", text)
    if m:
        y, mo, d = map(int, m.groups())
        y += 1911
        try:
            return datetime(y, mo, d)
        except:
            pass
    return None

def parse_school_year(text: str):
    m = re.search(r"(11[2-6])\s*學年", text)
    if m:
        return int(m.group(1))
    return None

def is_within_three_years(date_obj=None, text=""):
    if date_obj:
        return DATE_START <= date_obj <= DATE_END
    sy = parse_school_year(text)
    if sy and sy in VALID_SCHOOL_YEARS:
        return True
    d = parse_date(text)
    if d:
        return DATE_START <= d <= DATE_END
    return False

def should_keep_always(title: str) -> bool:
    return any(kw in title for kw in KEEP_ALWAYS_KEYWORDS)

def get_soup(url: str):
    """修正亂碼"""
    try:
        resp = SESSION.get(url, timeout=25)
        resp.raise_for_status()

        if resp.encoding is None or resp.encoding.lower() in ["iso-8859-1", "windows-1252"]:
            resp.encoding = resp.apparent_encoding

        try:
            text = resp.content.decode("utf-8")
        except UnicodeDecodeError:
            try:
                text = resp.content.decode("big5")
            except UnicodeDecodeError:
                text = resp.content.decode("utf-8", errors="replace")

        return BeautifulSoup(text, "html.parser")
    except Exception as e:
        print(f"  [錯誤] 無法取得頁面: {e}")
        return None

def download_file(file_url: str, save_path: Path, referer=None) -> bool:
    headers = SESSION.headers.copy()
    if referer:
        headers["Referer"] = referer
    try:
        with SESSION.get(file_url, stream=True, timeout=50, headers=headers) as r:
            r.raise_for_status()
            with open(save_path, "wb") as f:
                for chunk in r.iter_content(8192):
                    f.write(chunk)
        return True
    except Exception as e:
        print(f"    下載失敗: {e}")
        return False

def extract_text_from_file(file_path: Path) -> str:
    suffix = file_path.suffix.lower()
    try:
        if suffix == ".pdf" and PdfReader:
            reader = PdfReader(str(file_path))
            return "\n".join([p.extract_text() or "" for p in reader.pages]).strip()
        elif suffix in (".docx", ".doc") and Document:
            doc = Document(str(file_path))
            return "\n".join([p.text for p in doc.paragraphs]).strip()
        elif suffix == ".odt" and HAS_ODF:
            doc = odf_load(str(file_path))
            return "\n".join([teletype.extractText(p) for p in doc.getElementsByType(odf_text.P)]).strip()
        elif suffix in (".xlsx", ".xls") and load_workbook:
            wb = load_workbook(str(file_path), data_only=True)
            texts = []
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                for row in ws.iter_rows(values_only=True):
                    line = " | ".join([str(c) if c is not None else "" for c in row])
                    if line.strip(" |"):
                        texts.append(line)
            return "\n".join(texts)
        elif suffix in (".pptx", ".ppt") and Presentation:
            prs = Presentation(str(file_path))
            texts = []
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text.strip():
                        texts.append(shape.text)
            return "\n".join(texts)
        elif suffix in (".txt", ".csv"):
            return file_path.read_text(encoding="utf-8", errors="ignore")
    except Exception as e:
        return f"[讀取失敗: {e}]"
    return f"[不支援格式: {suffix}]"

def extract_page_content(soup: BeautifulSoup) -> str:
    if not soup:
        return ""
    for tag in soup.select("header, footer, nav, script, style, .breadcrumb"):
        tag.decompose()
    candidates = [
        soup.select_one(".md_middle"),
        soup.select_one("#Dyn_2_2"),
        soup.select_one(".page_content"),
        soup.select_one(".content"),
        soup.select_one("article"),
        soup.select_one("main"),
    ]
    area = next((c for c in candidates if c and len(c.get_text(strip=True)) > 30), soup.body)
    return area.get_text("\n", strip=True) if area else ""

def get_real_filename(file_url: str, link_text: str = "") -> str:
    """正確處理 Action=downloadfile 的檔名"""
    # 1. 從 fname 參數取
    m = re.search(r"[?&]fname=([^&]+)", file_url, re.I)
    if m:
        fname = unquote(m.group(1))
        try:
            fname = unquote(fname)
        except:
            pass
        if fname and len(fname) > 3:
            return clean_filename(fname)

    # 2. 用連結文字
    if link_text and re.search(r"\.(pdf|docx?|xlsx?|pptx?|odt|zip|rar)$", link_text, re.I):
        return clean_filename(link_text)

    # 3. 預設
    return "attachment.pdf"

def find_attachments(soup, base_url):
    results = []
    if not soup:
        return results
    for a in soup.select("a[href]"):
        href = a.get("href", "")
        text = a.get_text(strip=True)
        is_download = (
            re.search(r"\.(pdf|docx?|xlsx?|pptx?|odt|zip|rar|7z|txt|csv)$", href, re.I)
            or "Action=downloadfile" in href
            or "downloadfile" in href.lower()
            or (text and re.search(r"\.(pdf|docx?|xlsx?|pptx?|odt)$", text, re.I))
        )
        if is_download:
            full_url = urljoin(base_url, href)
            results.append((full_url, text or "attachment"))
    return results

def find_date_from_page_and_files(soup, content, attachments_info):
    if soup:
        for sel in [".date", ".post-date", "time", "i.mdate", ".mdate"]:
            el = soup.select_one(sel)
            if el:
                d = parse_date(el.get_text())
                if d:
                    return d.strftime("%Y-%m-%d"), d

    d = parse_date(content)
    if d:
        return d.strftime("%Y-%m-%d"), d

    sy = parse_school_year(content)
    if sy:
        return f"{sy}學年度", None

    for att in attachments_info:
        d = parse_date(att.get("檔名", ""))
        if d:
            return d.strftime("%Y-%m-%d"), d
        sy = parse_school_year(att.get("檔名", ""))
        if sy:
            return f"{sy}學年度", None

    return "無", None

def find_sub_links(soup, base_url):
    if not soup:
        return []
    links = []
    for a in soup.select("a[href*='/p/406-'], a[href*='/p/405-'], a[href*='showcontent']"):
        href = a.get("href", "")
        title = a.get_text(strip=True)
        if title and len(title) > 4 and "更多" not in title:
            links.append((urljoin(base_url, href), title))
    return links[:12]

# ==================== 主爬取 ====================
def crawl_one_page(category, title, url, depth=0):
    print(f"\n{'  ' * depth}▶ [{category}] {title}")
    print(f"{'  ' * depth}  {url}")

    # 直接是 PDF
    if url.lower().endswith(".pdf"):
        attach_dir = ATTACH_ROOT / clean_filename(category)
        attach_dir.mkdir(parents=True, exist_ok=True)
        fname = clean_filename(title) + ".pdf"
        save_path = attach_dir / fname
        if download_file(url, save_path):
            text = extract_text_from_file(save_path)
            return {
                "標題": title,
                "日期": "無",
                "網址": url,
                "內容": "",
                "附件": [{"檔名": fname, "文字內容": text}],
                "子頁面": []
            }
        return None

    soup = get_soup(url)
    if not soup:
        return {
            "標題": title,
            "日期": "無",
            "網址": url,
            "內容": "",
            "附件": [],
            "子頁面": []
        }

    content = extract_page_content(soup)
    attachments = []
    attach_dir = ATTACH_ROOT / clean_filename(category)
    attach_dir.mkdir(parents=True, exist_ok=True)

    for file_url, link_text in find_attachments(soup, url):
        # 正確取得檔名（解決 .php 問題）
        real_name = get_real_filename(file_url, link_text)
        if not re.search(r"\.(pdf|docx?|xlsx?|pptx?|odt|zip|rar|txt|csv)$", real_name, re.I):
            real_name += ".pdf"

        fname = clean_filename(f"{title}_{real_name}")
        save_path = attach_dir / fname

        if not save_path.exists():
            ok = download_file(file_url, save_path, referer=url)
        else:
            ok = True

        if ok:
            file_text = extract_text_from_file(save_path)
            attachments.append({
                "檔名": fname,
                "原始連結": file_url,
                "文字內容": file_text
            })
            print(f"{'  ' * depth}    ✅ 附件: {fname}")

    date_str, date_obj = find_date_from_page_and_files(soup, content, attachments)
    print(f"{'  ' * depth}  日期：{date_str}")

    # 日期過濾（固定頁例外）
    if not should_keep_always(title):
        if date_str != "無":
            if not is_within_three_years(date_obj=date_obj, text=content + title + date_str):
                print(f"{'  ' * depth}  ⏭️ 非近三年，跳過")
                return None

    result = {
        "標題": title,
        "日期": date_str,
        "網址": url,
        "內容": content,
        "附件": attachments,
        "子頁面": []
    }

    # 往下一層
    if depth < 1:
        for sub_url, sub_title in find_sub_links(soup, url):
            time.sleep(random.uniform(0.5, 1.0))
            sub_result = crawl_one_page(category, sub_title, sub_url, depth + 1)
            if sub_result:
                result["子頁面"].append(sub_result)

    return result

def main():
    SAVE_ROOT.mkdir(exist_ok=True)
    ATTACH_ROOT.mkdir(exist_ok=True)

    all_data = {
        "產生時間": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        "說明": "近三年（含學年度），多層頁面 + 附件日期判斷 + 修正亂碼 + 修正downloadfile",
        "資料": {}
    }

    for category, pages in TARGETS.items():
        print(f"\n{'='*60}")
        print(f"開始：{category}")
        print(f"{'='*60}")

        cat_results = []
        for title, url in pages.items():
            try:
                result = crawl_one_page(category, title, url)
                if result:
                    cat_results.append(result)
            except Exception as e:
                print(f"  [例外] {title}: {e}")
            time.sleep(random.uniform(0.8, 1.5))

        all_data["資料"][category] = cat_results

        out_path = SAVE_ROOT / f"{clean_filename(category)}.json"
        with open(out_path, "w", encoding="utf-8") as f:
            json.dump(cat_results, f, ensure_ascii=False, indent=2)
        print(f"\n✅ {category} 完成 → {out_path.name}")

    with open(SAVE_ROOT / "all_pages.json", "w", encoding="utf-8") as f:
        json.dump(all_data, f, ensure_ascii=False, indent=2)

    print(f"\n🎉 全部完成！")
    print(f"結果資料夾：{SAVE_ROOT.absolute()}")

if __name__ == "__main__":
    print("請先安裝：")
    print("pip install requests beautifulsoup4 pypdf python-docx odfpy openpyxl python-pptx")
    print()
    main()