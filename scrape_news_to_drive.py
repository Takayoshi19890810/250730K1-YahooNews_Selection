from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
from openpyxl import Workbook
import time
import os
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- 設定 ---
SPREADSHEET_ID = "1nphpu1q2cZuxJe-vYuOliw1azxqKKlzt6FFGNEJ76sw"
INPUT_SHEET_NAME = "input"
SERVICE_ACCOUNT_FILE = 'service_account.json'
TODAY_STR = datetime.now().strftime("%y%m%d")


def parse_relative_time(text):
    now = datetime.now()
    try:
        if "分前" in text:
            return now - timedelta(minutes=int(text.replace("分前", "").strip()))
        elif "時間前" in text:
            return now - timedelta(hours=int(text.replace("時間前", "").strip()))
        elif "日前" in text:
            return now - timedelta(days=int(text.replace("日前", "").strip()))
        elif "秒前" in text:
            return now - timedelta(seconds=int(text.replace("秒前", "").strip()))
    except ValueError:
        pass
    try:
        for fmt in ("%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M"):
            return datetime.strptime(text, fmt)
    except ValueError:
        pass
    return now


def format_datetime(dt):
    return dt.strftime("%y/%m/%d %H:%M")


def get_news_content(driver, url):
    driver.get(url)
    time.sleep(2)
    try:
        title = driver.title.replace(" - Yahoo!ニュース", "")
        article = driver.find_element(By.TAG_NAME, "article")
        paragraphs = article.find_elements(By.TAG_NAME, "p")
        content = [p.text for p in paragraphs if p.text.strip()]
    except:
        title = "タイトル取得失敗"
        content = ["本文取得失敗"]
    return title, content


def get_comments(driver, url):
    article_id = url.rstrip("/").split("/")[-1]
    comments_url = f"https://news.yahoo.co.jp/articles/{article_id}/comments"
    comments = []
    page = 1
    last_text = ""
    while True:
        comment_page_url = f"{comments_url}?page={page}" if page > 1 else comments_url
        driver.get(comment_page_url)
        time.sleep(2)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        comment_elements = soup.find_all('article', class_='sc-169yn8p-3')
        current_page = []
        for elem in comment_elements:
            comment_p = elem.find('p', class_='sc-169yn8p-10')
            comment_text = comment_p.text.strip() if comment_p else ''
            user_a = elem.find('a', class_='sc-169yn8p-7')
            user_name = user_a.text.strip() if user_a else ''
            time_a = elem.find('a', class_='sc-169yn8p-9')
            raw_time = time_a.text.strip() if time_a else ''
            dt = parse_relative_time(raw_time)
            current_page.append((comment_text, format_datetime(dt), user_name))
        joined = "\n".join([c[0] for c in current_page])
        if not current_page or joined == last_text:
            break
        comments.extend(current_page)
        last_text = joined
        page += 1
    return comments if comments else [("コメントなし", "", "")]


def main():
    # 認証とシート取得
    scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
    client = gspread.authorize(creds)
    sheet = client.open_by_key(SPREADSHEET_ID)
    ws_input = sheet.worksheet(INPUT_SHEET_NAME)
    rows = ws_input.get_all_values()

    # 出力ブック作成
    wb = Workbook()
    ws_in = wb.active
    ws_in.title = "input"
    for r_idx, row in enumerate(rows, 1):
        for c_idx, val in enumerate(row, 1):
            ws_in.cell(row=r_idx, column=c_idx, value=val)

    options = Options()
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=options)

    for idx, row in enumerate(rows[1:], 2):
        if len(row) < 2 or not row[1].startswith("http"):
            continue
        url = row[1]
        ws = wb.create_sheet(title=str(idx - 1))
        try:
            title, content = get_news_content(driver, url)
            ws.cell(row=1, column=1, value="タイトル")
            ws.cell(row=1, column=2, value=title)
            ws.cell(row=2, column=1, value="URL")
            ws.cell(row=2, column=2, value=url)
            for i, text in enumerate(content[:15], start=3):
                ws.cell(row=i, column=1, value=text)
        except Exception as e:
            ws.cell(row=1, column=1, value="本文取得エラー")
            ws.cell(row=2, column=1, value=str(e))

        try:
            comments = get_comments(driver, url)
            ws.cell(row=19, column=1, value="コメント本文")
            ws.cell(row=19, column=2, value="投稿日時")
            ws.cell(row=19, column=3, value="投稿者")
            for i, (text, date, user) in enumerate(comments, start=20):
                ws.cell(row=i, column=1, value=text)
                ws.cell(row=i, column=2, value=date)
                ws.cell(row=i, column=3, value=user)
            ws_in.cell(row=idx, column=6, value=len(comments) if comments[0][0] != "コメントなし" else 0)
        except Exception as e:
            ws.cell(row=20, column=1, value="コメント取得失敗")
            ws.cell(row=20, column=2, value=str(e))
            ws_in.cell(row=idx, column=6, value="取得失敗")

    driver.quit()
    output_path = f"{TODAY_STR}.xlsx"
    wb.save(output_path)
    print(f"✅ ローカル保存完了: {output_path}")

if __name__ == "__main__":
    main()
