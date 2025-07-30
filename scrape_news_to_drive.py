from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from bs4 import BeautifulSoup
from datetime import datetime, timedelta
import time
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# --- 設定 ---
SPREADSHEET_ID = "1nphpu1q2cZuxJe-vYuOliw1azxqKKlzt6FFGNEJ76sw"  # スプレッドシートID
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
    except:
        pass
    try:
        return datetime.strptime(text, "%Y/%m/%d %H:%M:%S")
    except:
        return now

def format_datetime(dt):
    return dt.strftime("%y/%m/%d %H:%M")

def get_news_pages(base_url, driver):
    page_contents = []
    page = 1
    last_content = ""
    while True:
        url = base_url if page == 1 else f"{base_url}?page={page}"
        driver.get(url)
        time.sleep(2)
        try:
            article = driver.find_element(By.TAG_NAME, "article")
            paragraphs = article.find_elements(By.TAG_NAME, "p")
            content = "\n".join([p.text for p in paragraphs if p.text.strip()])
            if not content or content == last_content:
                break
            page_contents.append(content)
            last_content = content
            page += 1
        except:
            break
    driver.get(base_url)
    time.sleep(1)
    try:
        title = driver.title.replace(" - Yahoo!ニュース", "")
    except:
        title = "タイトル取得失敗"
    return title, base_url, page_contents

def get_comments_pages(base_url, driver):
    comments_data = []
    page = 1
    last_comments_joined = ""
    article_id = base_url.rstrip("/").split("/")[-1]
    base_comment_url = f"https://news.yahoo.co.jp/articles/{article_id}/comments"
    while True:
        comment_url = base_comment_url if page == 1 else f"{base_comment_url}?page={page}"
        driver.get(comment_url)
        time.sleep(2)
        soup = BeautifulSoup(driver.page_source, 'html.parser')
        comment_elements = soup.find_all('article', class_='sc-169yn8p-3')
        page_comments = []
        for comment_article in comment_elements:
            comment_p = comment_article.find('p', class_='sc-169yn8p-10')
            comment_text = comment_p.text.strip() if comment_p else ''
            user_a = comment_article.find('a', class_='sc-169yn8p-7')
            user_name = user_a.text.strip() if user_a else ''
            time_a = comment_article.find('a', class_='sc-169yn8p-9')
            raw_time = time_a.text.strip() if time_a else ''
            dt = parse_relative_time(raw_time)
            formatted_time = format_datetime(dt)
            page_comments.append([comment_text, formatted_time, user_name])
        joined_current_page = "\n".join([c[0] for c in page_comments])
        if not page_comments or joined_current_page == last_comments_joined:
            break
        last_comments_joined = joined_current_page
        comments_data.extend(page_comments)
        page += 1
    return comments_data if comments_data else [["コメントなし", "", ""]]

def main():
    scope = ['https://spreadsheets.google.com/feeds','https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
    client = gspread.authorize(creds)

    # 入力シート読み込み
    sh = client.open_by_key(SPREADSHEET_ID)
    input_ws = sh.worksheet("input")
    rows = input_ws.get_all_values()
    urls = [row[1] for row in rows[1:] if len(row) > 1 and row[1].startswith("http")]
    if not urls:
        print("❌ URLが見つかりません")
        return

    # 出力シート準備
    try:
        output_ws = sh.worksheet(TODAY_STR)
        sh.del_worksheet(output_ws)
    except:
        pass
    output_ws = sh.add_worksheet(title=TODAY_STR, rows="1000", cols="20")
    output_ws.update("A1", "No")
    output_ws.update("B1", "URL")
    output_ws.update("C1", "投稿日時")
    output_ws.update("D1", "タイトル")
    output_ws.update("E1", "本文")
    output_ws.update("F1", "コメント数")

    # ブラウザ起動
    options = Options()
    options.add_argument("--lang=ja-JP")
    options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)

    for idx, url in enumerate(urls, 1):
        print(f"▶ ({idx}/{len(urls)}) 処理中: {url}")
        try:
            title, base_url, pages = get_news_pages(url, driver)
        except Exception as e:
            title, pages = "記事取得失敗", [str(e)]

        try:
            comments = get_comments_pages(url, driver)
            comment_count = len(comments) if comments[0][0] != "コメントなし" else 0
        except Exception as e:
            comments = [["コメント取得失敗", str(e), ""]]
            comment_count = "取得失敗"

        body_text = "\n\n".join(pages[:15])
        post_date = rows[idx][2] if len(rows[idx]) > 2 else ""

        # 本文・コメント数記入
        output_ws.update(f"A{idx+1}:F{idx+1}", [[idx, url, post_date, title, body_text, comment_count]])

    driver.quit()
    print("✅ 完了しました。")

if __name__ == "__main__":
    main()
