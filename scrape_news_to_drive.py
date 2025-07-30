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
from pydrive2.auth import GoogleAuth
from pydrive2.drive import GoogleDrive

# --- 設定 ---
SPREADSHEET_URL = "https://docs.google.com/spreadsheets/d/1nphpu1q2cZuxJe-vYuOliw1azxqKKlzt6FFGNEJ76sw/edit?gid=0#gid=0"
DRIVE_FOLDER_ID = "1MjNzGR57vsLtjbBJAZl06BKqZALYjGUO"
SERVICE_ACCOUNT_FILE = 'service_account.json'
TODAY_STR = datetime.now().strftime("%y%m%d")
OUTPUT_FILE = f"{TODAY_STR}.xlsx"

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
            content = "\n".join([p.text for p in paragraphs if p.text.strip()]).strip()
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
            page_comments.append((comment_text, formatted_time, user_name))
        joined_current_page = "\n".join([c[0] for c in page_comments])
        if not page_comments or joined_current_page == last_comments_joined:
            break
        last_comments_joined = joined_current_page
        comments_data.extend(page_comments)
        page += 1
    return comments_data if comments_data else [("コメントなし", "", "")]

def get_urls_from_google_sheet(spreadsheet_url):
    try:
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(SERVICE_ACCOUNT_FILE, scope)
        client = gspread.authorize(creds)
        sheet = client.open_by_url(spreadsheet_url).get_worksheet(0)
        rows = sheet.get_all_values()
        news_data = []
        for row_idx, row in enumerate(rows[1:], 2):
            if len(row) > 1 and row[1].startswith("http"):
                try:
                    post_date_str = row[2] if len(row) > 2 else ""
                    post_datetime = None
                    if post_date_str:
                        date_formats = ["%Y/%m/%d %H:%M:%S", "%Y/%m/%d %H:%M", "%m/%d %H:%M"]
                        for fmt in date_formats:
                            try:
                                dt = datetime.strptime(post_date_str, fmt)
                                if fmt == "%m/%d %H:%M":
                                    dt = dt.replace(year=datetime.now().year)
                                post_datetime = dt
                                break
                            except ValueError:
                                continue
                        if post_datetime is None:
                            raise ValueError(f"Unsupported date format: {post_date_str}")
                    news_data.append({
                        'url': row[1],
                        'post_datetime': post_datetime,
                        'original_A_col': row[0] if len(row) > 0 else '',
                        'original_B_col': row[1],
                        'original_C_col': row[2] if len(row) > 2 else '',
                        'original_row_index': row_idx
                    })
                except Exception as e:
                    print(f"行 {row_idx} の処理エラー: {e}")
        return rows, news_data
    except Exception as e:
        print(f"Google Sheetsの取得失敗: {e}")
        return [], []

def upload_file_to_drive(file_path, drive_folder_id=None):
    try:
        gauth = GoogleAuth()
        gauth.LoadServiceConfigFile(SERVICE_ACCOUNT_FILE)
        gauth.Authorize()

        drive = GoogleDrive(gauth)
        file_name = os.path.basename(file_path)

        metadata = {'title': file_name}
        if drive_folder_id:
            metadata['parents'] = [{"kind": "drive#fileLink", "id": drive_folder_id}]

        f = drive.CreateFile(metadata)
        f.SetContentFile(file_path)
        f.Upload()
        print(f"✅ Google Driveにアップロード完了: {f['id']}")
        return f['id']

    except Exception as e:
        print(f"Driveアップロード失敗: {e}")
        return None

def main():
    rows, news_list = get_urls_from_google_sheet(SPREADSHEET_URL)
    if not news_list:
        print("❌ Google Sheetsから処理対象のURLを1件も読み込めませんでした。プログラムを終了します。")
        return
    now = datetime.now()
    start_time = (now - timedelta(days=1)).replace(hour=15, minute=0, second=0, microsecond=0)
    end_time = now.replace(hour=14, minute=59, second=59, microsecond=999999)
    filtered_news = [n for n in news_list if n['post_datetime'] and start_time <= n['post_datetime'] <= end_time]
    if not filtered_news:
        print("❌ 条件に一致するニュースがありません。")
        return
    print(f"✅ フィルタ一致件数: {len(filtered_news)}")
    options = Options()
    options.add_argument("--lang=ja-JP")
    options.add_argument("--headless")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    driver = webdriver.Chrome(options=options)
    wb = Workbook()
    ws_input = wb.active
    ws_input.title = "input"
    for r, row_data in enumerate(rows, 1):
        for c, val in enumerate(row_data, 1):
            ws_input.cell(row=r, column=c, value=val)
    comment_col = max(len(rows[0]) if rows else 0, 5) + 1
    ws_input.cell(row=1, column=comment_col, value="コメント件数")
    for idx, news in enumerate(filtered_news, 1):
        url = news['url']
        row_idx = news['original_row_index']
        print(f"▶ ({idx}/{len(filtered_news)}) 処理中: {url}")
        ws = wb.create_sheet(title=f"News_{idx}")
        try:
            title, base_url, pages = get_news_pages(url, driver)
            ws.cell(row=1, column=1, value="タイトル")
            ws.cell(row=1, column=2, value=title)
            ws.cell(row=2, column=1, value="URL")
            ws.cell(row=2, column=2, value=base_url)
            for i, page_text in enumerate(pages[:15], 1):
                ws.cell(row=i + 2, column=1, value=page_text)
        except Exception as e:
            ws.cell(row=1, column=1, value="記事取得失敗")
            ws.cell(row=2, column=1, value=str(e))
        try:
            comments = get_comments_pages(url, driver)
            ws.cell(row=19, column=1, value="コメント本文")
            ws.cell(row=19, column=2, value="投稿日時")
            ws.cell(row=19, column=3, value="ユーザー名")
            for i, (text, dt, user) in enumerate(comments, start=20):
                ws.cell(row=i, column=1, value=text)
                ws.cell(row=i, column=2, value=dt)
                ws.cell(row=i, column=3, value=user)
            comment_count = len(comments) if comments[0][0] != "コメントなし" else 0
            ws_input.cell(row=row_idx, column=comment_col, value=comment_count)
        except Exception as e:
            ws.cell(row=20, column=1, value="コメント取得失敗")
            ws.cell(row=20, column=2, value=str(e))
            ws_input.cell(row=row_idx, column=comment_col, value="取得失敗")
    driver.quit()
    wb.save(OUTPUT_FILE)
    print(f"✅ ローカル保存完了: {OUTPUT_FILE}")
    upload_file_to_drive(OUTPUT_FILE, DRIVE_FOLDER_ID)
    try:
        os.remove(OUTPUT_FILE)
    except OSError:
        pass

if __name__ == "__main__":
    main()
