# -*- coding: utf-8 -*-
import glob
import os
import re
import shutil
import time
from datetime import datetime, timedelta
from pathlib import Path

import pandas as pd
import pdfplumber
from selenium import webdriver
from selenium.common.exceptions import (NoSuchElementException,
                                          TimeoutException)
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.support.ui import WebDriverWait
from webdriver_manager.chrome import ChromeDriverManager

# --------------------------- CẤU HÌNH ------------------------------------

from excel_fund_updater import update_fund_excel
URL = "https://www.hanwhalife.com.vn/vi/news#GiaDonViQuy"
EXCEL_FILE = "gia_don_vi_quy_TONGHOP.xlsx"
SHEET_NAME = "Hanwha"
DOWNLOAD_DIR = Path(__file__).with_name("_hanwha_downloads")
DEFAULT_START_DATE = "01/01/2021"  # ngày bắt đầu cho lần chạy đầu tiên

DATE_COLUMN = "Ngày"
FUND_COLUMNS = [
    "Quỹ Tăng trưởng chiến lược",
    "Quỹ Tăng trưởng",
    "Quỹ Cổ phiếu hàng đầu",
    "Quỹ Bền vững",
]

HEADLESS = True  # đặt False nếu bạn muốn xem trình duyệt chạy trực tiếp (để debug)

# Regex bắt 1 dòng dữ liệu trong text thô của PDF, ví dụ:
#   "1 18/06/2026 10.812,11 12.667,33 12.931,36 13.776,37"
# Nhóm: ngày (dd/mm/yyyy), rồi 4 giá trị số dạng "12.345,67"
PDF_ROW_RE = re.compile(
    r"^\d+\s+(\d{2}/\d{2}/\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$"
)

# Một số quỹ chưa ra mắt ở những ngày cũ -> cột đó bị TRỐNG trong PDF, dòng chỉ
# còn 3 giá trị số thay vì 4. Regex này bắt trường hợp thiếu giá trị ĐẦU TIÊN
# (cột "Quỹ Tăng trưởng chiến lược"), ví dụ:
#   "104 04/07/2024 11.077,38 8.587,28 12.118,75"
PDF_ROW_RE_MISSING_FIRST = re.compile(
    r"^\d+\s+(\d{2}/\d{2}/\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s*$"
)

# --------------------------------------------------------------------------


def build_driver() -> webdriver.Chrome:
    DOWNLOAD_DIR.mkdir(exist_ok=True)
    options = Options()
    if HEADLESS:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1366,900")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    prefs = {
        "download.default_directory": str(DOWNLOAD_DIR.resolve()),
        "download.prompt_for_download": False,
        "safebrowsing.enabled": True,
    }
    options.add_experimental_option("prefs", prefs)

    service = service = Service(r"C:\chromedriver\chromedriver.exe")
    driver = webdriver.Chrome(service=service, options=options)
    driver.implicitly_wait(5)
    return driver


def set_date_input(driver, label_text: str, value: str):
    """Tìm ô input ngày nằm gần nhãn (label) chứa label_text, ví dụ
    'Từ ngày' hoặc 'Đến ngày', rồi gán giá trị."""
    xpath = (
        f"//*[contains(text(), '{label_text}')]"
        "/following::input[@type='text' or not(@type)][1]"
    )
    el = driver.find_element(By.XPATH, xpath)
    driver.execute_script("arguments[0].removeAttribute('readonly')", el)
    driver.execute_script("arguments[0].value = arguments[1];", el, value)
    # bắn các event để JS của trang nhận biết giá trị đã đổi
    driver.execute_script(
        "arguments[0].dispatchEvent(new Event('change', {bubbles: true}));"
        "arguments[0].dispatchEvent(new Event('input', {bubbles: true}));",
        el,
    )


def click_button_by_text(driver, text: str):
    xpath = f"//*[self::button or self::a][contains(., '{text}')]"
    el = driver.find_element(By.XPATH, xpath)
    driver.execute_script("arguments[0].click();", el)


def get_results_table(driver):
    """Bảng kết quả nằm dưới phần 'TRA CỨU GIÁ ĐƠN VỊ QUỸ', có cột đầu là NGÀY."""
    xpath = "//table[.//th[contains(., 'NGÀY') or contains(., 'Ngày')]]"
    tables = driver.find_elements(By.XPATH, xpath)
    # Lấy bảng kết quả tra cứu (thường là bảng cuối cùng khớp điều kiện)
    return tables[-1] if tables else None


def parse_table_rows(table) -> list:
    rows = []
    for tr in table.find_elements(By.TAG_NAME, "tr")[1:]:  # bỏ header
        cells = [td.text.strip() for td in tr.find_elements(By.TAG_NAME, "td")]
        if len(cells) >= 5 and "/" in cells[0]:
            rows.append(cells[:5])
    return rows


def extract_pdf_table(filepath: Path) -> pd.DataFrame:
    """Đọc PDF bằng cách parse TEXT THÔ (đáng tin cậy hơn extract_tables(),
    vì pdfplumber nhận sai đường viền cột cuối của bảng PDF này - cột cuối
    "Quỹ Bền vững" bị extract_tables() trả về None). Mỗi dòng dữ liệu hợp lệ
    trong text thô có dạng:
        "<STT> <dd/mm/yyyy> <giá1> <giá2> <giá3> <giá4>"
    nên chỉ cần regex là bắt được, không cần dò cấu trúc bảng.
    """
    rows = []
    with pdfplumber.open(filepath) as pdf:
        print(f"  [debug] PDF có {len(pdf.pages)} trang")
        for i, page in enumerate(pdf.pages):
            text = page.extract_text() or ""
            page_rows = 0
            for line in text.split("\n"):
                line = line.strip()
                m = PDF_ROW_RE.match(line)
                if m:
                    date, v1, v2, v3, v4 = m.groups()
                    rows.append([date, v1, v2, v3, v4])
                    page_rows += 1
                    continue
                # Quỹ "Tăng trưởng chiến lược" chưa ra mắt ở ngày này -> cột
                # đầu bị trống trong PDF, dòng chỉ còn 3 giá trị số.
                m2 = PDF_ROW_RE_MISSING_FIRST.match(line)
                if m2:
                    date, v2, v3, v4 = m2.groups()
                    rows.append([date, "", v2, v3, v4])
                    page_rows += 1
            print(f"  [debug] Trang {i + 1}: khớp {page_rows} dòng dữ liệu")

    if not rows:
        print("  [!] Không đọc được dòng dữ liệu nào từ PDF.")
        return pd.DataFrame(columns=[DATE_COLUMN, *FUND_COLUMNS])

    return pd.DataFrame(rows, columns=[DATE_COLUMN, *FUND_COLUMNS])


def try_download_file(driver) -> pd.DataFrame:
    """Thử bấm nút 'Tải về'. Nếu có file mới xuất hiện trong DOWNLOAD_DIR,
    đọc file đó (PDF, Excel hoặc CSV) và trả về DataFrame. Nếu không có gì
    xảy ra trong vài giây, trả về None để dùng phương án lặp qua từng trang."""
    before = set(os.listdir(DOWNLOAD_DIR))
    try:
        click_button_by_text(driver, "Tải về")
    except NoSuchElementException:
        return None

    for _ in range(20):  # chờ tối đa ~20 giây cho file tải xong
        time.sleep(1)
        after = set(os.listdir(DOWNLOAD_DIR)) - before
        finished = [f for f in after if not f.endswith((".crdownload", ".tmp"))]
        if finished:
            filepath = DOWNLOAD_DIR / finished[0]
            suffix = filepath.suffix.lower()
            try:
                if suffix == ".pdf":
                    print(f"  Đã tải về file PDF: {filepath.name} -> đang đọc bảng...")
                    return extract_pdf_table(filepath)
                if suffix in (".xlsx", ".xls"):
                    return pd.read_excel(filepath)
                if suffix == ".csv":
                    return pd.read_csv(filepath, encoding="utf-8-sig")
                print(f"  [!] Định dạng file tải về không nhận diện được: {suffix}")
                return None
            except Exception as exc:
                print(f"  [!] Đọc file tải về lỗi: {exc}")
                return None
    return None


def scrape_by_pagination(driver) -> pd.DataFrame:
    """Lặp qua từng trang kết quả, đọc dữ liệu bảng cho đến khi hết trang."""
    all_rows = []
    page_num = 1
    while True:
        table = get_results_table(driver)
        if table is None:
            print("  [!] Không tìm thấy bảng kết quả.")
            break
        rows = parse_table_rows(table)
        print(f"  Trang {page_num}: {len(rows)} dòng")
        if not rows:
            break
        all_rows.extend(rows)

        # tìm nút "Sau" / ">" để qua trang kế tiếp
        try:
            next_btn = driver.find_element(
                By.XPATH,
                "//*[self::a or self::button]"
                "[contains(., 'Sau') or normalize-space(text())='>']",
            )
            classes = next_btn.get_attribute("class") or ""
            if "disabled" in classes:
                break
            driver.execute_script("arguments[0].click();", next_btn)
            page_num += 1
            time.sleep(1.5)  # chờ bảng load lại
        except NoSuchElementException:
            break

    if not all_rows:
        return pd.DataFrame(columns=[DATE_COLUMN, *FUND_COLUMNS])
    return pd.DataFrame(all_rows, columns=[DATE_COLUMN, *FUND_COLUMNS])


def clean_number(value):
    if pd.isna(value):
        return None
    text = str(value)
    text = "".join(ch for ch in text if ch.isdigit() or ch in ",.")
    text = text.replace(".", "").replace(",", ".")
    try:
        return float(text)
    except ValueError:
        return None


def normalize_dataframe(df: pd.DataFrame) -> pd.DataFrame:
    """Đưa DataFrame (lấy từ file tải về hoặc từ bảng web) về đúng format:
    cột Ngày kiểu datetime, các cột giá kiểu số."""
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    # đoán cột ngày là cột đầu tiên nếu tên không khớp chính xác
    date_col = df.columns[0]
    df.rename(columns={date_col: DATE_COLUMN}, inplace=True)
    df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN], dayfirst=True, errors="coerce")
    for col in df.columns[1:]:
        df[col] = df[col].apply(clean_number)
    df.dropna(subset=[DATE_COLUMN], inplace=True)
    return df


def load_existing(path: Path) -> pd.DataFrame:
    if path.exists():
        df = pd.read_excel(path)
        df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN], dayfirst=True)
        return df
    return pd.DataFrame(columns=[DATE_COLUMN, *FUND_COLUMNS])


def main():
    existing = load_existing(Path(EXCEL_FILE))
    if existing.empty:
        start_date = DEFAULT_START_DATE
    else:
        last_date = existing[DATE_COLUMN].max()
        start_date = (last_date - timedelta(days=3)).strftime("%d/%m/%Y")
        # lùi lại vài ngày để chắc chắn không bỏ lỡ dòng nào, sẽ tự loại trùng sau
    end_date = datetime.now().strftime("%d/%m/%Y")

    print(f"Khoảng ngày sẽ tra cứu: {start_date} -> {end_date}")

    if DOWNLOAD_DIR.exists():
        shutil.rmtree(DOWNLOAD_DIR)
    DOWNLOAD_DIR.mkdir(exist_ok=True)

    driver = build_driver()
    new_df = pd.DataFrame(columns=[DATE_COLUMN, *FUND_COLUMNS])
    try:
        driver.get(URL)
        WebDriverWait(driver, 20).until(
            EC.presence_of_element_located((By.XPATH, "//*[contains(text(), 'TRA CỨU GIÁ ĐƠN VỊ QUỸ')]"))
        )
        # cuộn tới khu vực tra cứu để các phần tử "khả kiến" với Selenium
        driver.execute_script(
            "document.querySelector(\"[id*='GiaDonViQuy'], a[href*='GiaDonViQuy']\")"
            "?.scrollIntoView();"
        )
        time.sleep(1)

        set_date_input(driver, "Từ ngày", start_date)
        set_date_input(driver, "Đến ngày", end_date)
        click_button_by_text(driver, "Tìm kiếm")

        WebDriverWait(driver, 20).until(
            lambda d: get_results_table(d) is not None
        )
        time.sleep(1.5)

        downloaded_df = try_download_file(driver)
        if downloaded_df is not None and not downloaded_df.empty:
            print("Đã lấy được dữ liệu qua nút 'Tải về'.")
            new_df = downloaded_df
        else:
            print("Không lấy được file tải về -> chuyển sang đọc từng trang.")
            new_df = scrape_by_pagination(driver)
    except TimeoutException:
        print("[!] Trang tải quá lâu hoặc không tìm thấy khu vực tra cứu.")
    finally:
        driver.quit()

    if new_df.empty:
        print("Không lấy được dữ liệu mới nào.")
        return

    new_df = normalize_dataframe(new_df)

    combined = pd.concat([existing, new_df], ignore_index=True)
    # ÉP LẠI kiểu datetime cho cột Ngày: nếu "existing" là bảng trống (object dtype)
    # thì pd.concat có thể làm cột Ngày của bảng kết quả bị "nhiễm" về dtype
    # object, khiến .dt accessor lỗi phía dưới -> luôn convert lại cho chắc.
    combined[DATE_COLUMN] = pd.to_datetime(combined[DATE_COLUMN], errors="coerce")
    combined.dropna(subset=[DATE_COLUMN], inplace=True)
    combined.drop_duplicates(subset=[DATE_COLUMN], keep="last", inplace=True)
    combined.sort_values(DATE_COLUMN, ascending=False, inplace=True)
    combined.reset_index(drop=True, inplace=True)

    if combined.empty:
        print("Không có dữ liệu hợp lệ nào để lưu.")
        return

    # ✅ convert dataframe -> dict
    new_data = {}

    for _, row in combined.iterrows():
        date_str = row[DATE_COLUMN].strftime("%d/%m/%Y")
        fund_prices = {}

        for col in FUND_COLUMNS:
            value = row.get(col)
            if pd.notna(value):
                fund_prices[col] = float(value)

        if fund_prices:
            new_data[date_str] = fund_prices

    # ✅ update Excel ngay trong main
    added = update_fund_excel(
        file_path=EXCEL_FILE,
        sheet_name=SHEET_NAME,
        new_data=new_data,
    )

    if added:
        print(f"[+] Đã thêm {added} dòng mới vào {EXCEL_FILE} (sheet {SHEET_NAME})")
    else:
        print("[=] Không có dòng mới (dữ liệu đã tồn tại)")


if __name__ == "__main__":
    main()