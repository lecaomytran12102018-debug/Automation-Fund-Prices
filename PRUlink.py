# -*- coding: utf-8 -*-
import re
import sys
import time
import unicodedata

from bs4 import BeautifulSoup
import undetected_chromedriver as uc
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait

from excel_fund_updater import update_fund_excel


URL = "https://www.prudential.com.vn/vi/cham-soc-khach-hang/thong-tin-cac-quy-dau-tu/quy-lien-ket-don-vi-prulink/"
EXCEL_FILE = "gia_don_vi_quy_TONGHOP.xlsx"  # file Excel gộp chung, đổi path nếu cần
SHEET_NAME = "PRUlink"  # sheet riêng cho PRUlink trong file gộp

DATE_HEADER_KEYWORD = "Ngày"

DATE_CELL_RE = re.compile(r"(\d{2}/\d{2}/\d{4})")
PRICE_CELL_RE = re.compile(r"^[\d]{1,3}(?:\.\d{3})*(?:,\d+)?$")


def _nfc(text):
    return unicodedata.normalize("NFC", text or "").strip()


def _vn_price_to_float(price_str):
    """'52.696' -> 52696.0"""
    return float(price_str.replace(".", "").replace(",", "."))


def _build_driver(headless=False):
    """
    Khởi tạo Chrome driver sử dụng undetected_chromedriver để vượt qua WAF.
    """
    options = uc.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    
    options.add_argument("--window-size=1366,900")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")

    
    driver = uc.Chrome(options=options, version_main=149)
    return driver


def _scroll_gradually(driver, step=400, pause=0.4, max_scrolls=40):
    """
    Cuộn dần từng đoạn nhỏ xuống cuối trang để trigger các widget lazy-load
    """
    for _ in range(max_scrolls):
        driver.execute_script(f"window.scrollBy(0, {step});")
        time.sleep(pause)
        height = driver.execute_script("return document.body.scrollHeight;")
        scroll_y = driver.execute_script("return window.scrollY;")
        if scroll_y + 900 >= height:  # đã gần chạm đáy trang
            break


def fetch_rendered_html(url=URL, headless=True, wait_seconds=25):
    """
    Mở trang bằng undetected_chromedriver, gom toàn bộ HTML của trang chính
    VÀ tất cả các iframe lại thành một cục để parser tự do tìm kiếm.
    """
    driver = _build_driver(headless=headless)
    try:
        driver.get(url)
        time.sleep(3)

        _scroll_gradually(driver)
        time.sleep(3)

        print("[*] Đang lấy HTML trang chính...")
        # Khởi tạo biến chứa toàn bộ HTML
        combined_html = driver.page_source
        
        iframes = driver.find_elements(By.TAG_NAME, "iframe")
        print(f"[*] Phát hiện {len(iframes)} iframe. Đang chui vào lấy thêm dữ liệu...")
        
        for index, iframe in enumerate(iframes):
            try:
                driver.switch_to.frame(iframe)
                # Nối thêm HTML của iframe vào biến combined_html
                combined_html += f"\n\n"
                combined_html += driver.page_source
                combined_html += f"\n\n"
                
                driver.switch_to.default_content()
            except Exception:
                # Nếu lỗi thì quay ra và đi tiếp
                driver.switch_to.default_content()
                continue

        print("[+] Đã gom xong toàn bộ dữ liệu. Bắt đầu bóc tách...")
        return combined_html
    finally:
        driver.quit()


def parse_fund_prices(html):
    """
    Parse bám sát vào định dạng của bảng giá thực sự (Có chữ Ngày và có chữ Quỹ/PRUlink).
    """
    soup = BeautifulSoup(html, "html.parser")
    rows = soup.find_all("tr")
    
    if not rows:
        print("[!] Không tìm thấy thẻ <tr> nào.")
        return {}

    fund_columns = None
    all_data = {}

    for tr in rows:
        cells = tr.find_all(["td", "th"])
        if len(cells) < 2:
            continue

        cell_texts = [_nfc(c.get_text(separator=" ", strip=True)) for c in cells]
        first_cell = cell_texts[0]

        # ĐIỀU KIỆN CHỐT HEADER CHẶT CHẼ:
        # 1. Ô đầu tiên chứa chữ "Ngày"
        # 2. Phải có từ 5 cột trở lên (bảng giá có khoảng 7-8 quỹ)
        # 3. Phải có chữ "Quỹ" hoặc "PRUlink" trong tên các cột
        if fund_columns is None and "Ngày" in first_cell and len(cell_texts) > 4:
            # Kiểm tra xem có cột nào chứa tên Quỹ không
            if any("Quỹ" in c or "PRUlink" in c for c in cell_texts):
                fund_columns = cell_texts[1:]
                print(f"  => [OK] Đã chốt ĐÚNG Header bảng giá: {fund_columns}")
                continue

        m = DATE_CELL_RE.search(first_cell)
        if not m:
            continue
        date_str = m.group(1)

        if fund_columns is None:
            continue
            
        # Nếu dòng dữ liệu này có số lượng cột không khớp với Header, bỏ qua (né bảng rác)
        if len(cell_texts) - 1 != len(fund_columns):
            continue

        row_prices = {}
        for fund_name, price_text in zip(fund_columns, cell_texts[1:]):
            clean_price = re.sub(r'[^\d\.,]', '', price_text)
            if not clean_price:
                continue
                
            try:
                price_val = float(clean_price.replace(".", "").replace(",", "."))
                row_prices[fund_name] = price_val
            except ValueError:
                pass

        if row_prices:
            all_data[date_str] = row_prices

    return all_data


def main():
    print("[*] Đang mở trình duyệt (undetected-chromedriver) và fetch dữ liệu từ Prudential (PRUlink) ...")
    try:
        html = fetch_rendered_html(headless=False)
    except Exception as e:
        print(f"[x] Lỗi khi mở trang: {e}")
        sys.exit(1)

    new_data = parse_fund_prices(html)

    if not new_data:
        print("[x] Không parse được dữ liệu giá nào. Có thể trang web đổi cấu trúc hoặc WAF vẫn đang chặn.")
        sys.exit(1)

    latest_date = max(new_data.keys(), key=lambda d: tuple(reversed(d.split("/"))))
    print(f"[*] Tìm được tổng {len(new_data)} ngày trong bảng. Ngày mới nhất: {latest_date}")
    for fund, price in new_data[latest_date].items():
        print(f"    {fund}: {price}")

    added = update_fund_excel(
        file_path=EXCEL_FILE,
        sheet_name=SHEET_NAME,
        new_data=new_data,
    )

    if added:
        print(f"[+] Đã thêm {added} dòng mới vào {EXCEL_FILE}")
    else:
        print("[=] Không có dòng mới (tất cả ngày đã tồn tại trong file).")


if __name__ == "__main__":
    main()