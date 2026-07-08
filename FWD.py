# -*- coding: utf-8 -*-
import re
import sys
import time
import unicodedata

from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from excel_fund_updater import update_fund_excel


URL = (
    "https://www.fwd.com.vn/dich-vu-truc-tuyen/ivr/"
    "thong-tin-lai-suat-va-quy-lien-ket-don-vi/quy-lien-ket-don-vi"
)
EXCEL_FILE = "gia_don_vi_quy_TONGHOP.xlsx"  # file Excel gộp chung, đổi path nếu cần
SHEET_NAME = "FWD"  # sheet riêng cho FWD trong file gộp

# Map tên quỹ hiển thị trên web -> tên cột chính xác trong file Excel
FUND_NAME_MAP = {
    "Quỹ Vươn mình": "Quỹ Vươn Mình",
    "Quỹ Năng động": "Quỹ Năng Động",
    "Quỹ Tăng trưởng": "Quỹ Tăng Trưởng",
    "Quỹ Chiến lược": "Quỹ Chiến Lược",
    "Quỹ Cân bằng": "Quỹ Cân Bằng",
    "Quỹ Kết hợp": "Quỹ Kết Hợp",
    "Quỹ Ổn định": "Quỹ Ổn Định",
    "Quỹ Tích lũy": "Quỹ Tích Lũy",
}

# Chuẩn hóa key về NFC để so khớp ổn định: HTML lấy từ trang web có thể trả
# về chuỗi tiếng Việt ở dạng tổ hợp dấu NFD (nhìn giống NFC nhưng so sánh
# == sẽ ra False) -> nếu không chuẩn hóa, các tên quỹ có dấu phức (ư, ố, ế...)
# rất dễ bị "Bỏ qua" do không match được trong FUND_NAME_MAP.
FUND_NAME_MAP_NFC = {
    unicodedata.normalize("NFC", k): v for k, v in FUND_NAME_MAP.items()
}

DATE_CELL_RE = re.compile(r"(\d{2}/\d{2}/\d{4})")
PRICE_CELL_RE = re.compile(r"^[\d]{1,3}(?:\.\d{3})*$")


def _vn_price_to_float(price_str):
    """'9.377' -> 9377.0  (giá FWD không có phần lẻ thập phân)"""
    return float(price_str.replace(".", ""))


def fetch_rendered_html(url=URL, headless=True, wait_seconds=20):
    """
    Mở trang bằng Selenium, đợi bảng giá render xong rồi trả về page_source.
    headless=False để Tran có thể xem trực quan khi debug (theo thói quen
    đang dùng cho AIA/Hanwha).
    """
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1366,900")

    driver = webdriver.Chrome(options=options)
    try:
        driver.get(url)

        # Đợi đến khi bảng giá thực sự có dữ liệu (xuất hiện 1 ngày dd/mm/yyyy
        # trong bảng), thay vì sleep cứng -> tránh trường hợp mạng chậm.
        def _table_has_data(drv):
            text = drv.find_element(By.TAG_NAME, "body").text
            return bool(re.search(r"\d{2}/\d{2}/\d{4}", text)) and "No data found" not in text

        try:
            WebDriverWait(driver, wait_seconds).until(_table_has_data)
        except Exception:
            print("[!] Hết thời gian chờ bảng giá render, vẫn thử đọc dữ liệu hiện có...")

        # Đợi thêm chút để các dòng còn lại trong bảng kịp render hết
        time.sleep(1.5)

        return driver.page_source
    finally:
        driver.quit()


def _extract_date_and_price_from_cells(cell_texts):
    """
    Quét các ô (đã bỏ ô [0] = tên quỹ) để tìm ngày và giá theo FORMAT của
    chính nó, không quan tâm ô nào là "công ty quản lý quỹ" hay nó chứa gì.
    Trả về (date_str, price_float) hoặc (None, None) nếu không tìm đủ.
    """
    date_str = None
    price_val = None
    for text in cell_texts:
        text = text.strip()
        if not text:
            continue
        if date_str is None:
            m = DATE_CELL_RE.search(text)
            if m:
                date_str = m.group(1)
                continue
        if price_val is None and PRICE_CELL_RE.match(text):
            price_val = _vn_price_to_float(text)
    return date_str, price_val


def parse_fund_prices(html):
    """
    Parse theo cấu trúc bảng thật (<table>/<tr>/<td>).
    Trả về: (date_str, {fund_col_name: price})
    """
    soup = BeautifulSoup(html, "html.parser")

    prices = {}
    found_date = None

    rows = soup.find_all("tr")
    if not rows:
        # Fallback: trang có thể không dùng <table> thật mà render bằng div
        # (kiểu CSS-grid của Next.js) -> báo rõ để biết hướng xử lý tiếp.
        print(
            "[!] Không tìm thấy <tr> nào trong HTML -> bảng có thể không phải "
            "<table> thật mà là div/grid. Cần xem lại cấu trúc HTML thực tế "
            "(View Page Source / Inspect) để chỉnh selector."
        )
        return None, {}

    for tr in rows:
        cells = tr.find_all(["td", "th"])
        if len(cells) < 3:
            continue

        cell_texts = [c.get_text(separator=" ", strip=True) for c in cells]
        fund_name_raw = unicodedata.normalize("NFC", cell_texts[0].strip())

        if "Quỹ" not in fund_name_raw:
            continue  # dòng header hoặc không phải dòng dữ liệu quỹ

        mapped_name = FUND_NAME_MAP_NFC.get(fund_name_raw)
        if not mapped_name:
            print(f"[!] Không map được tên quỹ: {fund_name_raw!r} (bỏ qua dòng)")
            continue

        # Bỏ qua hoàn toàn ô [1] = "Công ty quản lý quỹ" (text SSIAM hay
        # ảnh logo VCBF đều không quan trọng) -> chỉ quét các ô còn lại để
        # tìm ngày + giá theo định dạng riêng của chúng.
        date_str, price_val = _extract_date_and_price_from_cells(cell_texts[1:])

        if date_str is None or price_val is None:
            print(f"[!] Bỏ qua dòng '{fund_name_raw}': không tìm đủ ngày/giá trong ô: {cell_texts}")
            continue

        prices[mapped_name] = price_val
        found_date = date_str

    return found_date, prices


def main():
    print(f"[*] Đang mở trình duyệt và fetch dữ liệu từ FWD ...")
    try:
        html = fetch_rendered_html(headless=True)
    except Exception as e:
        print(f"[x] Lỗi khi mở trang bằng Selenium: {e}")
        sys.exit(1)

    date_str, prices = parse_fund_prices(html)

    if not date_str or not prices:
        print(
            "[x] Không parse được dữ liệu giá. Có thể cấu trúc bảng đã đổi, "
            "hoặc trang chưa kịp render xong trước khi đọc."
        )
        sys.exit(1)

    expected_funds = set(FUND_NAME_MAP.values())
    missing = expected_funds - set(prices.keys())
    if missing:
        print(f"[!] Cảnh báo: thiếu giá các quỹ: {missing}")

    print(f"[*] Ngày định giá: {date_str}")
    for fund, price in prices.items():
        print(f"    {fund}: {price}")

    added = update_fund_excel(
        file_path=EXCEL_FILE,
        sheet_name=SHEET_NAME,
        new_data={date_str: prices},
    )

    if added:
        print(f"[+] Đã thêm {added} dòng mới vào {EXCEL_FILE}")
    else:
        print(f"[=] Không có dòng mới (ngày {date_str} đã tồn tại trong file).")


if __name__ == "__main__":
    main()