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


URL = "https://www.sunlife.com.vn/vn/dich-vu-khach-hang/lai-suat-va-quy/gia-don-vi-quy/"
EXCEL_FILE = "gia_don_vi_quy_TONGHOP.xlsx"  # file Excel gộp chung, đổi path nếu cần
SHEET_NAME = "Sunlife"  # sheet riêng cho Sunlife trong file gộp

DATE_FULL_RE = re.compile(r"^\d{2}/\d{2}/\d{4}$")
DATE_CELL_RE = re.compile(r"(\d{2}/\d{2}/\d{4})")
PRICE_CELL_RE = re.compile(r"^[\d]{1,3}(?:\.\d{3})*(?:,\d+)?$")


def _nfc(text):
    return unicodedata.normalize("NFC", text or "").strip()


def _vn_price_to_float(price_str):
    """'22.111' -> 22111.0"""
    return float(price_str.replace(".", "").replace(",", "."))


def fetch_rendered_html(url=URL, headless=False, max_wait_seconds=45, poll_interval=2):
    """
    Mở trang bằng Selenium. Thay vì chỉ chờ "có ngày dd/mm/yyyy nào đó xuất
    hiện trong trang" (dễ bị đánh lừa bởi các ngày khác ở banner/footer
    trong khi bảng giá thật vẫn đang loading), hàm này LẶP LẠI VIỆC THỬ
    PARSE THẬT bảng giá nhiều lần cho tới khi ra được dữ liệu hoặc hết giờ.
    """
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1366,1200")

    driver = webdriver.Chrome(options=options)
    try:
        driver.get(url)

        # Cuộn tới khu vực có chữ "Giá đơn vị quỹ" / tên 1 quỹ đã biết, để
        # kích hoạt lazy-load nếu trang chỉ load bảng khi nó lọt vào khung
        # nhìn (rất phổ biến ở các trang dùng AEM/Next.js).
        try:
            anchor = WebDriverWait(driver, 10).until(
                EC.presence_of_element_located(
                    (By.XPATH, "//*[contains(text(), 'Giá đơn vị quỹ') or contains(text(), 'Quỹ liên kết đơn vị')]")
                )
            )
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", anchor)
        except Exception:
            pass

        elapsed = 0
        last_html = driver.page_source
        while elapsed <= max_wait_seconds:
            html = driver.page_source
            last_html = html
            soup = BeautifulSoup(html, "html.parser")
            data = _parse_via_table(soup) or _parse_via_flatten_text(soup)
            if data:
                print(f"[*] Bảng giá đã render xong sau ~{elapsed}s.")
                return html

            # Cuộn thêm 1 chút mỗi lần chờ, đề phòng cần cuộn thật mới load.
            try:
                driver.execute_script("window.scrollBy(0, 200);")
            except Exception:
                pass

            time.sleep(poll_interval)
            elapsed += poll_interval

        print(f"[!] Hết {max_wait_seconds}s chờ mà vẫn chưa parse được bảng giá, dùng HTML hiện tại để debug.")
        return last_html
    finally:
        driver.quit()


def _parse_via_table(soup):
    """Cách 1: parse theo cấu trúc bảng thật (<table>/<tr>/<td>)."""
    rows = soup.find_all("tr")
    if not rows:
        return {}

    date_columns = None
    fund_rows = {}

    for tr in rows:
        cells = tr.find_all(["td", "th"])
        if len(cells) < 2:
            continue

        cell_texts = [_nfc(c.get_text(separator=" ", strip=True)) for c in cells]

        date_like_cells = [c for c in cell_texts[1:] if DATE_FULL_RE.match(c)]
        if date_columns is None and len(date_like_cells) >= 2:
            date_columns = date_like_cells
            continue

        first_cell = cell_texts[0]
        if "Quỹ" in first_cell:
            fund_rows[first_cell] = cell_texts[1:]

    if date_columns is None or not fund_rows:
        return {}

    all_data = {}
    for fund_name, price_list in fund_rows.items():
        for date_str, price_text in zip(date_columns, price_list):
            price_text = price_text.strip()
            if not price_text or not PRICE_CELL_RE.match(price_text):
                continue
            all_data.setdefault(date_str, {})[fund_name] = _vn_price_to_float(price_text)
    return all_data


def _parse_via_flatten_text(soup):
    """
    Cách 2 (fallback): không quan tâm bảng dựng bằng <table> hay <div>/<grid>
    gì cả -> chỉ lấy TOÀN BỘ đoạn text hiển thị trên trang theo ĐÚNG THỨ TỰ
    xuất hiện (giống thứ tự đọc từ trên xuống, trái sang phải), rồi:
      1. Tìm dãy token liên tiếp đầu tiên là ngày dd/mm/yyyy -> đó là header
         các cột ngày.
      2. Sau đó, mỗi khi gặp 1 token chứa "Quỹ" (không phải số) -> coi là
         tên quỹ mới, gom N token số tiếp theo (N = số cột ngày) làm giá
         tương ứng theo đúng thứ tự cột ngày đã xác định ở bước 1.
    Cách này bền với mọi kiểu HTML (table, div, css-grid...) vì chỉ dựa vào
    THỨ TỰ xuất hiện của text, không dựa vào tag/class cụ thể.
    """
    tokens = [_nfc(t) for t in soup.stripped_strings if t.strip()]

    date_columns = []
    start_idx = None
    for i, t in enumerate(tokens):
        if DATE_FULL_RE.match(t):
            if start_idx is None:
                start_idx = i
            date_columns.append(t)
        elif start_idx is not None:
            break

    if not date_columns:
        return {}

    n_cols = len(date_columns)
    all_data = {}
    current_fund = None
    prices_buffer = []

    def _flush():
        if current_fund and prices_buffer:
            for date_str, price_text in zip(date_columns, prices_buffer):
                if PRICE_CELL_RE.match(price_text):
                    all_data.setdefault(date_str, {})[current_fund] = _vn_price_to_float(price_text)

    for t in tokens[start_idx + n_cols:]:
        if "Quỹ" in t and not PRICE_CELL_RE.match(t):
            _flush()
            current_fund = t
            prices_buffer = []
        elif current_fund and PRICE_CELL_RE.match(t) and len(prices_buffer) < n_cols:
            prices_buffer.append(t)
    _flush()

    return all_data


def parse_fund_prices(html):
    """
    Parse bảng XOAY NGƯỢC (dòng = quỹ, cột = ngày) rồi trả về dict đã xoay
    lại đúng format chuẩn của update_fund_excel():
        { "dd/mm/yyyy": { "Tên cột quỹ": giá, ... }, ... }

    Thử cách 1 (parse theo <table>) trước; nếu không ra kết quả (bảng dựng
    bằng div/grid, không có <tr> thật) thì fallback sang cách 2 (đọc theo
    thứ tự text thô).
    """
    soup = BeautifulSoup(html, "html.parser")

    all_data = _parse_via_table(soup)
    if all_data:
        return all_data

    print("[!] Không parse được qua <tr>, thử fallback đọc theo thứ tự text thô...")
    all_data = _parse_via_flatten_text(soup)
    if all_data:
        return all_data

    # Vẫn không ra gì -> lưu HTML lại để debug, gửi cho Tran xem cấu trúc thật.
    debug_path = "sunlife_debug.html"
    try:
        with open(debug_path, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"[!] Đã lưu HTML trang lúc đọc vào '{debug_path}' để debug. Gửi file này lại để kiểm tra cấu trúc thật.")
    except Exception as e:
        print(f"[!] Không lưu được file debug: {e}")

    return {}


def main():
    print(f"[*] Đang mở trình duyệt và fetch dữ liệu từ Sun Life ...")
    try:
        html = fetch_rendered_html(headless=False)
    except Exception as e:
        print(f"[x] Lỗi khi mở trang bằng Selenium: {e}")
        sys.exit(1)

    new_data = parse_fund_prices(html)

    if not new_data:
        print(
            "[x] Không parse được dữ liệu giá nào. Có thể cấu trúc bảng đã đổi, "
            "hoặc trang chưa kịp render xong trước khi đọc."
        )
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
        print(f"[=] Không có dòng mới (tất cả ngày đã tồn tại trong file).")


if __name__ == "__main__":
    main()