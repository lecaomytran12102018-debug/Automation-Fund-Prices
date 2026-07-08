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


URL = "https://generali.vn/lai-suat/quy-lien-ket-don-vi"

# Đổi path nếu file Excel không nằm cùng thư mục script.
# Lưu ý tên file có dấu "vị" (không phải "vi" thường) -- giữ đúng chính tả
# với file Excel hiện tại của bạn.
EXCEL_FILE = "gia_don_vi_quy_TONGHOP.xlsx"  # file Excel gộp chung, đổi path nếu cần
SHEET_NAME = "Generali"  # sheet riêng cho Generali trong file gộp

DATE_COL_HEADER = "Ngày"
DATE_CELL_RE = re.compile(r"(\d{2}/\d{2}/\d{4})")
# Giá Generali dùng dấu phẩy ngăn nghìn (vd "16,764"), khác FWD dùng dấu chấm.
PRICE_CELL_RE = re.compile(r"^[\d]{1,3}(?:,\d{3})*(?:\.\d+)?$")


def _vn_price_to_float(price_str):
    """'16,764' -> 16764.0"""
    return float(price_str.replace(",", ""))


def _nfc(text):
    return unicodedata.normalize("NFC", text or "").strip()


def fetch_rendered_html(url=URL, headless=True, wait_seconds=25):
    """
    Mở trang bằng Selenium, đợi bảng giá render xong, cuộn xuống đáy để
    đảm bảo lấy được dòng mới nhất, rồi trả về page_source.
    """
    options = webdriver.ChromeOptions()
    if headless:
        options.add_argument("--headless=new")
    options.add_argument("--window-size=1366,900")

    driver = webdriver.Chrome(options=options)
    try:
        driver.get(url)

        # Đợi đến khi bảng giá thực sự có dữ liệu (xuất hiện 1 ngày dd/mm/yyyy
        # trong bảng), thay vì sleep cứng.
        def _table_has_data(drv):
            text = drv.find_element(By.TAG_NAME, "body").text
            return bool(re.search(r"\d{2}/\d{2}/\d{4}", text))

        try:
            WebDriverWait(driver, wait_seconds).until(_table_has_data)
        except Exception:
            print("[!] Hết thời gian chờ bảng giá render, vẫn thử đọc dữ liệu hiện có...")

        # Thử bấm nút "TRA CỨU" để đảm bảo bảng query đúng theo khoảng ngày
        # đang hiển thị (mặc định web đã tự set 1 năm gần nhất, nhưng bấm
        # lại cho chắc, không lỗi nếu không tìm thấy nút).
        try:
            search_btn = WebDriverWait(driver, 5).until(
                EC.element_to_be_clickable(
                    (By.XPATH, "//*[contains(text(), 'TRA CỨU') or contains(text(), 'Tra cứu')]")
                )
            )
            search_btn.click()
            time.sleep(2)
        except Exception:
            pass  # không bắt buộc, bỏ qua nếu không có/không bấm được

        # Cuộn hết các vùng có thể cuộn trong trang xuống đáy vài lần, để
        # phòng trường hợp bảng dùng virtualized list (chỉ render dòng đang
        # nằm trong khung nhìn) -> không cuộn thì sẽ mất dòng mới nhất.
        _scroll_all_scrollables_to_bottom(driver)

        return driver.page_source
    finally:
        driver.quit()


def _scroll_all_scrollables_to_bottom(driver, max_iterations=15, pause=0.6):
    """
    Tìm mọi phần tử có thể cuộn (scrollHeight > clientHeight) trong trang và
    cuộn chúng xuống đáy, lặp lại vài lần để các dòng lazy-load kịp xuất
    hiện. Đồng thời cuộn cả window xuống đáy.
    """
    scroll_script = """
        let changed = false;
        const els = document.querySelectorAll('*');
        for (const el of els) {
            if (el.scrollHeight > el.clientHeight + 5) {
                const before = el.scrollTop;
                el.scrollTop = el.scrollHeight;
                if (el.scrollTop !== before) changed = true;
            }
        }
        const before = window.scrollY;
        window.scrollTo(0, document.body.scrollHeight);
        if (window.scrollY !== before) changed = true;
        return changed;
    """
    for _ in range(max_iterations):
        try:
            changed = driver.execute_script(scroll_script)
        except Exception:
            break
        time.sleep(pause)
        if not changed:
            break


def parse_fund_prices(html):
    """
    Parse theo cấu trúc bảng thật (<table>/<tr>/<td>).
    Tên cột (tên quỹ) được lấy TRỰC TIẾP từ header của bảng -> khớp thẳng
    với tên cột trong file Excel, không cần map tay.

    Trả về: dict { "dd/mm/yyyy": { "Tên cột quỹ": giá, ... }, ... }
    (đầy đủ tất cả các ngày tìm được trong bảng, không chỉ ngày mới nhất)
    """
    soup = BeautifulSoup(html, "html.parser")
    rows = soup.find_all("tr")
    if not rows:
        print(
            "[!] Không tìm thấy <tr> nào trong HTML -> bảng có thể không phải "
            "<table> thật mà là div/grid. Cần xem lại cấu trúc HTML thực tế."
        )
        return {}

    fund_columns = None  # list tên quỹ theo đúng thứ tự cột (đã bỏ cột Ngày)
    all_data = {}

    for tr in rows:
        cells = tr.find_all(["td", "th"])
        if len(cells) < 2:
            continue

        cell_texts = [_nfc(c.get_text(separator=" ", strip=True)) for c in cells]
        first_cell = cell_texts[0]

        # Dòng header: ô đầu tiên là "Ngày" -> các ô còn lại là tên quỹ.
        if fund_columns is None and first_cell == DATE_COL_HEADER:
            fund_columns = cell_texts[1:]
            continue

        # Dòng dữ liệu: ô đầu tiên match định dạng ngày dd/mm/yyyy.
        m = DATE_CELL_RE.match(first_cell)
        if not m:
            continue
        date_str = m.group(1)

        if fund_columns is None:
            # Chưa kịp thấy header nhưng đã thấy dòng dữ liệu -> bỏ qua,
            # không biết map cột nào ra quỹ nào.
            print(f"[!] Gặp dòng dữ liệu ({date_str}) trước khi xác định được header cột, bỏ qua.")
            continue

        row_prices = {}
        for fund_name, price_text in zip(fund_columns, cell_texts[1:]):
            price_text = price_text.strip()
            if not price_text or not PRICE_CELL_RE.match(price_text):
                continue  # cột trống (quỹ chưa định giá ngày đó) -> bỏ qua
            row_prices[fund_name] = _vn_price_to_float(price_text)

        if row_prices:
            all_data[date_str] = row_prices

    return all_data


def main():
    print(f"[*] Đang mở trình duyệt và fetch dữ liệu từ Generali ...")
    try:
        html = fetch_rendered_html(headless=True)
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

    # In ra ngày mới nhất tìm được để Tran dễ kiểm tra nhanh.
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