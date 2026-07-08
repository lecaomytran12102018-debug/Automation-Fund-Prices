# -*- coding: utf-8 -*-
import re
import sys
from datetime import datetime

import requests
from bs4 import BeautifulSoup

from excel_fund_updater import update_fund_excel

import ssl
ssl._create_default_https_context = ssl._create_unverified_context


URL = "https://dai-ichi-life.com.vn/quy-lien-ket-don-vi-36"
EXCEL_FILE = "gia_don_vi_quy_TONGHOP.xlsx"  # file Excel gộp chung, đổi path nếu cần
SHEET_NAME = "Daiichi"  # sheet riêng cho Daiichi trong file gộp

# Map tên quỹ hiển thị trên web -> tên cột chính xác trong file Excel
# (web ghi "Quỹ Tăng trưởng" chữ thường, file Excel ghi "Quỹ Tăng Trưởng" chữ hoa,
#  nên cần map tay để tránh tạo nhầm cột mới)
FUND_NAME_MAP = {
    "Quỹ Cân Bằng": "Quỹ Cân Bằng",
    "Quỹ Tăng trưởng": "Quỹ Tăng Trưởng",
    "Quỹ Tăng Trưởng": "Quỹ Tăng Trưởng",
    "Quỹ Phát Triển": "Quỹ Phát Triển",
    "Quỹ Bảo Toàn": "Quỹ Bảo Toàn",
    "Quỹ Thịnh Vượng": "Quỹ Thịnh Vượng",
    "Quỹ Đảm Bảo": "Quỹ Đảm Bảo",
    "Quỹ Dẫn Đầu": "Quỹ Dẫn Đầu",
    "Quỹ Tài Chính Năng Động": "Quỹ Tài Chính Năng Động",
}

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}

# Regex bắt pattern: "Quỹ Cân Bằng 10.052,58 25/06/2026"
# (?!Quỹ) chặn việc nuốt lan qua chữ "Quỹ" khác phía trước (vd tiêu đề
# "Các Quỹ Liên Kết Đơn Vị" đứng ngay trước dòng "Quỹ Cân Bằng ...")
ROW_PATTERN = re.compile(
    r"(Quỹ\s+(?:(?!Quỹ)[^\d\n])+?)\s+([\d]{1,3}(?:\.\d{3})*,\d{2})\s+(\d{2}/\d{2}/\d{4})"
)


def _vn_price_to_float(price_str):
    """'10.052,58' -> 10052.58"""
    return float(price_str.replace(".", "").replace(",", "."))


def fetch_page_html(url=URL):
    resp = requests.get(
        url,
        headers=HEADERS,
        timeout=20,
        verify=False,   # disable SSL check
    )
    resp.raise_for_status()
    return resp.text

import urllib3
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)



def parse_fund_prices(html):
    """
    Trả về: (date_str, {fund_col_name: price})
    Thử parse bằng BeautifulSoup (tìm bảng giá) trước, nếu không ra đủ data
    thì fallback sang regex trên toàn bộ text của trang.
    """
    soup = BeautifulSoup(html, "html.parser")
    text = soup.get_text(separator=" ")
    text = re.sub(r"\s+", " ", text)  # gộp khoảng trắng/newline thành 1 space

    matches = ROW_PATTERN.findall(text)

    if not matches:
        return None, {}

    prices = {}
    found_date = None
    for raw_name, raw_price, raw_date in matches:
        fund_name = raw_name.strip()
        # Bỏ qua các đoạn match nhiễu không phải tên quỹ thật (phòng hờ)
        mapped_name = FUND_NAME_MAP.get(fund_name)
        if not mapped_name:
            continue
        prices[mapped_name] = _vn_price_to_float(raw_price)
        found_date = raw_date  # tất cả các quỹ đều có cùng ngày định giá

    return found_date, prices


def main():
    print(f"[*] Đang fetch dữ liệu từ {URL} ...")
    try:
        html = fetch_page_html()
    except requests.RequestException as e:
        print(f"[x] Lỗi khi fetch trang: {e}")
        sys.exit(1)

    date_str, prices = parse_fund_prices(html)

    if not date_str or not prices:
        print(
            "[x] Không parse được dữ liệu giá từ trang (có thể trang cần JS render). "
            "Thử lại bằng Selenium nếu requests.get() không đủ."
        )
        sys.exit(1)

    expected_funds = set(FUND_NAME_MAP.values())
    missing = expected_funds - set(prices.keys())
    if missing:
        print(f"[!] Cảnh báo: thiếu giá các quỹ: {missing}")

    print(f"[*] Ngày định giá: {date_str}")
    for fund, price in prices.items():
        print(f"    {fund}: {price}")

    new_data = {date_str: prices}

    added = update_fund_excel(
        file_path=EXCEL_FILE,
        sheet_name=SHEET_NAME,
        new_data=new_data,
    )

    if added:
        print(f"[+] Đã thêm {added} dòng mới vào {EXCEL_FILE}")
    else:
        print(f"[=] Không có dòng mới (ngày {date_str} đã tồn tại trong file).")


if __name__ == "__main__":
    main()