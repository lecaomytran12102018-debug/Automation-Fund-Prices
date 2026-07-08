import re
import time
from datetime import datetime
from pathlib import Path
import pandas as pd
import requests
from bs4 import BeautifulSoup

# ----------------------------- CONFIG ----------------------------------
BASE_URL = "https://www.chubb.com/vn-vn/customer-service/unit-price-notice/{year}.html"

# Dùng CHUNG 1 file Excel với các quỹ khác (FWD, Generali, AIA, PRUlink,
# Sunlife,...), mỗi quỹ 1 sheet riêng -> đổi path nếu file gộp không nằm
# cùng thư mục script.
OUTPUT_FILE = Path(__file__).with_name("gia_don_vi_quy_TONGHOP.xlsx")
SHEET_NAME = "Chubb"  # sheet riêng cho Chubb trong file gộp

START_YEAR = 2023  # năm bắt đầu lấy dữ liệu
DATE_COLUMN = "Ngày"
FUND_COLUMNS = ["Quỹ Tăng trưởng", "Quỹ Cân bằng", "Quỹ Bền vững"]
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/124.0 Safari/537.36"
    )
}
# ----------------------------------------------------------------------


def clean_number(text: str) -> float:
    text = re.sub(r"[^\d,.\-]", "", text)
    text = text.replace(".", "").replace(",", ".")
    return float(text)


def fetch_year_table(year: int) -> pd.DataFrame:
    url = BASE_URL.format(year=year)
    try:
        resp = requests.get(url, headers=HEADERS, timeout=30)
        resp.raise_for_status()
    except requests.RequestException as exc:
        print(f"[!] Không tải được năm {year}: {exc}")
        return pd.DataFrame(columns=[DATE_COLUMN, *FUND_COLUMNS])

    soup = BeautifulSoup(resp.text, "html.parser")
    table = soup.find("table")
    if table is None:
        print(f"[!] Không có dữ liệu năm {year}")
        return pd.DataFrame(columns=[DATE_COLUMN, *FUND_COLUMNS])

    rows = []
    for tr in table.find_all("tr"):
        cells = [td.get_text(strip=True) for td in tr.find_all(["td", "th"])]
        if len(cells) < 4:
            continue
        date_text = cells[0]
        if not re.match(r"^\d{2}/\d{2}/\d{4}$", date_text):
            continue
        try:
            date_val = datetime.strptime(date_text, "%d/%m/%Y")
            prices = [clean_number(c) for c in cells[1:4]]
        except ValueError:
            continue
        rows.append([date_val, *prices])

    return pd.DataFrame(rows, columns=[DATE_COLUMN, *FUND_COLUMNS])


def load_existing(path: Path, sheet_name: str) -> pd.DataFrame:
    """
    Đọc dữ liệu cũ TỪ ĐÚNG SHEET 'Chubb' trong file Excel gộp.
    Nếu file gộp chưa tồn tại, hoặc tồn tại nhưng chưa có sheet 'Chubb'
    (lần chạy đầu tiên) -> coi như chưa có dữ liệu cũ, trả về DataFrame rỗng.
    """
    if not path.exists():
        return pd.DataFrame(columns=[DATE_COLUMN, *FUND_COLUMNS])
    try:
        df = pd.read_excel(path, sheet_name=sheet_name)
    except ValueError:
        # Sheet "Chubb" chưa tồn tại trong file gộp -> chưa có dữ liệu cũ.
        return pd.DataFrame(columns=[DATE_COLUMN, *FUND_COLUMNS])
    df[DATE_COLUMN] = pd.to_datetime(df[DATE_COLUMN], dayfirst=True)
    return df


def get_years_to_fetch(existing: pd.DataFrame) -> list:
    current_year = datetime.now().year
    if existing.empty:
        return list(range(START_YEAR, current_year + 1))
    last_year = existing[DATE_COLUMN].max().year
    return sorted({current_year, last_year})


def save_to_shared_workbook(export_df: pd.DataFrame, path: Path, sheet_name: str):
    if path.exists():
        with pd.ExcelWriter(path, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
            export_df.to_excel(writer, sheet_name=sheet_name, index=False)
    else:
        with pd.ExcelWriter(path, engine="openpyxl", mode="w") as writer:
            export_df.to_excel(writer, sheet_name=sheet_name, index=False)


def main():
    existing = load_existing(OUTPUT_FILE, SHEET_NAME)
    years = get_years_to_fetch(existing)
    print(f"Sẽ tải các năm: {years}")

    frames = [existing]
    for year in years:
        print(f"--> Đang lấy năm {year}")
        df_year = fetch_year_table(year)
        print(f"   Lấy được {len(df_year)} dòng")
        frames.append(df_year)
        time.sleep(1)

    combined = pd.concat(frames, ignore_index=True)
    if combined.empty:
        print("Không có dữ liệu!")
        return

    combined.drop_duplicates(subset=[DATE_COLUMN], keep="last", inplace=True)
    combined.sort_values(DATE_COLUMN, ascending=False, inplace=True)
    combined.reset_index(drop=True, inplace=True)

    export_df = combined.copy()
    export_df[DATE_COLUMN] = pd.to_datetime(export_df[DATE_COLUMN], errors="coerce")
    export_df[DATE_COLUMN] = export_df[DATE_COLUMN].dt.strftime("%d/%m/%Y")

    save_to_shared_workbook(export_df, OUTPUT_FILE, SHEET_NAME)

    print(f"\n✅ Hoàn tất! Đã ghi vào sheet '{SHEET_NAME}' trong file {OUTPUT_FILE.name}")
    print(export_df.head(3))


if __name__ == "__main__":
    main()