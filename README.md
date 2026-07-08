# Automation-Fund-Prices
Dự án bao gồm các script python tự động fetch giá cả 9 quỹ liên kết đơn vị tại Việt Nam và cập nhật định kỳ 2 lần 1 tuần vào 1 file excel duy nhất

# Cấu trúc thư mục
Dưới đây là mô tả chức năng của các file chính trong dự án:
1. `Fund_name.py`: Script tự động quét và thu thập dữ liệu unit/price của các fund.
2. `excel_fund_updater.py`: Script ghi nhớ format lưu trữ dữ liệu trong file excel.
3. `run_all.py`: Script để run tất cả scripts một lần - không cần chạy 9 scripts của các quỹ.
4. 'gia_don_vi_quy_TONGHOP.xlsx': File excel dùng để lưu thông tin scrape được từ quỹ theo định dạng nhất quán.
