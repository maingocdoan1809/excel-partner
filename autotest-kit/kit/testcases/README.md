# Testcases

Đặt các file testcase Excel (`.xlsx`) của dự án vào đây. Lệnh chọn file bằng **tên file đầy đủ** hoặc đường dẫn (`/system-test run file=Testcase_01_Login.xlsx`), không có viết tắt; nên đặt tên rõ ràng và không đổi tên giữa các lần chạy.

Mỗi file cần một bảng có **cột ID** và **cột kết quả mong đợi** (hoặc thao tác). Tên cột được nhận theo từ đồng nghĩa tiếng Việt / Anh / Nhật: `ID`, `No.`, `Kết quả mong đợi`, `Expected Result`, `期待結果`, `Thao tác`, `Steps`, `Role`, `Vai trò`, `Điều kiện tiên quyết`, `Phân loại`, `Loại`… Tên cột riêng của dự án: thêm vào `qa.config.json` mục `testcases.columns`.

Cột vai trò / môi trường (nếu có) quyết định case chạy với những vai trò nào: từ khóa khai báo ở `qa.config.json` mục `matrix`.

File ở đây **không bao giờ bị công cụ ghi**: đọc để lập catalog, kết quả ghi ra file Excel evidence riêng.
