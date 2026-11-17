# Quản lý Test Suite (`update`, `suite …`)

Test Suite = catalog `test_suite/catalog.json` do `python -m qakit.tools.catalog` sinh từ các workbook nguồn (glob trong `test_suite/sources.json`, mặc định `testcases.sources` của config; file đứng trước là **nguồn chính** khi một ID có trong nhiều file). File nguồn không bao giờ bị ghi.

Mỗi ID lưu: file / sheet / dòng, phân loại (cls), loại (typ), cột vai trò / môi trường + các slot `env:role` runner sẽ chạy, số bước, hash nội dung (quan điểm, vai trò, chuẩn bị, thao tác, mong đợi, phân loại), phiên bản file (ô "Version / Phiên bản" ở sheet bìa), suite tự động hóa (`automation`: suite, `matrix|plain`, `primary`), trạng thái vòng đời, `first_seen`, `history`.

**Catalog cũ (stale):** mỗi build lưu hash từng file nguồn. File lưu lại sau build / file mới / file mất → `catalog validate` báo ERROR, `qakit.tools.scope` đánh dấu `stale` + mã thoát 2 khi file liên quan tới phạm vi (`changed` / `all`: mọi file), `runs new` từ chối. Cách gỡ duy nhất: `update`.

**Vòng đời ID:** `new` (xuất hiện sau lần build đầu) → `active` → `changed` (hash khác; đã có automation thì `needs_review`) → `deprecated` (không còn trong nguồn nào; GIỮ NGUYÊN, không xóa; xuất hiện lại → `active`, ghi `reactivated`). ID luôn là ID trong file testcase, không đánh số lại.

**Sandbox `<tên>`:** catalog riêng `test_suite_<tên>/` (có `sources.json` riêng), không đụng `test_suite/` — dùng cho demo / thử nghiệm file testcase trước khi đưa vào catalog thật. Tạo: thư mục `test_suite_<tên>/sources.json` = `{"sources": ["<glob>"]}`. Lệnh: `qakit.tools.catalog --sandbox <tên> …`, `qakit.tools.scope … <tên>`.

## `update [<sandbox>] [note="…"]` (tạo / cập nhật)
1. File testcase mới đặt đúng thư mục nguồn (hoặc thêm glob vào `test_suite/sources.json`, có lý do). Không sửa file testcase.
2. `python -m qakit.tools.catalog build --dry-run` → đọc new / changed / deprecated / reactivated / needs_review / conflict.
3. Agent xem thay đổi **mơ hồ**; mỗi loại cần quyết định là một câu hỏi riêng, trình bày theo màn hình + tên case (mã trong ngoặc), nội dung đổi dạng trước → sau ([confirm.md](confirm.md) §4):
   - ID bị xóa mà vẫn có automation → giữ `deprecated` (mặc định) hay user xác nhận đổi ID (ghi vào note, không tự gộp).
   - ID đổi nội dung mà có automation (`needs_review`) → so nội dung cũ (`test_suite/history/<build trước>.json`) với mới; automation còn đúng → `ack`; không → đưa vào `automate`.
   - `conflict` (cùng ID, nội dung khác giữa các file) → thường là bộ con có chủ đích: ghi nhận.
   - File mới có ID trùng file khác / thiếu cột bắt buộc → báo user (gợi ý thêm tên cột vào `testcases.columns`), không đoán.
4. User đồng ý lưu (câu hỏi cuối của confirm.md §4) → `python -m qakit.tools.catalog build --note "<vì sao: file nào, phiên bản>"` → ghi `catalog.json`, `history/<build>.json`, `CHANGELOG.md`.
5. `python -m qakit.tools.catalog validate`; lỗi → báo kèm cách sửa (cột vai trò không nhận ra → thêm từ khóa vào `matrix`).
6. Đã review xong ID `needs_review`: `python -m qakit.tools.catalog ack <ID…> --note "<kết luận>"`.
7. Trình bày chênh lệch (`python -m qakit.tools.catalog diff`) và báo: số new / changed / deprecated, ID cần review / cần automation, lệnh `plan select=changed`.

Mọi lệnh `qakit.tools.catalog` của `update <sandbox>` thêm `--sandbox <sandbox>`.

## `suite validate`
`python -m qakit.tools.catalog validate` (mã thoát 1 khi có ERROR): file nguồn đổi sau build, suite không nạp được, spec thiếu, suite không đăng ký case, `t.todo()` còn lại, ID của suite không có trong spec, ID deprecated vẫn có automation, `needs_review`, case thiếu bước, cột vai trò / môi trường không nhận ra, nhiều suite chính cho một ID, ID trùng nội dung khác. Agent tóm tắt theo nhóm và đề xuất xử lý; không tự sửa testcase / suite khi chưa được yêu cầu.

## `suite show [<phạm vi>]`
`python -m qakit.tools.catalog show [--file 03] [--id …]` → bảng theo file (phiên bản, số ID theo trạng thái, đã tự động hóa, cần review) + suite; `--file` liệt kê từng ID; `--id` in bản ghi đầy đủ.

## `suite diff [base=<build>]`
`python -m qakit.tools.catalog diff [--base …]` → thêm / bỏ / đổi nội dung / đổi trạng thái / đổi automation giữa hai build.

## `suite coverage`
`python -m qakit.tools.catalog coverage` → tỉ lệ ID active có automation theo file và theo loại. Đây là độ phủ **tự động hóa**; độ phủ yêu cầu ↔ testcase chỉ đánh giá khi user yêu cầu (đối chiếu `docs/system_map/*` / tài liệu thiết kế), nêu rõ là đánh giá của agent.

## `suite impact <thay đổi>` (chỉ đọc)
Cần `source.use = true` (black-box → từ chối như `source`). Là bước 1–4 của `source` ([update-source.md](update-source.md)) mà không sửa gì, không chạy: `python -m qakit.tools.impact --range <repo>=<range> …` (bảng ánh xạ `test_suite/source_map.json`) → agent thu hẹp → bảng ID · lý do · mức (trực tiếp / hồi quy) · có automation · code test có khả năng gãy; gợi ý `plan id=<ID,…>` / `automate id=<ID,…>` / `source ref=…`. Nêu rõ phần nào là suy luận.
