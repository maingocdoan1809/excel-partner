# Thiết kế + cài đặt automation cho case (`automate <phạm vi>`)

Mục tiêu: case trong phạm vi có code chạy được trong một suite, đúng testcase, mỗi thao tác một ảnh khoanh đỏ. Không chạy thật toàn bộ ở đây (đó là `run`); chỉ chạy thử trên trạng thái hiện tại (`--dev`).

## 1. Xác định việc cần làm
- `python -m qakit.tools.scope <phạm vi>` → các ID NOT AUTOMATED và file nguồn. Một file testcase ↔ một suite chính (`suites/<tên>/cases.py`).
- File chưa có suite: `python -m qakit.tools.gen_suite --spec testcases/<file>.xlsx --name <suite> [--ids …]` → khung `cases.py` (mỗi bước là `t.todo`, hiện NOT EXECUTED, không bao giờ thành PASS).
- File đã có suite: thêm hàm case vào suite đó, không tạo suite trùng.

## 2. Nắm ứng dụng + phân tích (trước khi viết)
- **`source.use = false` (black-box, G12):** không đọc source. Căn cứ theo thứ tự: testcase → `inputs/autotest_input.xlsx` (sheet 4 màn hình, 5 phân quyền, 6 dữ liệu, 7 quy tắc + thông báo nguyên văn) → `docs/system_map/*` → DOM thật (`dom_probe`) và thao tác chỉ đọc trên trình duyệt test. Thông báo lỗi chưa có nguyên văn ở đâu → kỳ vọng theo đúng chữ testcase, không tự viết; quy tắc chưa rõ → ghi vào câu hỏi của cổng `run`.
- **`source.use = true`:** đọc đúng phần liên quan của app: route / quyền theo vai trò, màn hình / component, API, model, validation và **thông báo lỗi nguyên văn**, thao tác nào GHI dữ liệu, cái gì quan sát được và cái gì **không** (không quan sát được → case BLOCKED từ đầu). `docs/system_map/*` trước, source sau.
- Dò DOM chỉ đọc: `python -m qakit.tools.dom_probe /route …` → `docs/dom/<route>.json` (heading, nút, input, tab, bảng, dialog + thuộc tính làm selector ổn định).
- Ghi khác biệt spec ↔ thực tế vào remark / báo cáo; không tự sửa testcase.
- Dữ liệu tiền điều kiện: cái gì UI tạo được (dữ liệu test `<prefix><suite>_<ID>`), cái gì không (hỏi ở cổng của `run`, hoặc PENDING / BLOCKED). File / câu hỏi mẫu do user cung cấp dùng nguyên văn; thiếu → PENDING nêu đúng thứ còn thiếu, không tạo nội dung giả.

## 3. Viết case
- Khai báo `@suite.case("ORD-001")`. Suite matrix (theo cột vai trò / môi trường): hàm `fn(t, role, env)`, chạy bằng `--matrix` (runner đổi vai trò / môi trường theo slot qua adapter). Suite thường: `fn(t)`, vai trò riêng: `@suite.case(id, role="user")`.
- **Page object** trong `pages/<màn hình>.py` (kế thừa `pages/base.py`): ẩn selector, mọi thao tác đi qua `t` để tự chụp (`t.click`, `t.fill`, `t.hover`, `t.snap`), mọi thao tác GHI gọi `guard.check(guard.WRITE_TEST, name, "…")` trước và `guard.track(kind, name)` sau khi tạo. Selector ưu tiên role + tên, label, `data-testid`, id; chữ hiển thị qua bộ nhãn đa ngôn ngữ (`ui.button(scope, ("Save", "Lưu", "保存"))`). Mong đợi luôn lấy từ testcase.
- **Viết từng case một và chạy thử từng case một** ([addin.md](addin.md) mục 1a): không viết cả suite rồi mới chạy. Khi cần một điều kiện biên chưa có trên màn hình dùng `t.reach(...)`.
- **Bước trong testcase là đường lý tưởng; code test phải tự xử lý vướng mắc** (chi tiết ở [repair.md](repair.md)): thao tác Lưu / Gửi dùng `t.submit(...)` (chụp lúc bấm, đọc validation, tự điền ô bắt buộc, bấm lại); khi testcase cần dữ liệu trên màn hình dùng `t.ensure(...)` để tạo nếu chưa có (không đặt BLOCKED); ghi FAIL do lỗi ứng dụng chỉ bằng `t.defect(..., ruled_out=[...])`.
- Mỗi bước: đọc action + expected → thao tác THẬT → chờ trạng thái hoàn tất → xác minh (UI + **đọc lại dữ liệu** qua API / hooks cho mọi kỳ vọng dữ liệu) → `t.step(n, caption, action, ok, actual, hl=[selector])`. Không làm bước phụ thuộc khi trạng thái sai.
- Không quan sát được → `t.blocked(n, caption, action, "<lý do + điều kiện gỡ>")`. Cần người → `raise Pending(reason, needs=...)` hoặc `t.pending(...)`. Case idempotent (dọn dư trước khi tạo, tên có thời gian). Loại đối tượng mới mà test tạo ra → thêm cleaner vào `CLEANERS` của `project/hooks.py` (và `discover()`).

## 4. Chuẩn evidence (bắt buộc)
- **Khoanh ĐÚNG phần tử mà bước nói tới, không khoanh cả bảng / cả vùng.** Kiểm tra "tên dài bị cắt" → khoanh ô tên đó; "có nút New" → khoanh nút New; "tag nằm trong ô" → khoanh ô chứa tag; "danh sách có 2 dòng" → khoanh hai dòng. Khung quanh cả trang hoặc cả bảng rộng hết màn hình vô nghĩa với người đọc báo cáo. Bước không truyền `hl=` bị đánh dấu "NO RED BOX" trong kết quả và dashboard: sửa trước khi coi case là xong.
- **Mỗi thao tác nhỏ = một ảnh có khung đỏ đúng phần tử**: mỗi click, mỗi ô nhập (một ảnh / field, chú thích ghi giá trị; chuỗi dài rút gọn + số ký tự), bật / tắt switch / checkbox (ảnh trước + sau), chọn ngày / dropdown, mở / đóng dialog, xác nhận lưu / xóa. Thao tác tiền điều kiện qua UI cũng chụp. Không gộp nhiều field vào một ảnh.
- Ảnh kết quả cuối bước qua `t.step(...)` (đánh số `n-k`); nhãn ảnh là ghi chú bước, không phải tên file. Toast biến mất nhanh: đọc rồi chụp ngay (`wait<=100`, `scroll=False`). Chờ trạng thái hoàn tất thật (hết loading). Khung đỏ theo `bounding_box()` DOM; phần tử ngoài viewport được cuộn vào giữa.
- Kỳ vọng về dữ liệu không thấy trên UI: hiển thị kết quả đọc (chỉ đọc) trong chính trình duyệt test (trang API JSON / màn hình quản trị chỉ đọc), ghi rõ trong caption.
- Caption / actual viết theo `language.testcase`, rõ để dịch 1-1; ngôn ngữ Excel khác → bản dịch trong `suites/<suite>/<lang>_text.py` (report.md).

## 5. Chạy thử + đăng ký
- Chạy thử trên trạng thái hiện tại, không đổi vai trò / môi trường: `QA_OUT=test_execution_output/adhoc/dev_<ID> python -m suites.<suite>.cases --dev=<env>:<role> <ID…>` (trong add-in: thư mục này hiện ở dashboard `/tests`; báo giai đoạn bằng `python -m qakit.tools.status trial "…"`); chỉ đọc hoặc chỉ dữ liệu test; xem ảnh trước khi kết luận. Kết quả chạy thử không phải kết quả kiểm thử. Dọn dữ liệu test của lần thử (`python -m qakit.tools.cleanup --apply` với cùng QA_OUT).
- Case tạo dữ liệu test → ghi một dòng mô tả bằng lời vào `suites/<suite>/notes.json` (`{"<ID>": {"data": ["…"]}}`) để phiên xác nhận hiển thị; file testcase mới → tên màn hình trong `test_suite/screens.json`.
- Xong: `python -m qakit.tools.catalog build --note "automate <ID…>"`; `catalog validate` sạch lỗi của suite vừa sửa. Báo user ID đã tự động hóa, ID còn `t.todo` / BLOCKED dự kiến, lệnh `plan <ID,…>` để chạy.
