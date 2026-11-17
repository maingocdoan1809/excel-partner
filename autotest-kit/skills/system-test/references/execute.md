# Thực thi (`check`, `plan`, `run`, `retry`, `clean`)

Mọi lệnh chạy trong thư mục autotest (chứa `qa.config.json`). Biến môi trường viết dạng bash (`QA_OUT=… python -m …`); PowerShell: `$env:QA_OUT='…'; python -m …`. Môi trường khác URL trong config: `QA_BASE_URL=<url>` (và dữ liệu tương ứng trong hooks nếu cần). Mật khẩu / token chỉ lấy từ biến môi trường mà config đặt tên — không bao giờ in ra, không ghi vào file.

## Phạm vi → `qakit.tools.scope`
Phạm vi do user viết bằng tham số có tên (`file=` `id=` `suite=` `select=` `status=` `run=` `tag=` `env=` `role=` `exclude=` `sandbox=`; bảng ở SKILL.md) và đã qua `python -m qakit.tools.cmd` (bước 0). Truyền **đúng** các token ở dòng `scope:` của nó cho `python -m qakit.tools.scope` (dạng dài `file=` / `id=` / `suite=` / `exclude=` / `sandbox=` / `all` / `changed` / `failed@<run>` / `status=…@<run>`); không tự thêm token trần. Token chọn hợp (union) → lọc → loại trừ. Giá trị không khớp → "Not recognised", mã thoát 2 (không đoán). ID không có suite = **NOT AUTOMATED** (liệt kê, không chạy). Mã thoát 2 = giá trị không nhận ra, catalog cũ (`stale`), hoặc không còn gì để chạy. `url=` → đặt `QA_BASE_URL` cho mọi bước của lệnh.

## `plan <phạm vi>` (dry-run, chỉ đọc)
`python -m qakit.tools.scope <token scope:> --brief` → bản cho người đọc (theo màn hình, mỗi case bằng lời, vai trò, dữ liệu test sẽ tạo, phần không chạy + lý do); bản kỹ thuật (suite, slot, lệnh) khi bỏ `--brief`. Không tạo run, không mở trình duyệt.

## `check` (chỉ đọc)
1. `QA_OUT=<scratchpad>/check python -m qakit.tools.preflight` (không tạo thư mục trong `test_execution_output`). Kiểm: config, URL, trình duyệt (launch / cdp), đăng nhập, `hooks.check_data` (công cụ đọc ĐÚNG nguồn dữ liệu app dùng), snapshot, adapter vai trò / môi trường so với matrix, catalog.
2. `browser.mode = cdp`: trình duyệt test phải đang chạy với `--remote-debugging-port=<cdp_port> --user-data-dir=<profile riêng>`; đăng nhập MFA / SSO lần đầu: nhờ user đúng một lần.
3. Thiếu biến môi trường mật khẩu → báo tên biến cần đặt (không hỏi giá trị trong chat).
4. Báo vướng mắc: bảng mục · vấn đề · đã xác minh / chưa · ảnh hưởng · cần ai / quyền gì · hướng xử lý. Phân biệt rõ đã kiểm tra / giả định / sự cố thật.

## `run <phạm vi>` / `retry [run=<run_id>]`
`retry` = `run select=failed run=<run>` (scope `failed@<run>`, nhãn `retry-<run>`). Thiếu phạm vi → `qakit.tools.cmd` đã từ chối; không mặc định `all`.

**Giai đoạn và điều kiện qua bước** (ghi `reports/run_state.md` của run sau MỖI giai đoạn; sau khi ngữ cảnh bị nén → đọc lại rules.md + run_state.md trước khi làm tiếp). Thời gian ở mỗi mốc ghi `reports/metrics.md`.

| # | Giai đoạn | Việc | Điều kiện qua bước |
|---|---|---|---|
| 0 | Kế hoạch | `python -m qakit.tools.catalog validate`; `python -m qakit.tools.scope <token scope:> --out <scratchpad>/plan.json`; `python -m qakit.tools.runs new --plan <scratchpad>/plan.json --label <nhãn>` → `run_id`, thư mục run (= QA_OUT) | Plan không có token lạ, không ⛔ catalog cũ (→ `update` trước; `runs new` cũng từ chối); có ≥ 1 case chạy được |
| 1 | Preflight | như `check` với `QA_OUT=<thư mục run>` | Không lỗi chặn (URL, đăng nhập hoặc PENDING có chủ đích, nguồn dữ liệu) |
| 2 | Phân tích | Rà từng case trong plan: vai trò / môi trường (slot), dữ liệu tiền điều kiện (UI tạo được?), thao tác GHI (kể cả sinh dữ liệu phụ), cách quan sát, cản trở → nhóm **A** tự động · **B** sau chuẩn bị dữ liệu / quyền · **C** cần user quyết · **D** BLOCKED dự kiến (kèm điều kiện gỡ). Nhóm C: câu hỏi + rủi ro + mặc định an toàn. Đã có trong `decisions` của config → không hỏi lại. Ghi `reports/testcase_matrix.md` | Mọi case có nhóm |
| 3 | **Cổng xác nhận (một phiên)** | Theo [confirm.md](confirm.md): viết `<run>/reports/gate.md` (đầu ra `scope … --brief` + dữ liệu sẽ tạo, vai trò / môi trường sẽ đổi, dự kiến không kiểm được, việc cần người, quyết định nhóm C, EXTERNAL, thời gian) → hỏi **mỗi chủ đề một câu** (Phạm vi · Dữ liệu · Vai trò · từng quyết định), diễn đạt theo màn hình / thao tác, mã chỉ trong ngoặc. Câu trả lời ghi vào `gate.md`; trả lời "bỏ" → lập lại plan; quyết định lâu dài → `decisions` trong config. User hủy → `python -m qakit.tools.runs finish <run_id> --status cancelled` | Mọi câu của phiên đã trả lời; plan khớp câu trả lời; từ đây KHÔNG hỏi lại |
| 4 | Chạy | `python -m qakit.tools.autorun --plan <run>/plan.json` (+ `QA_ALLOW_EXTERNAL=1` chỉ khi được phép), chạy nền, rồi **đứng canh bằng `python -m qakit.tools.watch`** (không chỉ chờ thông báo): nó thoát ngay khi có case cần sửa (mã 10), sửa tại chỗ rồi chạy lại **đúng case đó trong cùng run** bằng `python -m qakit.tools.rerun <ID>` (ghi đè, không tạo run mới) và `rerun --pending` cho phần chưa tới; xem addin.md mục 1a. autorun: khóa → trình duyệt + đăng nhập + `check_data` → snapshot before → từng suite (`--matrix [--targets]`) → LUÔN: cleanup `--apply`, snapshot after + diff → Excel mỗi ngôn ngữ → bản tổng hợp → `reports/run_report.md` → index + `SUMMARY.md`. Runner đổi vai trò / môi trường theo slot qua adapter và khôi phục cuối cùng. Việc cần người → PENDING | Mọi case có trạng thái + lý do; mọi bước đã chạy có ảnh |
| 5 | Toàn vẹn | Đọc `logs/autorun_*.log`: cleanup, diff. Mọi CREATED / DELETED / CHANGED ngoài tiền tố test = vi phạm: báo ngay kèm id, KHÔNG tự sửa (có thể do người dùng thật đang dùng hệ thống). `data.snapshot = none` → ghi rõ "toàn vẹn dữ liệu chưa được kiểm". Còn mục MANUAL của cleanup → liệt kê | diff sạch hoặc đã báo; vai trò / môi trường đã khôi phục |
| 6 | Excel + nghiệm thu | report.md §Excel | Đã xem sheet tổng hợp + sheet có ảnh + sheet FAIL / BLOCKED / PENDING ở mọi bản ngôn ngữ |
| 7 | Báo cáo | `python -m qakit.tools.runs compare <run_id>` + report.md §Báo cáo A–F → `<run>/reports/report.md`; tóm tắt cho user, PENDING đầu tiên | Checklist của report.md đã tick |

**Chạy thủ công (khi không dùng autorun):** giữ thứ tự với `QA_OUT=<thư mục run>`: `python -m qakit.reporting.snapshot take before` → lệnh của từng item (`python -m suites.<suite>.cases --matrix <ID…> [--targets=env:role,…]`) → `python -m qakit.tools.cleanup --apply` → snapshot after + `diff before after` → `python -m qakit.reporting.excel_builder --suite <suite> --spec <spec> --only-run --lang <lang>` → `python -m qakit.tools.runs finish <run_id>`.

**Lỗi script ≠ lỗi ứng dụng:** case NOT EXECUTED do runner → sửa suite (automate.md), chạy lại bằng `retry` (run mới); không sửa / xóa kết quả của run cũ.

## `clean`
`python -m qakit.tools.cleanup` (kế hoạch, chỉ đọc: registry của các run + `hooks.discover()`) → trình bày theo màn hình và hỏi theo [confirm.md](confirm.md) §4 → `python -m qakit.tools.cleanup --apply`. Chỉ đối tượng có tiền tố test, qua `CLEANERS` của project hooks (thứ tự `CLEANUP_ORDER`); loại chưa có cleaner → MANUAL (liệt kê, đề xuất viết cleaner). Giữ evidence / log / báo cáo / snapshot.
