# Luật chung (nguồn duy nhất; mọi lệnh của /system-test)

Khi hai luật mâu thuẫn: an toàn dữ liệu / riêng tư thắng. Reference của từng lệnh chỉ bổ sung, không nới lỏng các luật này. Các giá trị riêng của dự án (URL, tiền tố dữ liệu test, vai trò, môi trường, ngôn ngữ) đọc từ `qa.config.json` của thư mục autotest; không đoán.

## G1–G12

- **G1 Không mock, không suy đoán.** Trạng thái chỉ từ assertion trên UI + dữ liệu đọc lại (DB / API qua project hooks). Bước không có ảnh thật = không được ghi nhận (runner ném lỗi). Thiếu căn cứ → BLOCKED / PENDING / NOT EXECUTED. Không tạo evidence giả, không chạy trên localhost / HTML mock thay cho môi trường đã chọn.
- **G2 Toàn vẹn dữ liệu.** Dữ liệu có sẵn chỉ ĐỌC. Chỉ ghi vào đối tượng do test tạo, tên bắt đầu bằng `data.test_prefix` (mặc định `ZZ_AUTOTEST_`), qua UI (hoặc API công khai của app khi dọn). Gọi `guard.check(...)` trước mọi thao tác ghi. **Thao tác sinh dữ liệu phụ cũng là GHI**: gửi chat / AI generate / chạy đánh giá / gửi mail / thanh toán. Dialog chọn sẵn "tất cả" → kiểm số bản ghi (`guard.generate_scope`) trước khi bấm. Không ghi trực tiếp vào DB.
- **G3 Riêng tư.** Chỉ chụp trang của trình duyệt test (`page.screenshot`). Cấm chụp màn hình / cửa sổ khác; cấm đọc cookie / localStorage / token để làm bằng chứng. Lỡ thấy nội dung riêng tư → xóa ảnh ngay, báo user.
- **G4 Bị chính sách / hệ thống từ chối → dừng đúng hành động đó**, không lặp lại, không thử biến thể; làm tiếp phần độc lập; xin **đúng một** quyền cần thiết. Không vượt MFA / CAPTCHA, không dùng tài khoản không được cấp.
- **G5 Phạm vi.** Chỉ URL của môi trường đã chọn; chỉ testcase trong phạm vi đã giải (`qakit.tools.scope`); không ghi đè file testcase nguồn, file mẫu hay file người dùng đang mở (`~$*.xlsx` → lưu bản có hậu tố thời gian).
- **G6 Một cổng xác nhận cho mỗi lần chạy.** Với `run` / `retry`: một phiên xác nhận ở cổng (execute.md; cách hỏi: confirm.md — mỗi chủ đề một câu, diễn đạt theo màn hình / thao tác, mã chỉ trong ngoặc). Sau phiên đó chạy liền mạch, không hỏi lại; việc cần người → **PENDING** kèm "NEEDS / CẦN: …"; tình huống ngoài dự kiến → mặc định an toàn (không ghi dữ liệu có sẵn, không EXTERNAL). Lệnh chỉ đọc không hỏi (trừ `update` trước khi lưu catalog và `clean` trước khi dọn, theo confirm.md §4).
- **G7 Đổi vai trò / môi trường chỉ qua adapter đã cấu hình** (`roles.switch`, `envs.switch` trong `qa.config.json`): tài khoản theo vai trò, hoặc script của dự án. Sao lưu trước lần đổi đầu nếu script đổi dữ liệu dùng chung (`hooks.backup_roles`), luôn khôi phục cuối cùng (kể cả khi lỗi) và kiểm lại. Không có adapter → case của vai trò / môi trường đó là PENDING, không tự nghĩ cách khác.
- **G8 Xóa an toàn.** Chỉ xóa đường dẫn cụ thể do agent tạo, liệt kê từng file. Không xóa đệ quy thư mục chứa dữ liệu người dùng. Luôn giữ: file user cung cấp, evidence / log / báo cáo của mọi lần chạy, `test_suite/history/`.
- **G9 Trung thực.** Không đánh đồng lỗi môi trường / script với lỗi chức năng; không khẳng định hoàn thành / 100% khi chưa kiểm bằng công cụ thật; báo lỗi của chính mình ngay, đo mức ảnh hưởng trước khi báo, không tự "sửa thêm" khi chưa được phép.
- **G10 Loại trừ.** Thao tác EXTERNAL (tạo tài nguyên ngoài app, gửi mail / ticket thật, thanh toán) chỉ chạy khi user cho phép ở cổng (`QA_ALLOW_EXTERNAL=1`); chức năng user chỉ định loại trừ → NOT EXECUTED, lý do "loại trừ theo chỉ định". Quyết định đã chốt của user ghi ở `qa.config.json` mục `decisions` (tạo khi cần) để không hỏi lại.
- **G11 Không ghi đè lịch sử.** Mỗi lần chạy ghi vào `test_execution_output/runs/<run_id>/` (QA_OUT). Không sửa / xóa run đã xong (chỉ khi user yêu cầu dọn). Test Suite không bao giờ xóa ID: ID không còn trong nguồn → `deprecated`.
- **G12 Không dùng source khi `source.use = false`** (chế độ black-box, `qa.config.json`). Không mở, tìm (Read / Grep / Glob / `cat` / `git`) bất kỳ file nào của ứng dụng, kể cả khi code nằm ngay trong thư mục dự án; chỉ dùng thư mục autotest (`inputs/`, `testcases/`, `docs/system_map/`, `docs/dom/`, `pages/`, `suites/`, `project/`) và trình duyệt test. Kỳ vọng / thông báo nguyên văn lấy từ testcase và `inputs/autotest_input.xlsx` (sheet 4–7); selector lấy từ DOM thật (`qakit.tools.dom_probe`). Không đủ căn cứ → hỏi ở cổng hoặc BLOCKED / PENDING, không đọc source để lấp. Lệnh `source` / `suite impact` từ chối và nói rõ lý do. Muốn đọc source → user đổi `source.use` thành `true`.

## Thứ tự ưu tiên khi thực thi
1 Toàn vẹn dữ liệu + riêng tư · 2 Đúng testcase và kết quả thực tế · 3 Đủ evidence từng thao tác · 4 Chất lượng ảnh · 5 Toàn vẹn Excel nguồn / đầu ra · 6 Truy vết được · 7 Tự phát hiện, xử lý, báo blocker · 8 Giảm thao tác thủ công.
Không dừng toàn bộ chỉ vì một case / công đoạn bị chặn; làm tiếp phần độc lập an toàn.

## Trạng thái case
| Trạng thái | Nghĩa |
|---|---|
| PASS | Mọi bước bắt buộc đã chạy + có ảnh + assertion đúng |
| FAIL | Có bằng chứng ứng dụng không đáp ứng expected (chấm đúng chữ của testcase; ghi hành vi thực tế + nguyên nhân trong source nếu xác định được và `source.use = true`) |
| BLOCKED | Không thể thực hiện / quan sát dù có người làm — luôn kèm **điều kiện gỡ** |
| PENDING | Cần người (đăng nhập MFA / SSO, quyết định chưa chốt, xác nhận thủ công, EXTERNAL chưa cho phép, adapter vai trò chưa cấu hình) — kèm "NEEDS: …" |
| NOT EXECUTED | Chưa chạy được do script / runner hoặc bước `t.todo()` — việc của agent |
| NOT AUTOMATED | ID trong phạm vi nhưng chưa có suite cài đặt — không chạy, không báo PASS |

Thứ tự tổng hợp: FAIL > PENDING > BLOCKED > NOT EXECUTED > PASS.

## Phân loại nguyên nhân lỗi
`product` (ứng dụng sai) · `testcase lỗi thời` (app đổi đúng ý đồ, testcase còn nội dung cũ — vẫn FAIL theo testcase, báo để cập nhật testcase) · `automation` (selector, timing, assertion, runner) · `environment/permission` (đăng nhập, quyền, slot, adapter, policy) · `test data / precondition` · `needs human`. `qakit.tools.runs` chỉ đưa **gợi ý**; agent xác nhận bằng ảnh evidence (và source nếu cần, chỉ khi `source.use = true`) trước khi kết luận `product`.

## Sự thật của dự án (đọc nguồn, không đoán)
- `qa.config.json`: URL, đăng nhập, vai trò / môi trường (matrix + adapter), tiền tố dữ liệu test, nguồn testcase, ngôn ngữ Excel, `decisions`.
- `project/hooks.py`: cách đọc dữ liệu (snapshot), dọn dữ liệu test (CLEANERS), kiểm nguồn dữ liệu (check_data).
- `docs/system_map/*`: màn hình, quyền, API, dữ liệu. `lessons_learned.md` (nếu có): bẫy đã gặp của dự án.
- Bẫy thường gặp ở mọi dự án: DB / API mà công cụ đọc khác cái app dùng (luôn `check_data`) · quyền chỉ nạp lại khi đăng nhập lại · toast biến mất nhanh · phần tử ngoài viewport → không có khung đỏ · thao tác sinh dữ liệu là GHI · dialog chọn sẵn "tất cả" · UI đa ngôn ngữ đổi nhãn · dữ liệu legacy không tạo được bằng UI · không có log thì case BLOCKED từ đầu.

## Nơi lưu (truy vết)
```
<autotest>/qa.config.json                    cấu hình dự án
<autotest>/test_suite/                       catalog.json · history/<build>.json · CHANGELOG.md · sources.json · source_map.json
<autotest>/test_suite_<sandbox>/             catalog riêng (thử nghiệm / demo), cùng cấu trúc
<autotest>/test_execution_output/
  runs/index.json                            danh sách lần chạy
  runs/<YYYYMMDD_HHMMSS>_<project>_<nhãn>/   = QA_OUT của run
    SUMMARY.md · plan.json · excel/ (bàn giao, mỗi ngôn ngữ một file, + _tong_hop khi chạy một phần)
    evidence/base/<suite>/<ID>/ · evidence/<lang>/<suite>/<ID>/
    reports/ (report.md, compare.md, run_report.md, gate.md, run_state.md, metrics.md, impact.md)
    data/ (results_<suite>.json, consolidated/, preflight.json) · safety/ (created_registry.json, snapshots/) · logs/ · tmp/
  adhoc/<YYYYMMDD>/                          lệnh lẻ không đặt QA_OUT
  _system/                                   autorun.lock
```

## Ngôn ngữ
Trao đổi với user bằng `language.talk` của config. Evidence: một file Excel cho mỗi ngôn ngữ trong `language.excel` (report.md). Giữ nguyên nhãn UI / thông điệp của app trong 「」.
