# Kiểm thử lại theo thay đổi source (`source`)

**Cần `source.use = true`.** `false` (black-box) → không làm, báo user: lệnh này cần đọc diff của source; muốn dùng thì đổi `source.use` và khai `repos` trong `test_suite/source_map.json`.

Cú pháp: `/system-test source ref=<thay đổi> [url=<URL đã deploy>] [execute=yes|no]`
`ref=`: số PR, branch, commit, hoặc range; nhiều repo: `ref="<repo>:A...B <repo2>:C...D"`. `execute=no`: dừng sau bước 6 (chỉ phân tích + sửa script + system_map + kế hoạch chạy).

Công cụ: `python -m qakit.tools.impact` (chỉ đọc: git diff + catalog; bảng ánh xạ `test_suite/source_map.json` — `repos` = tên → đường dẫn repo). Phần chạy dùng lại quy trình `run` của [execute.md](execute.md) (một cổng xác nhận). Phần sửa script theo [automate.md](automate.md).

## Các bước

| # | Bước | Việc | Đầu ra |
|---|---|---|---|
| 1 | Xác định thay đổi | PR / branch: base = branch đích → range `origin/<đích>...<branch>`. Commit đơn (kể cả commit merge) dùng trực tiếp. Ref chưa có trên máy → `git fetch origin` trong repo đó (không checkout, không sửa working tree). Không xác định được → hỏi user 1 câu | range theo repo |
| 2 | Catalog mới nhất | `python -m qakit.tools.catalog build --dry-run`; workbook nguồn đã đổi → `update` (suite.md) trước | catalog hiện hành |
| 3 | Bản cắt thô | `python -m qakit.tools.impact --range <repo>=<range> … --out <scratchpad>/impact` | file → khu vực → ID ứng viên; file chưa ánh xạ; ảnh hưởng rộng; code test dùng chuỗi UI đã xóa; mục system_map cần sửa |
| 4 | Thu hẹp (agent) | Đọc diff từng file (`git diff <range> -- <file>`), quyết định từng ID: **trực tiếp** (case kiểm đúng thay đổi), **hồi quy** (dùng chung code đã sửa), **loại** (ghi lý do). File chưa ánh xạ: xếp vào khu vực, **thêm quy tắc vào `source_map.json`**. Ảnh hưởng rộng (layout chung, file ngôn ngữ, nâng cấp thư viện) → bộ hồi quy đại diện (hiển thị ban đầu + truy cập màn hình của các màn bị chạm), ghi rõ. Tài liệu thiết kế đổi → ghi "spec đổi: testcase có thể cần cập nhật" | bảng ID · mức · lý do · có automation? |
| 5 | Sửa code test (trước khi chạy) | Mục "code test dùng chuỗi UI đã xóa / đổi" và mọi thay đổi DOM trong diff (nhãn, placeholder, name / id, field thêm / bỏ, dialog mới): sửa page object → suite theo automate.md; mong đợi vẫn lấy từ testcase. Kiểm locator chỉ đọc trên bản đã deploy (`--dev`). ID trực tiếp chưa có automation → đưa vào cổng: automate ngay hay để NOT AUTOMATED | file test đã sửa + lý do |
| 6 | Cập nhật system_map | Chỉ các mục `impact` liệt kê và đúng phần đã đổi: đọc source mới, sửa nội dung, thêm dòng `> Cập nhật <ngày> theo <repo>@<commit>: <tóm tắt>` dưới tiêu đề mục | mục đã sửa |
| 7 | Kiểm bản deploy | Môi trường (`url=` hoặc config) phải chứa thay đổi: so commit / version endpoint / chuỗi mới trong bundle tĩnh của app. Chưa deploy → dừng, báo user, giữ kế hoạch | xác nhận phiên bản |
| 8 | Chạy | `run <ID,… sau bước 4>` theo execute.md (`runs new --label src-<commit ngắn>`), thêm vào nội dung cổng: tóm tắt thay đổi, số ID trực tiếp / hồi quy / loại, script đã sửa, ID NOT AUTOMATED, spec có thể đã đổi | run_id |
| 9 | Phân loại + sửa vòng lặp | `runs show <run>`; xác nhận từng case không PASS bằng ảnh. `automation` → sửa script → `retry` (**tối đa 2 vòng**, mỗi vòng là run mới). FAIL mà app đổi đúng ý đồ nhưng testcase còn cũ → **`testcase lỗi thời`**: liệt kê để cập nhật testcase. FAIL còn lại → `product` | kết quả đã phân loại |
| 10 | Báo cáo | `<run>/reports/impact.md`: range, file → khu vực, ID (mức, lý do, kết quả), ID đã loại + lý do, script đã sửa, quy tắc mới trong `source_map.json`, mục system_map đã cập nhật, testcase lỗi thời. Thêm `runs compare <run>` + báo cáo A–F | báo cáo |

## Quy tắc riêng
- Không sửa source app — chỉ đọc diff. Không checkout / reset branch của user; xem file ở commit khác bằng `git show <commit>:<path>`.
- Không sửa testcase; testcase lỗi thời chỉ được báo cáo.
- Mỗi ID chạy có lý do; ID loại khỏi bản cắt thô cũng có lý do.
- Phạm vi lớn (> 150 lượt case × slot): đề xuất bộ hồi quy đại diện ở cổng, chạy toàn bộ chỉ khi user chọn.
- Thay đổi chỉ ở tài liệu hoặc test của repo → không cần chạy; báo phần spec đổi.
