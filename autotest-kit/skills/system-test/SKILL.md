---
name: system-test
description: Kiểm thử hệ thống (System Test) bằng framework qakit theo bộ testcase Excel của dự án - kiểm tra môi trường, xem trước và chạy theo phạm vi (file / ID / suite / changed / failed, loại trừ, vai trò / môi trường), chạy lại case lỗi, kết quả + so sánh với lần trước + báo cáo + Excel evidence (mỗi thao tác một ảnh khoanh đỏ, nhiều ngôn ngữ, bản tổng hợp), cập nhật Test Suite khi có file testcase mới, dọn dữ liệu test, viết automation, kiểm thử lại theo thay đổi source. Dùng khi user chạy /system-test, nói "chạy system test", "chạy testcase file 03", "chạy lại case fail", "có testcase mới / cập nhật test suite", "so sánh kết quả lần trước", "PR này ảnh hưởng case nào". Chưa có qa.config.json → dùng /autotest-init trước. Args - lệnh + tham số có tên, ví dụ "run file=Testcase_03_Orders.xlsx exclude=USR-010".
---

# System Test (điểm vào duy nhất)

**Thư mục làm việc:** thư mục autotest của dự án = thư mục chứa `qa.config.json` (mặc định `<repo>/autotest`). Mọi lệnh `python -m qakit...` / `python -m suites...` chạy từ thư mục này. Chưa có `qa.config.json` → dừng, đề nghị `/autotest-init`.
Nói chuyện với user bằng `language.talk` của config (mặc định tiếng Việt).

**Được gọi từ add-in Office Partner (Excel):** đọc [references/addin.md](references/addin.md) trước tiên (thư mục làm việc, file testcase đang mở, tự sửa lỗi, tự ngắt, tri thức, playwright-cli). Khi một bước không như mong đợi (không lưu được, màn hình không có dữ liệu, thiếu phần tử): đọc [references/repair.md](references/repair.md) TRƯỚC khi ghi BLOCKED hoặc FAIL. Lệnh `/autotest-init` không cần: user cấu hình môi trường ở **Settings > Testing**, add-in sinh `qa.config.json` và chép kit.

**Luôn đọc trước:** [references/rules.md](references/rules.md) (luật G1–G12, trạng thái, phân loại nguyên nhân, nơi lưu) và `qa.config.json`. Sau đó chỉ đọc **một** reference của lệnh đang làm.

**Chế độ black-box:** `source.use = false` trong `qa.config.json` → không mở bất kỳ file source nào của ứng dụng, ở mọi lệnh (G12); `source` và `suite impact` không dùng được.

## Lệnh

Cú pháp cố định: `/system-test <lệnh> [tham số=giá trị …]`. Mọi tham số đều có tên; nhiều giá trị ngăn bằng dấu phẩy, không dấu cách; giá trị có dấu cách đặt trong ngoặc kép (`tag="Phân quyền"`). Thứ tự tham số tùy ý.

**Bước 0 của mọi lệnh (bắt buộc, trước khi đọc reference hay chạy công cụ khác):** `python -m qakit.tools.cmd <nguyên văn args của user>`.
- Thoát 2 → in nguyên phần lỗi + cú pháp đúng cho user, **dừng**. Không tự sửa, không đoán, không chạy bản "gần đúng" (giá trị trần như `03`, `-ORD-010`, `full`, `demo`; file viết tắt như `file=03`, `file=orders`; tham số lạ / lặp / sai lệnh; giá trị ngoài danh sách; thiếu phạm vi).
- Thoát 0 → dùng dòng `OK …` (lệnh chuẩn hóa, ghi vào báo cáo / nhãn run) và dòng `scope:` (token truyền cho `qakit.tools.scope`, không tự ghép token khác). `python -m qakit.tools.cmd --help [<lệnh>]` = cú pháp đầy đủ.

**7 lệnh chính**
| Lệnh | Tham số | Việc | Reference | Ghi dữ liệu? |
|---|---|---|---|---|
| *(trống)* | — | Tình trạng: catalog + lần chạy gần nhất + gợi ý bước tiếp (`qakit.tools.catalog show`, `qakit.tools.runs list`) | — | không |
| `check` | `url=` | Kiểm tra môi trường (config, URL, trình duyệt, đăng nhập, nguồn dữ liệu, adapter, catalog) — `qakit.tools.preflight` | [execute.md](references/execute.md) | không |
| `plan` | phạm vi | Xem trước kế hoạch, không chạy — `qakit.tools.scope` | execute.md | không |
| `run` | phạm vi, `url=` | Chạy thật: preflight → **cổng xác nhận** → chạy → dọn → diff → Excel (+ bản tổng hợp) | execute.md | **có** (dữ liệu test) |
| `retry` | `run=` `url=` `sandbox=` | Chạy lại FAIL + NOT EXECUTED của một lần chạy (mặc định mới nhất) = `run select=failed run=…` | execute.md | **có** |
| `result` | `run=` `view=summary\|full\|excel` `base=` | `summary` (mặc định): kết quả + nguyên nhân lỗi + **so với lần trước** (`base=` chọn mốc); `full` = thêm báo cáo A–F; `excel` = build lại Excel | [report.md](references/report.md) | file báo cáo |
| `update` | `note=` `sandbox=` | File testcase mới / đã sửa: nạp thử → trình bày thay đổi → user đồng ý → lưu → validate → gợi ý `plan select=changed` | [suite.md](references/suite.md) | chỉ catalog |
| `clean` | `url=` | Xem kế hoạch dọn dữ liệu test → user đồng ý → dọn — `qakit.tools.cleanup [--apply]` | execute.md | **có** |

**Nâng cao**
| Lệnh | Tham số | Việc | Reference | Ghi dữ liệu? |
|---|---|---|---|---|
| `automate` | phạm vi | Viết / sửa code test cho case chưa có automation hoặc cần review | [automate.md](references/automate.md) | code suite / page object |
| `source` | `ref=` (bắt buộc) `url=` `execute=yes\|no` | Source đổi: case bị ảnh hưởng → sửa code test gãy → cập nhật system_map → chạy đúng các case đó (`execute=no`: dừng sau phân tích + sửa) | [update-source.md](references/update-source.md) | code test, system_map; **có** khi chạy |
| `suite show` | `file=` `id=` `suite=` `sandbox=` | Xem catalog | suite.md | không |
| `suite diff` | `base=<build>` `sandbox=` | Chênh lệch giữa 2 build | suite.md | không |
| `suite validate` · `suite coverage` | `sandbox=` | Kiểm lỗi / độ phủ automation | suite.md | không |
| `suite impact` | `ref=` hoặc `desc=` | Case bị ảnh hưởng (chỉ phân tích) | suite.md | không |

## Phạm vi (cho `plan` / `run` / `automate`)
Cần **ít nhất một tham số chọn**; không có thì `qakit.tools.cmd` từ chối (không bao giờ mặc định chạy tất cả).
| Nhóm | Tham số | Giá trị |
|---|---|---|
| Chọn (hợp) | `file=` | tên file `.xlsx` **đầy đủ** thuộc nguồn testcase (`testcases.sources` / `sources.json`) hoặc đường dẫn tới file đó (tương đối từ thư mục autotest hoặc tuyệt đối). **Không nhận viết tắt** (số file, một phần tên) |
| | `id=` | `ORD-001,ORD-002` |
| | `suite=` | tên suite |
| | `select=` | `all` · `changed` · `failed` |
| | `status=` | `FAIL,PENDING,BLOCKED,NOT_EXECUTED,PASS` |
| | `run=` | run_id đầy đủ, chỉ đi cùng `select=failed` / `status=` (bỏ trống = lần mới nhất) |
| Lọc (thu hẹp) | `tag=` | phân loại / loại chứa chữ đó, hoặc tiền tố ID |
| | `env=` | khóa của `matrix.envs` trong config (không khai báo → không dùng được) |
| | `role=` | khóa của `matrix.roles` trong config |
| Loại trừ | `exclude=` | ID, tên file đầy đủ / đường dẫn, suite, tag, env, role (env / role = bỏ slot đó) |
| Catalog | `sandbox=` | tên thư mục `test_suite_<tên>/` có sẵn |

Ví dụ: `run file=Testcase_03_Orders.xlsx` · `run file=Testcase_03_Orders.xlsx role=user exclude=ORD-010` · `run id=ORD-001,ORD-002` · `plan select=changed` · `run select=failed run=<run_id>` · `result view=full`.

## Quy tắc định tuyến
0. Mọi lệnh qua bước 0 (`qakit.tools.cmd`) trước; lỗi cú pháp → báo + cú pháp đúng, dừng. User yêu cầu bằng lời (không gõ `/system-test …`) → viết lệnh chuẩn tương ứng, cho user thấy một dòng, rồi đưa lệnh đó qua bước 0.
1. Không rõ lệnh → tình trạng (lệnh trống) rồi hỏi user muốn làm gì (một câu). Không tự đoán thành `run`.
2. Lệnh chỉ đọc (trống, `check`, `plan`, `result`, `suite …`) chạy ngay, không hỏi.
3. Lệnh ghi dữ liệu trên môi trường (`run`, `retry`, `clean`, bước chạy của `source`) luôn qua **một phiên xác nhận** trước thao tác ghi đầu tiên; `update` xác nhận trước khi lưu catalog. Cách hỏi: [references/confirm.md](references/confirm.md) (mỗi chủ đề một câu, theo màn hình / thao tác, mã chỉ trong ngoặc).
4. Phụ thuộc: chưa có catalog → `update` trước. Kế hoạch báo ⛔ Catalog out of date → `update` trước `plan` / `run`. `result` / `retry` cần một lần chạy đã có kết quả.
5. Kết thúc mỗi lệnh: tóm tắt ngắn (đã làm gì, số liệu, đường dẫn file), nêu phần chưa xác minh và lệnh nên chạy tiếp.

## Tài liệu liên quan
`qa.config.json` · `project/hooks.py` · `pages/` · `suites/` · `docs/system_map/*` · `test_suite/` (catalog, lịch sử, source_map) · `test_execution_output/runs/` (lịch sử chạy). Hướng dẫn cho người dùng: `docs/02_usage_guide.md` của bộ kit (nếu đã chép vào dự án).
