# Plan: skill `odoo-test`

Một skill độc lập để kiểm thử Odoo từ file testcase. Add-in Office Partner chỉ là một công cụ gắn vào (mở form, mở dashboard, nút Confirm). Skill chạy được hoàn toàn từ terminal.

## 1. Mục tiêu
- AI tự sửa lỗi có kỷ luật (tối đa 3 lần/case), agent loop gọn, không chia thành nhiều run rời rạc.
- Một workspace (cấu hình dự án) + một session cho mỗi yêu cầu test. Mọi case của session test trên cùng cấu hình đó.
- User chỉ chat: chỉ định file testcase (add-in: file đang mở hoặc file chỉ định; terminal: file chỉ định).
- AI tự phân tích case (rõ ràng, đủ dữ kiện), đề ra chiến lược, user xác nhận, rồi mới chạy.
- Evidence từng bước, có khoanh vùng, PASS hay FAIL đều chụp.
- Tri thức học được được lưu lại và tra cứu lại.
- Không chép file nào vào thư mục dự án của user. Mọi cấu hình, log, evidence nằm ở Home.

## 2. Kiến trúc

```
skills/odoo-test/
  SKILL.md              <=150 dòng: quy trình 6 pha + bảng lệnh
  references/           odoo.md (kiến thức Odoo) · repair.md · evidence.md
  oqa/                  CLI Python: python -m oqa <lệnh> (máy trạng thái, mọi quy trình nằm ở đây)
  oqa/core/             tái dùng từ qakit: evidence, ui (khoanh đỏ), guard, diagnose, retry, browser, tc
  oqa/web/              form init + dashboard (HTML tự chứa)
  hooks/                hook Claude Code (lớp phụ)

~/.odoo-test/                       (đổi bằng ODOO_TEST_HOME hoặc ô "Thư mục lưu" trên form)
  workspaces/<ws>/
    workspace.json                  url, db, profile, thư mục tham chiếu (không chứa mật khẩu)
    knowledge/index.md              tri thức, một dòng một sự kiện, grep được
    scripts/                        script case đã chạy được, tái dùng
    sessions/<id>/
      state.json  cases.json  analysis.json  strategy.md  events.jsonl
      evidence/<case>/<step>_{raw,annotated}.png   logs/
```

**Nguyên tắc cứng:** quy trình do CLI ép. Mỗi lệnh đọc `state.json` và từ chối khi pha trước chưa xong. Hook chỉ chặn đường vòng, không phải nơi giữ quy trình (provider khác Claude Code không có hook).

## 3. Quy trình 6 pha

| Pha | Lệnh | CLI kiểm tra |
|---|---|---|
| (trước) | `oqa init` | workspace phải tồn tại trước khi tạo session |
| INTAKE | `oqa session new --ws W --file F` | đọc xlsx → `cases.json`; dò tham chiếu file, ID bản ghi, dữ liệu giả định |
| ANALYZE | `oqa analyze --submit analysis.json` | validate schema. Mỗi case bắt buộc có: điểm rõ ràng, đủ dữ kiện, tham chiếu thiếu, dữ liệu cần + bước `ensure`/`reach` dự phòng, edge case, phụ thuộc case khác, danh sách step. Thiếu mục nào thì từ chối |
| CONNECT | `oqa connect` | URL, đăng nhập từng profile, lưu `storageState`, ghi trang đích sau login vào knowledge |
| CONFIRM | `oqa strategy` → user xác nhận | sắp xếp case theo đồ thị phụ thuộc, sinh `strategy.md`. `confirm` chỉ nhận từ nút trong add-in/dashboard hoặc TTY thật (Bash của AI không phải TTY) |
| EXECUTE | `oqa run <ID>` | từng case một, chỉ sau confirm; case trước để lại trạng thái cho case sau; đếm lần sửa, lần thứ 4 bị từ chối và buộc `oqa ask` |
| REPORT | `oqa report` | Excel evidence từng case + tổng hợp; `learn` kiểm tra trùng rồi mới ghi |

Lệnh phụ: `ask`/`answer` (session sang WAITING_USER), `learn`, `status`, `doctor [--fix]`, `ws show|set`, `dashboard`.

## 4. `oqa init`
1. Đọc Home; có `workspace.json` thì parse và điền sẵn form, chưa có thì tạo mới.
2. Bật server tạm `127.0.0.1:<cổng ngẫu nhiên>`, token dùng một lần, kiểm tra Origin, tự mở trình duyệt với form HTML.
3. Form chỉ có: tên workspace · URL (+ DB nếu cần) · profile (tên như Admin/Teamlead, tài khoản, mật khẩu, ghi chú vai trò) · thư mục tham chiếu (nhiều dòng, tùy chọn) · nâng cao thu gọn (ngôn ngữ, tiền tố dữ liệu test, thư mục Home). Nút "Kiểm tra kết nối" dùng chính code của `connect`.
4. User bấm Lưu → ghi file → tắt server → in tóm tắt không có mật khẩu → thoát mã 0. Quá 15 phút → mã 2, không ghi gì.
5. AI nhận input: `oqa init` tự chặn cho tới khi lưu. Claude Code chạy nó bằng background task và được báo khi thoát; provider khác chạy ở foreground với `--timeout`. `oqa init --status` chỉ báo đã cấu hình hay chưa.

**Mật khẩu:** không ghi vào `workspace.json`. Lưu trong Windows Credential Manager (`keyring`); form không hiện lại mật khẩu đã lưu (để trống = giữ nguyên); CLI không in mật khẩu ra log. Không có Credential Manager thì hỏi user có muốn dùng file mã hóa DPAPI, không ghi plaintext âm thầm. Hook chặn AI đọc `credentials*` và `storageState*`.

## 5. Evidence (CLI ép)
- Danh sách step của từng case có từ pha ANALYZE, nên dashboard hiện đủ step ở trạng thái PLANNED trước khi chạy.
- `t.step(...)` không ghi nhận nếu không có ảnh thật: ảnh gốc + ảnh khoanh đỏ đúng phần tử cần kiểm.
- Step lỗi (PASS hay FAIL đều chụp): CLI tự cuộn phần tử vào khung nhìn, khoanh đỏ phần tử lỗi (hoặc hộp thoại/thông báo lỗi), rồi mới ghi FAIL.
- Step thiếu `hl` hoặc khung đỏ không nằm trong ảnh → "evidence không hợp lệ", case không thể PASS.
- Sau mỗi case, `oqa run` kiểm lại: mỗi step có ≥1 ảnh tồn tại, không rỗng, có điểm ảnh đỏ. Thiếu thì hạ xuống NOT EXECUTED, lý do "thiếu evidence".
- Excel evidence do `oqa report` dựng từ chính `events.jsonl`, nên dashboard và Excel luôn khớp.

## 6. `oqa dashboard`
Cùng cơ chế server tạm với `init`.
- Trái: lịch sử (quét `workspaces/*/sessions/*`: workspace, file testcase, thời gian, pha, số PASS/FAIL), lọc theo workspace.
- Giữa: các case của session đã chọn, thanh tiến độ, chip trạng thái.
- Chi tiết case: từng step (mô tả, kỳ vọng, thực tế, trạng thái, ảnh); bấm để phóng to, chuyển ảnh gốc/khoanh đỏ; tab theo lần thử (1/3..3/3) với ảnh lỗi.
- Trên cùng: thanh 6 pha, dòng "AI đang làm gì", khung hỏi/đáp khi WAITING_USER, nút Confirm.
- Session đang chạy: tự cập nhật bằng cách đọc đuôi `events.jsonl`; session cũ chỉ đọc.
- Chỉ ghi qua hai endpoint (confirm, answer). Ảnh chỉ phục vụ từ thư mục evidence của session, chặn path traversal.

## 7. Hook (Claude Code)
Đã kiểm trên Claude Code 2.1.325 bằng một tiến trình `claude -p` thật: trong hook ở frontmatter của skill, `${CLAUDE_SKILL_DIR}` **khai triển thành rỗng**, còn `${CLAUDE_PLUGIN_ROOT}` là thư mục của skill; SKILL.md dùng biến thứ hai. Hook `guard` chặn lệnh `oqa confirm` trong một lượt chạy thật của skill. `oqa hooks` in snippet với đường dẫn tuyệt đối cho nơi không khai triển được.

- PreToolUse: chặn sửa trực tiếp `state.json`/`events.jsonl`, chạy script case ngoài `oqa run`, thao tác ghi khi chưa CONFIRM, đọc file credentials.
- Stop: từ chối dừng khi session đang EXECUTE còn case chưa xong mà không ở WAITING_USER.
- UserPromptSubmit: nhắc pha hiện tại (một dòng).

## 8. Gói kiến thức Odoo
`references/odoo.md` + helper `odoo_ui`: đăng nhập rồi redirect khác nhau theo version (`/web`, `/odoo`), many2one autocomplete, wizard/dialog, notebook tab, onchange, toast, chờ `o_web_client`. Selector dò bằng DOM thật (`dom_probe`/`playwright-cli`), không giả định.

## 9. Thay đổi trong add-in
- Màn /test chỉ mở `oqa dashboard`; Settings > Testing mở `oqa init`.
- Mỗi lượt chat truyền `OQA_WS`, `OQA_SESSION` và đường dẫn file đang mở; script từ chối ghi sang session khác.
- Add-in chỉ tìm skill đã cài và gọi `oqa doctor --fix` thay cho cài kit/pip.
- Làm trên nhánh `feat/odoo-test-skill` cắt từ `origin/main-office`, nên main-office chưa có code autotest: chỉ **thêm** phần gắn mỏng, không xóa gì. Nhánh `feat/autotest-kit-integration` giữ làm tham khảo (không mang `autotest-state.ts`, dashboard, runs, runtime, home sang).
- Cần kiểm chứng ở M4: cách bridge nạp skill cho provider (`%USERPROFILE%\.excel-partner\skills`).

## 10. Lấy gì từ kit cũ (`feat/autotest-kit-integration`)
Tái dùng: `core/evidence`, `core/ui` (khoanh đỏ), `core/guard`, `core/diagnose`, `core/retry`, `core/browser`, `core/tc`, `dom_probe`, `cleanup`, `excel_builder`. Không mang: `catalog`, `scope`, `cmd`, `runs`, `autorun`, `impact`, `watch`, `gen_suite`, `autotest-init`, 8/9 reference. `preflight` → `connect`; `rerun` → sửa-trong-session của `run`.

## 11. Các mốc

| Mốc | Nội dung | Nghiệm thu |
|---|---|---|
| **M0** ✅ | Chốt schema (`workspace.json`, `cases.json`, `analysis.json`, `events.jsonl` gồm step/ảnh/lần thử, `state.json`) trong `skills/odoo-test/references/contract.md`; sửa dòng "No Python" trong AGENTS.md | schema có test |
| **M1** ✅ | `oqa` lõi: `init` (server tạm, form, Credential Manager), workspace, session, máy trạng thái, `analyze`/`strategy`/`confirm`/`ask`/`answer`/`knowledge`/`doctor` | pytest: mọi lệnh sai pha bị từ chối; `init` lưu/đọc lại đúng, không lộ mật khẩu |
| **M2** ✅ (trên Odoo giả) | `connect` + `scaffold` + `run` + vòng sửa 3 lần + khoanh đỏ + kiểm evidence + `finish` + `report` + `probe` | có test với trình duyệt thật trên app giả dùng class của Odoo; **chưa chạy trên Odoo thật** |
| **M2b** ✅ | `oqa dashboard` (`view.py` là reducer duy nhất, dùng chung với report) | xem bằng trình duyệt: lịch sử, case, step, ảnh, lightbox; confirm/answer qua API có test |
| **M3** ✅ | SKILL.md (64 dòng), 2 reference (odoo, repair) + contract, hook `guard`/`stop`, `oqa hooks` | hook có test như Claude Code gọi; **chưa kiểm** `${CLAUDE_SKILL_DIR}` trong frontmatter, dự phòng bằng `oqa hooks` |
| **M4** ✅ (code + test) | Add-in: mục "Odoo Test" trong panel Testing mở form `oqa init` và dashboard qua `bridge/src/odoo-test.ts`; `${CLAUDE_SKILL_DIR}` được thay trong nội dung skill; `oqa install --to` chép skill vào thư mục skills | `typecheck`, `build`, vitest (2218 test) xanh; **chưa thử tay trong Excel** |
| **M5** | Chạy một file testcase thật (phân trang thiếu dữ liệu, tham chiếu thiếu) | case thiếu dữ liệu tự `ensure`; tham chiếu thiếu bị hỏi từ pha ANALYZE |

M1–M3 không cần add-in nên dùng và kiểm thử được ngay trên terminal; M4 chỉ gắn giao diện lên.

## 12. Rủi ro
- Confirm trên terminal cần user tự gõ `oqa confirm` ở cửa sổ khác. Chấp nhận, vì đó là mục đích.
- Hook chỉ có ở Claude Code; provider khác an toàn nhờ CLI nhưng không có lớp chặn thứ hai.
- Odoo khác version: helper phải dò theo DOM thật. Cần một Odoo test truy cập được cho M2 và M5.
- Quy tắc AGENTS.md "package không bundle skill": skill nằm trong repo này khi phát triển, tách repo riêng sau nếu cần.
