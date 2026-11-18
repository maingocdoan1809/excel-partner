---
name: autotest-init
description: Khởi tạo bộ autotest (framework qakit + skill /system-test) cho MỘT dự án web bất kỳ - quét dự án (công nghệ, route / màn hình, cách đăng nhập, vai trò, nguồn dữ liệu, file testcase), hỏi user một lần, sinh qa.config.json + project/hooks.py + source_map.json + system_map bản đầu, cài thư viện, chạy check và nạp Test Suite, rồi tự động hóa thử một case để chứng minh chạy được. Dùng khi user chạy /autotest-init, nói "cài autotest cho dự án này", "áp dụng system test cho project mới", "khởi tạo autotest". Args (tùy chọn) - URL môi trường test, thư mục testcase.
---

# Khởi tạo autotest cho dự án

Nói chuyện với user bằng tiếng Việt (hoặc ngôn ngữ user đang dùng). Kết quả cuối: `/system-test check` OK, catalog đã nạp, ít nhất một case chạy được thật với evidence. **Không ghi dữ liệu lên môi trường của dự án ở đây** ngoài lần chạy thử cuối (chỉ dữ liệu test có tiền tố, sau khi user đồng ý).

## Luật
- Không sửa source của app. Chỉ ghi vào thư mục autotest (mặc định `<repo>/autotest`) và `.claude/skills/`.
- Không hỏi / không ghi mật khẩu vào file hay chat: config chỉ chứa **tên** biến môi trường; user tự đặt giá trị.
- Chế độ black-box (bước 0.5) → từ lúc quyết định trở đi **không mở file nào của app**, kể cả khi code nằm trong thư mục dự án.
- Không đoán điều chưa thấy trong source / tài liệu: ghi "chưa xác định" và đưa vào câu hỏi.
- Hỏi user **một phiên**, mỗi chủ đề một câu (tối đa 4 câu / lần gọi), theo cách viết của `system-test` references/confirm.md: nói bằng lời điều agent đã thấy trong dự án (màn hình nào, đăng nhập kiểu gì, vai trò nào) kèm bằng chứng ngắn, đề xuất mặc định ở phương án đầu; phần còn lại tự quyết theo mặc định an toàn và ghi rõ trong báo cáo.

## Bước 0 — Bộ kit đã có trong dự án chưa
- Có `<repo>/autotest/qakit/` và `.claude/skills/system-test/` → tiếp bước 1.
- Chưa có: cần đường dẫn bộ kit (thư mục chứa `install.py`; hỏi user nếu không biết) → `python <kit>/install.py <repo> [--dir autotest]` (không ghi đè file có sẵn). Đã có `qa.config.json` → đây là chạy lại init: chỉ bổ sung phần thiếu, không ghi đè cấu hình user đã chỉnh (hiện diff và hỏi trước khi đổi).

## Bước 0.5 — Đọc file đầu vào (nếu có)
- `<autotest>/inputs/autotest_input.xlsx` (mẫu ở `templates/` của bộ kit) đã điền → đọc trước mọi bước quét: cột "Khóa cấu hình" chỉ khóa `qa.config.json` tương ứng; ô còn là ví dụ (nền xám, chưa sửa) hoặc "chưa rõ" = chưa có thông tin, đưa vào câu hỏi bước 2. Thông tin trong file thắng thông tin suy ra từ source; mâu thuẫn → hỏi.
- Quyết định chế độ **trước khi đọc bất kỳ file nào ngoài thư mục autotest**:
  - sheet 1 "Có giao mã nguồn không" = Không, hoặc user nói không dùng source → **black-box**;
  - sheet 1 = Có → đọc source (bước 1);
  - không có file đầu vào / ô trống: thư mục dự án không có code app → black-box; có code → hỏi user một câu ("Dự án có code ở <thư mục>; agent có được đọc source không?"), chưa trả lời thì chưa quét.
- **Chế độ black-box** → `qa.config.json` ghi `"source": {"use": false}` (mọi lệnh `/system-test` sau đó cũng không đọc source, luật G12):
  - bỏ các mục quét source ở bước 1; màn hình / phân quyền / thông báo lấy từ sheet 4 / 5 / 7; thiếu sheet 4 → sau khi đăng nhập, dò menu và từng màn hình bằng `python -m qakit.tools.dom_probe` (chỉ đọc) để dựng `docs/system_map/A_screens.md`;
  - `source_map.json` để `repos` rỗng; báo user lệnh `source` / `suite impact` không dùng được;
  - sheet 6 không có đường đọc dữ liệu → `data.snapshot = none`; báo cáo init và mọi báo cáo run ghi rõ "dữ liệu chỉ kiểm qua UI, không kiểm toàn vẹn";
  - selector lấy từ DOM thật (`dom_probe`), kỳ vọng thông báo lấy nguyên văn từ testcase / sheet 7.

## Bước 1 — Quét dự án (chỉ đọc)
Ghi kết quả vào `<autotest>/docs/init_scan.md` (bảng: mục · phát hiện · bằng chứng file:dòng hoặc sheet:dòng của file đầu vào · độ chắc chắn).
1. **Công nghệ FE / BE:** `package.json` (next, react, vue, angular, svelte…), `*.csproj` (ASP.NET MVC / Blazor), `pom.xml` / `build.gradle` (Spring), `requirements.txt` / `pyproject.toml` (Django, Flask, FastAPI), `composer.json` (Laravel), `Gemfile` (Rails). Repo FE / BE tách nhau → ghi cả hai đường dẫn.
2. **Màn hình / route:** router (Next `app/**/page.tsx`, React Router, Vue router, Angular routes, controller + view, urls.py…) → danh sách route + file. Đây là khung cho `docs/system_map/A_screens.md`.
3. **Đăng nhập:** trang login + form (selector user / password / submit), SSO (NextAuth / Azure AD / OAuth / SAML → thường `manual` hoặc `cdp`), basic auth, không đăng nhập. Cách đăng xuất (URL / nút). Dấu hiệu "đã đăng nhập" (URL không chứa /login, hoặc phần tử header chỉ có khi đăng nhập).
4. **Vai trò:** enum / bảng role / middleware / decorator / policy → danh sách vai trò + màn hình mỗi vai trò thấy (khung cho `B_permissions.md`). Cách đổi vai trò khi test: nhiều tài khoản test (`accounts`) hay script / API quản trị của dự án (`command`).
5. **Môi trường / tenant** (nếu testcase phân biệt): khác nhau thế nào, đổi bằng gì.
6. **Dữ liệu:** DB (Postgres, MySQL, SQL Server, Mongo, Cosmos, Firestore…), ORM / model chính; API JSON đọc được; biến kết nối (`.env.example`). Chọn cách snapshot đọc-only: API công khai > truy vấn DB chỉ đọc > none.
7. **Testcase:** tìm `*.xlsx` (thư mục testcase / docs / QA). Đọc header bằng `python -c "from qakit.reporting import spec; print(spec.read(r'<file>')[:2])"` (chạy trong thư mục autotest) — lỗi "no testcase table" → ghi tên cột thật để đưa vào `testcases.columns`. Đọc cột vai trò / môi trường → từ khóa cho `matrix`. Ngôn ngữ testcase.
8. **Môi trường test:** URL (README, `.env.example`, pipeline deploy, user cung cấp). Không dùng production thật nếu chưa được user xác nhận.

## Bước 2 — Hỏi user một lượt
Những gì quét chưa chắc, **mỗi chủ đề một câu**, mỗi câu tự đứng được: "Agent thấy <điều gì, ở đâu trong dự án>; đề xuất <cấu hình bằng lời>; đồng ý?" — không hỏi bằng tên khóa config. Thường là:
1. **URL môi trường test + tài khoản test** (tên biến môi trường cho user / mật khẩu theo từng vai trò; với SSO: user sẽ đăng nhập một lần trên trình duyệt test).
2. **Vai trò cần kiểm và cách đổi** (nhiều tài khoản / script quản trị / chỉ một vai trò).
3. **Nguồn dữ liệu để kiểm toàn vẹn** (API / DB chỉ đọc / không kiểm) + tiền tố dữ liệu test (mặc định `ZZ_AUTOTEST_`).
4. **Ngôn ngữ evidence** (vi / ja / en, có thể nhiều) + thư mục file testcase.

## Bước 3 — Sinh cấu hình (từ `qa.config.example.json`)
1. `qa.config.json`: project, language, app.url / home_path / test_account, browser (SSO / MFA → `mode: cdp` hoặc `launch` + `auth.type: manual`; ngược lại `launch` + `form`), auth (selector thật từ bước 1), matrix (khóa vai trò / môi trường + từ khóa đúng như chữ trong cột testcase; `default_role` = vai trò của tài khoản mặc định), roles / envs adapter, data, testcases (sources, columns), `source.use` (theo bước 0.5). Xóa các khóa `_help` không cần.
2. `project/hooks.py`: viết thật `check_data()` (so một giá trị thấy trên UI với nguồn dữ liệu), `snapshot()` (chỉ đọc, các bảng / collection mà màn hình trong testcase có thể ghi), `discover()` + `CLEANERS` cho loại đối tượng test sẽ tạo (ưu tiên API công khai của app; nếu chỉ có UI thì cleaner mở trình duyệt qua page object). Mọi cleaner gọi `guard.check(guard.WRITE_TEST, name, …)` trước khi xóa. Chưa đủ thông tin → để mặc định trả "not implemented" và ghi vào báo cáo.
3. `test_suite/screens.json` (số / tên file testcase → tên màn hình dễ hiểu, dùng cho phiên xác nhận), `test_suite/sources.json` (= `testcases.sources`), `test_suite/source_map.json`: `repos` (đường dẫn repo FE / BE so với thư mục autotest) + quy tắc thư mục source → số file testcase từ bước 1 (direct cho màn hình riêng, cross theo tag cho phân quyền, broad cho layout / thư viện chung).
4. `docs/system_map/A_screens.md`, `B_permissions.md`, `C_api_data.md` bản đầu từ bước 1 (mỗi mục một tiêu đề `## A §n <màn hình>` kèm route + file source).
5. `.gitignore` của thư mục autotest đã có `browser_profile/`, `test_execution_output/` — giữ.

## Bước 4 — Cài đặt và kiểm tra
1. Tạo + kích hoạt venv riêng của `autotest/` trước khi cài gì (không cài vào Python hệ thống): `python -m venv .venv` rồi `.venv\Scripts\Activate.ps1` (Windows) hoặc `source .venv/bin/activate` (macOS/Linux). Sau đó `pip install -r requirements.txt` và `python -m playwright install chromium` (hoặc dùng `browser.channel: msedge|chrome` có sẵn trên máy). Lệnh cài cần quyền → xin user. Nhắc user mở Claude Code trong cùng terminal đã activate venv, nếu không các lệnh `python -m qakit...` sau này sẽ chạy nhầm Python hệ thống.
2. Nhắc user đặt biến môi trường mật khẩu (đưa đúng tên biến, cách đặt trên Windows / macOS / Linux). SSO / MFA: mở trình duyệt test một lần để user đăng nhập (launch không headless, hoặc cdp).
3. `/system-test check` (`python -m qakit.tools.preflight`): mọi mục OK hoặc WARN có lý do. Sửa config theo lỗi (selector đăng nhập, URL, header testcase…).
4. `/system-test update note="init"`: nạp catalog; `suite validate` — cột vai trò không nhận ra → bổ sung từ khóa `matrix`.

## Bước 5 — Chạy thử một case (chứng minh end-to-end)
1. Chọn một case **chỉ đọc** (mở màn hình, kiểm hiển thị) ở file testcase quan trọng nhất; hỏi user nếu không chắc.
2. `python -m qakit.tools.gen_suite --spec <file> --name <suite> --ids <ID>` → viết page object + case theo `system-test` references/automate.md.
3. Chạy thử `QA_OUT=<scratchpad>/init python -m suites.<suite>.cases --dev=<env>:<role> <ID>`; xem ảnh `_annotated`.
4. `python -m qakit.tools.catalog build --note "init: automate <ID>"` → `/system-test run id=<ID>` (qua cổng xác nhận) → mở `SUMMARY.md` và Excel.

## Bước 6 — Báo cáo cho user
`<autotest>/docs/init_report.md` + tóm tắt: đã phát hiện gì (bảng bước 1, mức chắc chắn), cấu hình đã chọn và vì sao, việc user còn phải làm (biến môi trường, đăng nhập SSO lần đầu, tài khoản vai trò, quyền DB chỉ đọc), phần chưa xác minh (toàn vẹn dữ liệu nếu snapshot = none, adapter vai trò chưa có), số case trong catalog / đã tự động hóa, lệnh tiếp theo (`/system-test automate file=Testcase_01_Login.xlsx`, `/system-test plan file=Testcase_01_Login.xlsx`).
