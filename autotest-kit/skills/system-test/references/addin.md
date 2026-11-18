# Chạy trong add-in Office Partner (đọc đầu tiên khi bạn đang được gọi từ Excel)

User chỉ mở file testcase trong Excel và nhắn yêu cầu. Add-in lo phần còn lại; bạn **không hỏi user đường dẫn file hay thư mục**.

## 1. Tìm thư mục làm việc và file testcase
- Thư mục autotest (có `qa.config.json`): đọc `autotest.json` trong thư mục dữ liệu của add-in (`%LOCALAPPDATA%\Excel Partner\autotest.json`, khóa `home`). Không có file đó → `%LOCALAPPDATA%\Excel Partner\autotest`. Mọi lệnh `python -m qakit...` chạy với thư mục này làm thư mục hiện tại.
- File testcase đang mở nằm ở `testcases/<tên>.xlsx`; tên đúng là giá trị `file` trong `testcases/ACTIVE.json`. Đây là **bản chụp** workbook user đang mở (kể cả phần chưa lưu), add-in tự cập nhật mỗi khi user nhắn về test. Khi user nói "file này", "sheet này", "các case đã chọn" là nói file ACTIVE. Cần chắc là bản mới nhất → `savedAt` trong ACTIVE.json phải mới hơn lúc user gửi yêu cầu; nếu cũ, nhờ user bấm **Send this workbook to the test kit** ở màn hình Testing.
- Kiểm tra môi trường chạy được: `python -m qakit.tools.preflight`. Python / thư viện thiếu → nhờ user mở **Settings > Testing > Test kit** bấm Install, không tự cài.

## 1a. Làm từng case cho chắc, sửa NGAY khi thấy lỗi (bắt buộc)
User không muốn mất thời gian vì AI làm sai cả suite rồi chạy lại. Nguyên tắc: **case nào xong phải chắc case đó**.
1. **Viết và chạy thử từng case một**, không viết cả suite rồi mới chạy: viết case → `QA_OUT=test_execution_output/adhoc/dev_<ID> python -m suites.<suite>.cases --dev=<env>:<role> <ID>` → đọc kết quả và ảnh (`diagnose/`, `evidence/`) → chưa PASS (và không phải `t.defect` có căn cứ) thì **sửa ngay** theo [repair.md](repair.md) rồi chạy lại **chính case đó** → xanh mới sang case kế tiếp.
2. Chạy chính thức (`autorun`) chỉ khi mọi case trong phạm vi đã xanh từng cái một. Chạy nền, rồi **đứng canh bằng `python -m qakit.tools.watch`** (không phải "đợi hoàn tất"). `watch` in từng sự kiện và **thoát ngay** khi có việc cần bạn: mã 10 = có case cần sửa (run tạm dừng) hoặc run tự dừng; 0 = xong; 11 = chưa có gì mới (gọi lại).
3. Khi `watch` thoát mã 10, làm theo thứ tự:
   - **Tra cứu trước khi sửa** (đừng đoán): `grep -i "screen:<màn hình>" knowledge/index.md` và `grep -i "kind:error\|kind:fix" knowledge/index.md | grep -i "<từ khóa lỗi>"`; đọc [repair.md](repair.md) và reference liên quan (`rules.md`, `automate.md`, `confirm.md` nếu đã có quyết định trước đó); xem `docs/system_map/*`, `inputs/`, và `pages/`, `suites/` hiện có của chính màn hình đó; mở ảnh `diagnose/*.png`. Cách sửa đã từng thành công thì dùng lại.
   - **Sửa nguyên nhân tại chỗ**, rồi **chạy lại ĐÚNG case đó trong CÙNG run**: `python -m qakit.tools.rerun <ID>`. Kết quả mới **ghi đè** kết quả cũ của case; không tạo run mới, không tạo thư mục mới, lịch sử các lần chạy không bị bẩn thêm.
   - Case xanh rồi thì chạy phần chưa tới trong cùng run: `python -m qakit.tools.rerun --pending`. Ghi điều đã học vào `knowledge/index.md` (kind:fix).
   - **Không dùng `retry` để sửa một run đang làm**: `retry` tạo run mới, tốn thời gian chuẩn bị và làm bẩn lịch sử. Chỉ tạo run mới khi user yêu cầu.
   - Không chạy lại cả suite.
4. Cấu hình add-in đã ghi sẵn: `run.stop_on_problem = true` (kit tự dừng tại case đầu tiên cần sửa, các case chưa tới giữ nguyên PLANNED), `run.case_retries = 1` (case vừa lỗi được chạy lại ngay một lần từ trang sạch trước khi coi là cần sửa), `run.max_attempts = 2` (số lần thử trong một thao tác, có đọc màn hình ở giữa).
5. **Thiếu điều kiện thì tạo, không đặt BLOCKED / PENDING:** danh sách trống → `t.ensure`; điều kiện biên chưa thể hiện (tag chưa tràn khỏi ô, danh sách chưa quá một trang, giá trị chưa đạt độ dài tối đa) → `t.reach`, thêm dữ liệu từng vòng tới khi điều kiện đúng.

## 1c. Ảnh bằng chứng phải có khung đỏ đúng chỗ (bắt buộc)
Mọi bước trong `t.step(...)` phải truyền `hl=[...]` trỏ **đúng phần tử** bước đang kiểm (ô, dòng, nút, vùng chứa tag), không phải cả bảng hay cả trang. Sau khi chạy thử mỗi case, **mở vài ảnh `*_annotated.png`** và xác nhận khung đỏ nằm đúng chỗ. Dòng `NO RED BOX` trong kết quả nghĩa là bước chưa có khung: sửa ngay.

## 1b. Báo tiến độ (BẮT BUỘC, làm trước khi làm việc đó, không làm sau)
User theo dõi bạn ở dashboard `https://localhost:8765/tests` và trong add-in. Họ chỉ thấy những gì bạn ghi ra. **Mỗi lần đổi giai đoạn, ghi một dòng trước khi bắt đầu giai đoạn đó:**

    python -m qakit.tools.status <giai đoạn> "<đang làm gì / tìm thấy gì, một câu>"

Giai đoạn: `setup` · `learn` (tìm hiểu hệ thống) · `automate` (viết code test) · `trial` (chạy thử) · `preflight` · `plan` · `run` · `fix` · `cleanup` · `report` · `done` · `blocked`. Ví dụ: `status automate "Viết code cho 12 case sheet Orders, đã xong 3"`, `status fix "T5 lỗi selector nút Lưu, sửa lần 1/2"`, `status blocked "Cần user đăng nhập MFA"`.

**Kế hoạch phải hiện ra TRƯỚC khi chạy** (từng case, các bước, kết quả mong đợi, dòng trong sheet):
1. `python -m qakit.tools.scope <phạm vi> --out <thư mục tạm>/plan.json` (như execute.md bước 0).
2. `python -m qakit.tools.runs new --plan <thư mục tạm>/plan.json --label <nhãn>` → tạo thư mục run và ghi sự kiện `run_planned`. Việc này nằm **trước cổng xác nhận** (execute.md bước 0 → 3), nên trong lúc user duyệt, dashboard đã hiện toàn bộ case của kế hoạch với trạng thái PLANNED. Ghi thêm `status plan "Chờ user xác nhận phạm vi N case"`.
3. Cổng xác nhận, rồi `qakit.tools.autorun --plan <run>/plan.json` (không tạo run thứ hai).

Chạy thử khi viết automation (`--dev`): đặt `QA_OUT=test_execution_output/adhoc/dev_<ID>` (không dùng thư mục tạm ngoài), để dashboard thấy bước và ảnh của lần thử.
Không có dòng `status` nào trong 10 phút khi bạn đang làm việc là lỗi của bạn: user sẽ thấy màn hình trống.

## 2. Kết quả vào đâu
- Kit ghi từng bước vào `test_execution_output/runs/<run>/events.jsonl` (hoặc `adhoc/<ngày>/`). Add-in đọc file này và **tự ghi trạng thái / kết quả / ngày vào đúng dòng của file testcase đang mở** (cột Result, Date, Note), đồng thời hiện tiến độ trong add-in và ở dashboard `/tests`.
- Bạn **không** ghi vào sheet bằng công cụ Excel cho việc này, không sửa `events.jsonl`, không sửa bản chụp `testcases/*.xlsx`.
- Excel evidence (ảnh khoanh đỏ) vẫn do `result view=excel` sinh ra như cũ.

## 3. Tự sửa lỗi và tự ngắt
- Kit tự dừng run khi **cùng một lỗi lặp lại** `run.max_repeat` lần (mặc định 3, cấu hình ở màn hình Testing). Khi dừng: các case chưa chạy ghi NOT EXECUTED kèm lý do, sự kiện `run_stopped`. **Không chạy lại ngay**: nêu nguyên nhân cho user và đề xuất cách xử lý.
- Ngân sách tự sửa: mỗi case tối đa **2 lần** sửa code test (selector, chờ, thứ tự thao tác) rồi chạy lại case đó bằng `--dev`. Lần thứ 3 vẫn lỗi → dừng, đánh dấu cần xem lại và nêu bằng chứng. Ghi từng lần sửa vào `knowledge/index.md` (kind:fix).
- Chỉ sửa được: code trong `suites/`, `pages/`, `project/hooks.py`. **Cấm** đổi kỳ vọng, chữ của testcase, hoặc làm yếu một assertion để ra PASS.
- Lỗi do ứng dụng (kết quả thật khác kỳ vọng) là FAIL hợp lệ: không "sửa" cho qua.
- Cùng một thao tác thất bại hai lần liên tiếp mà chưa có thông tin mới → đổi cách tiếp cận (dò lại DOM) hoặc dừng; không thử lại y nguyên.

## 4. Tri thức: lưu và tra cứu (bắt buộc, mỗi lần test)
`knowledge/index.md` là **một dòng một sự kiện**, để `grep` ra ngay:

    - [screen:orders] [kind:error] [tag:timeout] Nút Lưu chỉ bật sau khi chọn khách hàng -> pages/orders.py

`kind` ∈ login · navigation · locator · data · error · fix · quirk · rule. `screen` = tên màn hình theo `test_suite/screens.json`.

**Trước khi thao tác** (đừng đọc cả file):
- `grep -i "screen:<màn hình>" knowledge/index.md`
- gặp lỗi: `grep -i "kind:error\|kind:fix" knowledge/index.md | grep -i "<từ khóa lỗi>"`
- có dòng khớp → áp dụng trước khi tự dò lại.

**Sau mỗi lần chạy / sửa** thêm dòng vào CUỐI file (mới nhất ở cuối): điều mới học về màn hình, locator ổn định, quy tắc nhập, lỗi gặp + cách xử lý đã thành công, điều bất ngờ. Sửa tại chỗ dòng đã sai, không xóa dòng còn đúng, không thêm dòng trùng (grep trước). Không ghi mật khẩu, token, cookie, dữ liệu cá nhân. Chỉ ghi điều đã thấy tận mắt trong lần chạy này.

## 5. playwright-cli (dò màn hình, tìm selector)
Dùng cho việc **khám phá**: `playwright-cli open <url>`, `snapshot`, `find "<chữ>"`, `eval "el => el.id" <ref>` để lấy selector ổn định (role + tên, label, `data-testid`, id) trước khi viết / sửa `pages/*.py`. Đăng nhập: nếu `auth.type = manual`, dùng đúng trình duyệt user đã đăng nhập (`browser.mode`) thay vì tự đăng nhập; không bao giờ nhập mật khẩu.
- `playwright-cli` **không** tạo kết quả kiểm thử. Kết quả và evidence chỉ do suite Python của kit (`t.click`, `t.fill`, `t.step`) sinh ra.
- Đóng trình duyệt dò (`playwright-cli close`) trước khi chạy suite, tránh hai trình duyệt tranh một profile.
- Phát hiện gì hữu ích (selector, hành vi) → ghi ngay vào `knowledge/index.md` (kind:locator / quirk).

## 5b. Gặp lỗi trên màn hình (validation, không lưu được, không thấy nút) — quy trình bắt buộc
**Đọc [repair.md](repair.md) trước**: nó có cây quyết định đầy đủ và các công cụ `t.submit` (lưu có kiểm tra + tự điền ô bắt buộc), `t.ensure` (tạo dữ liệu còn thiếu thay vì BLOCKED), `t.defect` (cách duy nhất để ghi FAIL do lỗi ứng dụng). Phần dưới là cơ chế bên dưới chúng.

Trình duyệt chờ mặc định 8 giây (`browser.action_timeout_ms`), không phải 30 giây. Khi một thao tác lỗi, **đọc màn hình trước, rồi mới thử lại**; không bao giờ chạy lại y nguyên.
1. Trong code test, bọc thao tác dễ lỗi (lưu bản ghi, tạo dữ liệu tiền điều kiện) bằng `qakit.core.retry.attempt("tên việc", lambda: ..., fix=lambda d: ..., pg=t.pg)`. Lần lỗi đầu, kit tự **chụp màn hình + đọc thông báo lỗi, ô bị đánh dấu invalid, ô bắt buộc còn trống, hộp thoại đang mở** (`qakit.core.diagnose`), ghi vào `diagnose/<giờ>_<tên>.png` và hiện ở dashboard.
2. `fix(d)` nhận `d["fields"]` (danh sách ô, ví dụ `Customer Name (required and empty)`), `d["messages"]`. Sửa **nguyên nhân** (điền ô bắt buộc còn trống, chọn giá trị trong dropdown, đóng hộp thoại chắn) rồi kit thử lại một lần.
3. **MỞ ẢNH** `diagnose/*.png` bằng công cụ đọc ảnh khi `summary` nói "no validation message" hoặc khi chưa hiểu. Ghi điều thấy vào `status fix "…"`.
4. Cùng một lỗi màn hình lặp lại sau khi đã sửa (cùng `summary`) → kit ném `Diagnosed`, **dừng thao tác đó**, case ghi lý do thật ("Customer Name required and empty"), không ghi "timeout". Không tự thử lần thứ ba. Run vẫn tiếp tục các case khác; lỗi giống nhau nhiều lần thì bộ tự ngắt dừng run.
5. Điều học được (ô nào bắt buộc, cách chọn tag/dropdown của màn hình đó) ghi vào `knowledge/index.md` (kind:locator hoặc kind:rule) để lần sau không vấp lại: `grep -i "screen:<màn hình>" knowledge/index.md` TRƯỚC khi viết code cho màn hình đó.
6. Khoanh đỏ vào phần tử không hiện (bảng rỗng...) không làm lỗi bước nữa: kit khoanh khung chứa nó (hoặc cả màn hình) và ghi chú trong chú thích ảnh.

## 5c. Không hỏi user những việc này (user chỉ nhập thông tin ở màn hình Testing rồi mong test chạy)
- **Việc chỉ đọc hoặc chỉ ghi vào thư mục autotest** (dựng catalog, viết `cases.py` / `pages/`, dò màn hình, đọc config, chạy preflight, chạy thử `--dev` không ghi dữ liệu): làm luôn, không hỏi "Tiến hành?". Báo bằng `status`.
- **Tên cột testcase song ngữ**: kit tự nhận; chưa nhận được cột nào mới thêm vào `testcases.columns` rồi báo một dòng, không hỏi.
- **Đăng nhập**: dùng `auth.type` của config. `session` = phiên đã lưu ở màn hình Testing: không hỏi, không mở cửa sổ nhờ đăng nhập. Phiên hết hạn (`NeedsHuman`) → báo user bấm **Sign in again** ở Settings > Testing rồi dừng; **không tự đổi `browser.mode` sang cdp, không sửa `auth.type`, không nhờ gõ mật khẩu**.
- **Tạo / xóa dữ liệu test có tiền tố `data.test_prefix`**: nếu `run.ask_before_test_data` là `false` (mặc định của add-in) thì **bỏ chủ đề "Dữ liệu" khỏi cổng xác nhận**, vẫn dùng `guard` và dọn sau khi chạy. Nếu là `true` thì hỏi một câu như confirm.md.
- **Thiếu tài liệu / thông tin để kiểm một bước (file thiết kế BD, bố cục màn hình, giá trị mong đợi mà testcase không nêu): HỎI user NGAY**, một câu: bước nào, cần tài liệu gì, và nhờ họ đính kèm bằng `@` hoặc cho đường dẫn. Trước khi hỏi, tìm trong **mọi sheet** của `testcases/<file>.xlsx` và thư mục `inputs/`. Trong code test dùng `t.need_input(n, caption, action, "<cần gì>")` (ghi PENDING kèm điều cần và phát tín hiệu cho `watch`). **Không được kết thúc báo cáo với PENDING "thiếu tài liệu" mà chưa hỏi.** User trả lời "không có" thì mới để PENDING và nói rõ lý do trong báo cáo.
- **Tạo thêm dữ liệu tham chiếu để đạt điều kiện của case là ĐƯỢC PHÉP**: tag, danh mục, bản ghi phụ... tên bắt đầu bằng tiền tố `data.test_prefix`. "Không sửa dữ liệu có sẵn" nghĩa là không sửa hay xóa bản ghi có sẵn, **không** cấm tạo bản ghi mới. Cần tag tràn khỏi ô → tạo tag có tên rất dài (màn hình quản lý tag); cần danh sách dài hơn một trang → tạo thêm bản ghi. Câu "hệ thống chỉ có N tag nên không kiểm được" không phải lý do để dừng: tạo thêm. Chỉ khi màn hình tạo dữ liệu đó thật sự không dùng được qua UI (đã thử và có ảnh chẩn đoán) mới BLOCKED, kèm điều kiện gỡ chặn.
- Chỉ hỏi (ngắn, mỗi chủ đề một câu) khi: thao tác **ngoài** tiền tố hoặc không thể hoàn nguyên, hành động EXTERNAL (gửi mail, gọi hệ thống ngoài), kỳ vọng của testcase không rõ nghĩa, hoặc cần quyết định mà testcase / hệ thống không trả lời được.

## 6. Báo cáo cho user
Mở đầu bằng kết quả. **Mỗi PENDING / BLOCKED trong báo cáo phải đi kèm điều user cần làm hoặc câu hỏi bạn đã hỏi**; không để một dòng "cần thêm tag dài" hay "thiếu tài liệu" như một việc bạn không làm được: nếu làm được thì làm, nếu cần user thì hỏi.
Ngắn, bằng `language.talk`: đã chạy bao nhiêu case, PASS / FAIL / còn lại, run có bị tự dừng không và vì sao, đã sửa gì, đã học thêm gì (số dòng mới trong knowledge), và đường dẫn dashboard (`https://localhost:8765/tests`). Không dán lại cả bảng kết quả (add-in và sheet đã hiện).
