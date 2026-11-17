# Tự sửa lỗi và phân loại: bước trong testcase là đường LÝ TƯỞNG, hệ thống thật có vướng mắc

Testcase mô tả điều lý tưởng ("lưu bản ghi", "kiểm tra dữ liệu trong danh sách"). Nó không nói ô nào bắt buộc, dữ liệu có sẵn chưa, đăng nhập còn hạn không. **Việc của bạn là tự đi từ đường lý tưởng tới trạng thái testcase cần, rồi mới kiểm tra.** Một trường hợp chỉ là FAIL khi bạn chứng minh được lỗi nằm ở code của ứng dụng.

## 0. Trước khi sửa: tra cứu những gì đã biết
Một lỗi trông mới thường đã gặp rồi. Trước khi đổi code hay thử cách khác, **tìm trong tri thức đã học và các reference** (mất vài giây, đỡ cả vòng thử sai):
1. `grep -i "screen:<màn hình>" knowledge/index.md`, rồi lọc `kind:error`, `kind:fix`, `kind:rule`, `kind:locator` theo từ khóa của lỗi. Dòng khớp thì áp dụng trước, đừng tự dò lại.
2. Reference của skill: file này, `rules.md` (luật, nơi lưu, phân loại nguyên nhân), `automate.md` (cách viết case, chuẩn evidence), `confirm.md` (quyết định đã chốt với user).
3. Tài liệu hệ thống: `docs/system_map/*`, `inputs/` (nếu có), page object `pages/<màn hình>.py` và các case đã chạy xanh ở `suites/` có cùng màn hình (cách thao tác đã đúng).
4. Ảnh `diagnose/*.png` của lần lỗi (mở ảnh, đọc thông báo trên màn hình).
Sửa xong, **ghi lại** điều học được vào `knowledge/index.md` để lần sau (của chính bạn hoặc phiên khác) tra cứu được.

## 0b. Chạy lại sau khi sửa: cùng run, ghi đè
Sửa case thuộc một run đang làm thì chạy lại **đúng case đó trong run đó**: `python -m qakit.tools.rerun <ID>` (kết quả mới thay kết quả cũ, ảnh lần trước bị xóa). Phần chưa tới: `python -m qakit.tools.rerun --pending`. **Không dùng `retry`** (nó tạo run mới: tốn thời gian chuẩn bị và làm bẩn lịch sử). Run mới chỉ khi user yêu cầu.

## 1. Cây quyết định khi một bước không như mong đợi
Đi từ trên xuống, dừng ở mục đầu tiên đúng. Mỗi mục có công cụ trong `t` (TC); không tự viết vòng thử lại mù.

| Hiện tượng | Nguyên nhân thường gặp | Việc phải làm (tự làm, không hỏi user) |
|---|---|---|
| Bấm Lưu mà không lưu | Thiếu ô bắt buộc / giá trị không hợp lệ / đang có hộp thoại | `t.submit("<nút Lưu>", "Click Lưu", ok=lambda: <bản ghi đã hiện>)`: kit chụp ảnh lúc bấm, đọc thông báo + ô sai, **tự điền ô bắt buộc còn trống**, bấm lại. Ô đã có giá trị mà vẫn sai (quá dài, trùng, sai định dạng) → tự sửa giá trị theo thông báo rồi gọi lại (`fix=` nhận `d["fields"]`, `d["errors"]`) |
| Testcase bảo xem dữ liệu trên màn hình mà màn hình **không có dữ liệu** | Chưa có bản ghi | `t.ensure(n, "<mô tả dữ liệu>", present=lambda: <đang có>, create=lambda: <tạo qua UI>)`. Tạo bằng tên có tiền tố `data.test_prefix`, đủ theo điều kiện testcase (vd "nhiều tag" → tạo bản ghi có nhiều tag; "đạt maxlength" → dữ liệu đúng độ dài). Dọn sau khi chạy. **Không được đặt BLOCKED vì thiếu dữ liệu khi chưa chạy `ensure`** |
| Thiếu **tài liệu đầu vào** để kiểm một bước (file BD, thiết kế màn hình) | Không có trong testcase / `inputs/` | Tìm trong mọi sheet của workbook; không thấy → **hỏi user ngay** (`t.need_input`), nhờ đính kèm file. Không ghi PENDING rồi báo cáo |
| Testcase cần một **điều kiện biên** mà dữ liệu hiện có chưa thể hiện (tag chưa tràn khỏi ô, danh sách chưa quá một trang, chưa đạt độ dài tối đa) | Dữ liệu quá ít | `t.reach(n, "<điều kiện>", condition=lambda: <đã đúng>, grow=lambda i: <thêm một mẩu dữ liệu>, rounds=N)`: thêm từng vòng tới khi điều kiện đúng. **`grow` được tạo cả dữ liệu tham chiếu (tag, danh mục) có tiền tố**: cần tag tràn thì tạo tag tên rất dài; "hệ thống chỉ có 4 tag" không phải lý do để dừng. **Không ghi PENDING kiểu "chưa tràn"** |
| Bước chỉ đọc (sắp xếp, phân trang, lọc) mà danh sách **trống** | Thiếu bản ghi | `t.ensure` trước bước đó (đủ số bản ghi để sắp xếp / phân trang có nghĩa), rồi kiểm tra. Không để bước ghi BLOCKED "values: []" |
| Không thấy phần tử / hover, click quá hạn | Selector sai, menu ẩn, trong iframe, cần cuộn, chưa tải xong | Dò lại DOM (`playwright-cli snapshot` / `qakit.tools.dom_probe`), thử role + tên, label, text; kiểm tra iframe; chờ hết loading; mở menu cha. Sửa `pages/*.py` rồi chạy lại **case đó** (tối đa 2 lần sửa) |
| Màn hình khác lạ (hộp thoại, trang lỗi 500, bị đăng xuất) | Môi trường | Đóng hộp thoại hợp lệ; 5xx / mạng → thử lại 1 lần sau vài giây; hết phiên đăng nhập → PENDING, báo user bấm "Sign in again" (không tự đổi chế độ trình duyệt) |
| Dữ liệu tiền điều kiện cần quyền / thao tác ngoài tiền tố / thao tác không hoàn nguyên được | Ngoài phạm vi an toàn | PENDING hoặc hỏi một câu theo confirm.md |
| Mọi thứ hợp lệ, dữ liệu có, không thông báo lỗi, nhưng kết quả sai so với kỳ vọng | **Lỗi ứng dụng** | Mục 3 (`t.defect`) |

## 2. Những điều phải làm trước khi từ bỏ một đường
- Nhìn **ảnh** chụp lúc lỗi (`diagnose/*.png`, hiện cả ở dashboard). Thông báo lỗi trên màn hình là thông tin quan trọng nhất; đọc nó, không đoán.
- Thử **ít nhất hai cách khác nhau** (giá trị khác, đường vào khác, tạo dữ liệu khác) trước khi kết luận không làm được. Ghi từng cách vào `status fix "…"`.
- Sinh các **trường hợp biên** khi cần tới được trạng thái: dữ liệu ở độ dài tối đa, giá trị duy nhất (thêm số thứ tự / giờ vào tên), phụ thuộc giữa các ô (chọn ô cha trước), bản ghi ở trạng thái cần thiết (nháp / đã duyệt).
- Mỗi lần sửa thành công → ghi `knowledge/index.md` (`kind:fix`, `kind:rule`, `kind:data`) để lần sau không vấp lại.

## 3. Khi nào mới được ghi FAIL: `t.defect(...)`
```python
t.defect(n, "Tổng tiền = 40", "Bấm Tính", expected="40", actual="45",
         ruled_out=["đầu vào hợp lệ theo quy tắc (đọc lại từ màn hình)", "dữ liệu đã có (3 dòng)", "không có thông báo lỗi trên màn hình", "tải lại trang kết quả vẫn vậy"])
```
`ruled_out` bắt buộc, liệt kê điều bạn đã kiểm để chứng minh lỗi **không** do đầu vào, dữ liệu, selector hay môi trường. Kit tự chụp màn hình và đính kèm thông báo trên màn hình. **Không dùng `t.step(ok=False)` để báo lỗi ứng dụng khi chưa loại trừ được các nguyên nhân phía test.**

## 4. Bốn trạng thái, không lẫn
| Trạng thái | Khi nào |
|---|---|
| PASS | Đã làm đúng và kết quả khớp kỳ vọng |
| **FAIL** | Chỉ `t.defect`: ứng dụng sai, đã loại trừ lỗi phía test |
| BLOCKED | Đã thử `ensure` / `submit` / sửa selector mà vẫn không làm được **vì điều nằm ngoài tầm** (dữ liệu không thể tạo qua UI, không quan sát được giá trị phía sau). Ghi **điều kiện gỡ chặn** và **đã thử những cách nào** |
| PENDING | Cần người: đăng nhập hết hạn, MFA, quyết định, file mẫu, tài liệu đầu vào còn thiếu. **Không dùng cho "dữ liệu chưa đủ"** (đó là việc của `t.ensure` / `t.reach`) |
Một case thất bại vì chính test (ô trống, thiếu dữ liệu, selector sai) mà chưa sửa là **NOT EXECUTED** kèm lý do, **không bao giờ** FAIL.

## 5. Lỗi nằm trong kit (traceback trỏ vào `qakit/...`)
Đó không phải lỗi của case. **Đừng hỏi xin phép sửa**: nếu sửa đúng và nhỏ (vài dòng, hiển nhiên đúng) thì sửa tối thiểu, ghi `knowledge/index.md` (`kind:fix`, "kit patch: <file>: <việc>"), báo user một câu, chạy lại. Nếu không chắc: `status blocked "Lỗi kit: <file:dòng> <thông điệp>"`, dừng run (đừng để mọi case thành NOT EXECUTED hàng loạt) và báo user. File `qakit/` bị làm mới khi user bấm Save ở màn hình Testing nên bản vá tay sẽ mất; vì vậy nói rõ trong báo cáo để lỗi được sửa ở gốc.
