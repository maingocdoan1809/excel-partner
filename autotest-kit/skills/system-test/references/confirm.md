# Chuẩn xác nhận với user (cổng của `run` / `retry` / `source`, xác nhận của `update` / `clean`)

Mục tiêu: người duyệt **thấy rõ và duyệt được từng phần** mà không cần nhớ mã ID, tên suite hay tên kỹ thuật. Một cổng = một phiên xác nhận trước thao tác ghi đầu tiên; phiên có thể gồm nhiều câu hỏi riêng, nhưng sau khi phiên kết thúc thì không hỏi lại (G6). Viết bằng `language.talk` của config.

## 1. Bản xác nhận (`<run>/reports/gate.md`), viết trước khi hỏi
1. Phần phạm vi: dán nguyên đầu ra `python -m qakit.tools.scope <phạm vi> --brief`: theo **màn hình** (tên trong `test_suite/screens.json`), mỗi case một dòng "kiểm tra gì · chạy với vai trò / môi trường nào · dữ liệu test sẽ tạo (`suites/<suite>/notes.json`) · mã"; phần "Không chạy" kèm lý do. Màn hình chưa có tên đẹp → bổ sung `screens.json`; case tạo dữ liệu mà chưa có ghi chú → bổ sung `notes.json` trước khi hỏi.
2. Agent bổ sung, mỗi mục một tiêu đề, **bỏ mục không có nội dung** (ghi "không có"):
   - **Dữ liệu sẽ tạo / sửa:** ở màn hình nào, loại đối tượng gì, bao nhiêu, tên bắt đầu bằng tiền tố test, khi nào xóa. Thao tác sinh dữ liệu phụ (gửi chat, AI generate, gửi mail) ghi riêng.
   - **Vai trò / môi trường sẽ đổi:** tài khoản test nào, tạm thành vai trò / môi trường gì (tên dễ hiểu), cho những case nào, khôi phục về gì khi xong.
   - **Dự kiến không kiểm được:** case (tên bằng lời + mã), vì sao, cần gì để gỡ.
   - **Việc cần người trong lúc chạy:** đăng nhập SSO / MFA, thao tác tay… (sẽ thành PENDING nếu không có).
   - **Cần bạn quyết định:** mỗi điểm một dòng, kèm phương án mặc định an toàn.
   - **Thời gian dự kiến.**
3. Trước khi hỏi: gửi một tin nhắn ngắn tóm tắt (số case, số màn hình, có ghi dữ liệu / đổi vai trò không) + đường dẫn `gate.md`.

## 2. Câu hỏi: mỗi chủ đề một câu
- Dùng công cụ hỏi user của agent (AskUserQuestion); **mỗi câu một chủ đề**, tối đa 4 câu / lần gọi; nhiều hơn → gọi tiếp (cùng phiên, chưa ghi gì). Chủ đề không có nội dung → không hỏi.
- Thứ tự cố định:
  1. **Phạm vi** — "Chạy N case trên các màn hình A, B (danh sách bên cạnh)?" Kèm bảng rút gọn (màn hình → các case bằng lời) trong phần xem trước của câu hỏi. Phương án: chạy đúng phạm vi (đề xuất) · bỏ bớt (ghi trong ô khác) · hủy.
  2. **Dữ liệu test** — "Ở <màn hình>, test sẽ tạo <số> <loại đối tượng> tên bắt đầu <tiền tố> và xóa ngay sau khi chạy. Đồng ý?" Phương án: đồng ý · không (các case cần tạo dữ liệu bị bỏ, liệt kê tên case).
  3. **Đổi vai trò / môi trường** — "Tài khoản test <tên> sẽ tạm thành <vai trò> để kiểm <màn hình / việc gì>, xong trả về <vai trò gốc>. Đồng ý?" Phương án: đồng ý · chỉ chạy với vai trò hiện tại (case khác → PENDING, nêu số case).
  4. **Mỗi quyết định nhóm C một câu**; **thao tác ra ngoài hệ thống (EXTERNAL) một câu riêng**, nói rõ đó là gì (tạo tài nguyên cloud, gửi mail / ticket thật, thanh toán…) và hệ quả.
- **Cách viết mỗi câu:**
  - Mở đầu bằng **ở đâu + làm gì** bằng lời của người dùng app (màn hình, nút, loại dữ liệu); mã ID chỉ trong ngoặc sau mô tả, không bao giờ là danh sách mã trần.
  - Không đưa tên kỹ thuật thô (tên suite, khóa slot `env:role`, tên group / quyền nội bộ, biến môi trường, đường dẫn file); cần thì dịch nghĩa, tên kỹ thuật đặt sau trong ngoặc.
  - Số cụ thể: bao nhiêu case, đối tượng, vai trò, khoảng bao lâu.
  - Câu hỏi tự đứng được (user có thể chỉ thấy hộp thoại): không viết "như trên".
  - Phương án đầu tiên = đề xuất an toàn, ghi "(Recommended)"; mỗi phương án nói **hệ quả**.
  - Header ngắn: `Phạm vi`, `Dữ liệu`, `Vai trò`, `Quyết định`, `Ra ngoài`.
- Sau khi trả lời: ghi từng câu + câu trả lời vào `gate.md`; trả lời "bỏ" → lập lại plan (`scope … -<mã>`), cập nhật `gate.md`, không hỏi lại. Quyết định lâu dài → `decisions` trong `qa.config.json`.

## 3. Mẫu (ví dụ minh họa)
> **Phạm vi** — Chạy 4 case trên Màn hình danh sách item: mở danh sách (với Admin và User), chỉ admin thấy nút New item (với Admin và User), tạo item mới, giới hạn tên 30 ký tự. Không chạy: xuất báo cáo PDF (chưa có code), xóa item (bạn đã loại). Chạy đúng phạm vi này?
> - Chạy đúng 4 case (Recommended) — khoảng 1 phút
> - Hủy

> **Dữ liệu** — Ở Màn hình danh sách item, test sẽ tạo 2 item thử tên bắt đầu ZZ_AUTOTEST_ (một cho "tạo item mới", một cho "giới hạn 30 ký tự") và xóa khi dọn ngay sau lần chạy. Không sửa item có sẵn. Đồng ý?
> - Đồng ý (Recommended)
> - Không — bỏ "tạo item mới" và "giới hạn tên 30 ký tự" khỏi lần chạy

> **Vai trò** — Để kiểm màn hình với người dùng thường, test sẽ đăng xuất và đăng nhập bằng tài khoản User (viewer), xong đăng nhập lại Admin. Đồng ý?
> - Đồng ý (Recommended)
> - Chỉ chạy với Admin — phần "User" của 2 case đầu thành PENDING

## 4. Xác nhận của `update` và `clean`
- **`update`:** trình bày thay đổi theo màn hình, mỗi case bằng tên + mã trong ngoặc: case mới; case đổi nội dung (**trước → sau** của đúng cột đã đổi, rút gọn); case bị bỏ khỏi file. Mỗi loại thay đổi mơ hồ một câu hỏi riêng (case bị bỏ mà đã có code test; case đổi nội dung mà đã có code test). Câu cuối: "Lưu bản cập nhật testcase của <màn hình> (… mới, … đổi, … bỏ)?" kèm bảng thay đổi.
- **`clean`:** liệt kê theo màn hình: loại đối tượng · tên · tạo ở lần chạy nào; câu hỏi "Xóa N đối tượng test này (tên đều bắt đầu <tiền tố>)? Không đụng dữ liệu khác." kèm danh sách; mục MANUAL ghi riêng (cần xóa tay ở đâu).
