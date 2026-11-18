# System map

Mô tả hệ thống cho agent đọc khi viết / sửa code test (`automate`, `source`): màn hình, route, quyền theo vai trò, API, dữ liệu. Mỗi file một mảng, ví dụ:
- `A_screens.md` — từng màn hình: route, thành phần, thao tác, thông báo lỗi, file source FE.
- `B_permissions.md` — ma trận vai trò × màn hình / thao tác.
- `C_api_data.md` — API, bảng / collection dữ liệu, trường quan trọng.

Viết theo mục có tiêu đề (`## A §1 Login`) để `qakit.tools.impact` trỏ được mục cần cập nhật khi source đổi. Agent tạo bản đầu ở bước `/autotest-init` (quét source) và cập nhật dần.
