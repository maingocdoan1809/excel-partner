---
name: qa-testcases
description: "Sinh test case từ yêu cầu (requirement) trong workbook: đủ positive, negative, boundary, phân quyền; gắn Req_ID; kiểm tra độ phủ."
---
# Sinh test case từ yêu cầu

Trả lời bằng ngôn ngữ của người dùng. Làm việc trong workbook đang mở (mẫu: `TestPlan.xlsx` với các sheet Requirements, TestCases, RTM, Dashboard, Bugs).

## Trước khi viết
1. Gọi `qa_status` để biết sheet test case, tên cột và các role (profile) đang có. Dùng đúng tên cột đó.
2. Đọc sheet **Requirements** và các dòng đã có trong **TestCases** (`read_range`). Không tạo lại test case đã tồn tại; tiếp tục số thứ tự TC_ID từ số lớn nhất hiện có (TC_001, TC_002...).
3. Nếu yêu cầu mơ hồ, thiếu tiêu chí chấp nhận hoặc mâu thuẫn, **hỏi lại** (hoặc ghi vào cột Notes) thay vì tự đoán quy tắc nghiệp vụ.

## Mỗi requirement cần đủ các loại
- **Positive**: luồng đúng, dữ liệu hợp lệ.
- **Negative**: dữ liệu sai, thiếu trường bắt buộc, thao tác không được phép.
- **Boundary**: giá trị biên (0, 1, tối đa, tối đa + 1, chuỗi rỗng, ký tự đặc biệt, ngày biên).
- **Permission**: người dùng không đủ quyền không thấy hoặc không làm được.
- Thêm **Regression** khi requirement sửa chức năng cũ.
Không sinh test trùng ý nhau chỉ để cho đủ số lượng.

## Cách ghi từng dòng
| Cột | Nội dung |
| --- | --- |
| TC_ID | Duy nhất, dạng TC_001 |
| Req_ID | **Bắt buộc**, đúng một Req_ID có trong sheet Requirements |
| Module | Tên module/màn hình |
| Title | Một câu, bắt đầu bằng động từ, nói rõ điều kiện và kết quả |
| Preconditions | Dữ liệu và trạng thái cần có trước |
| Steps | Đánh số 1., 2., 3.; mỗi bước một hành động của người dùng |
| Expected | Kết quả quan sát được và **kiểm chứng được** (thông báo, giá trị trên màn hình, trạng thái bản ghi) |
| Type | Positive / Negative / Boundary / Permission / Regression / Usability |
| Priority | High / Medium / Low |
| Role | Tên profile đăng nhập (ví dụ admin) lấy từ `qa_status`; không bao giờ ghi mật khẩu |

Các cột **Automation_Status, Script_Path, Last_Run, Last_Result, Duration_s, Run_Evidence do công cụ điền**, không tự gõ vào.

## Ghi vào workbook
- Dùng `write_range`/`set_cells` vào các dòng trống ngay dưới dữ liệu hiện có, không ghi đè dòng cũ.
- Giữ định dạng và danh sách chọn của mẫu. Không gộp ô.
- Sau khi ghi, đọc lại vùng đã ghi và kiểm tra: mọi dòng có TC_ID và Req_ID; không có TC_ID trùng.

## Báo cáo kết quả
Nêu số test case đã thêm theo loại, các requirement chưa đủ thông tin để viết test (và câu hỏi cần trả lời), và nhắc người dùng bước tiếp theo: chọn các test case muốn tự động hóa rồi dùng `/qa-automate`. Không khẳng định "đã kiểm thử" vì mới chỉ là thiết kế test case.
