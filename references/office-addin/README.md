# Dữ liệu tham khảo về add-in Office

Thư mục này chứa hai bản JavaScript được giải nén từ cache của một add-in Office, định dạng lại và thay các mẫu nhận diện bằng thông tin chung. Đây là bản trích để đọc tham khảo, không phải bản build có thể chạy.

- `main-bundle.reference.txt`: bundle chính, có schema công cụ và logic tích hợp Office.
- `support-bundle.reference.txt`: bundle phụ trợ.
- `sanitization-report.json`: số lượng thay thế, phạm vi rà soát và mã kiểm tra của các bản đã xử lý.

Các bước xử lý thay URL và email bằng địa chỉ ví dụ, thay đường dẫn người dùng, một số mẫu thông tin xác thực, mã UUID, mã build và tên sản phẩm bằng giá trị chung. Việc thay thế có thể làm thay đổi cú pháp và ý nghĩa của code; chỉ sử dụng các file này để đọc tham khảo. Không có cookie, Local Storage, cơ sở dữ liệu phiên hoặc nội dung tài liệu nào được sao chép.

Rà soát theo mẫu không bảo đảm nhận diện được mọi giá trị nhạy cảm; cần kiểm tra thủ công trước khi chia sẻ ra ngoài. Dữ liệu trong thư mục này không được đưa vào các đường dẫn `files` của gói npm hiện tại.
