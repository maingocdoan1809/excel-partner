# Prompt kiểm thử Gateway (qua Claude for Excel)

Dùng khi Claude for Excel (hoặc client khác dùng Messages API) đã đăng nhập vào gateway: URL `https://localhost:8766` và token từ `excel-partner info --copy`. Chọn model muốn thử trong menu model (ví dụ `Copilot — Default`, `Claude — Sonnet 5.5`, `Codex — GPT-6-Sol`). Mỗi model/provider chạy một lần, chat mới.

Prompt này **không giả định tên tool**. Client tự gửi tool của nó cho gateway, nên model sẽ tự liệt kê những gì nó có.

Trước khi chạy: nhấn **Ctrl+S** một lần. Cuối bài test, nếu có gì sai, đóng file và chọn **Don't Save**.

````text
Bạn đang được gọi qua một Gateway (Messages API cục bộ) tới một CLI AI. Hãy kiểm thử xem chuỗi gọi tool và hội thoại có hoạt động đúng không, trên workbook đang mở. Tôi KHÔNG có bản sao của file này.

# Quy tắc an toàn (bắt buộc)
1. KHÔNG sửa, xóa hay định dạng bất kỳ sheet có sẵn. Chỉ được ĐỌC chúng.
2. Mọi thao tác ghi chỉ trên sheet nháp "_GW_TEST" do bạn tạo (nếu trùng tên, dùng "_GW_TEST2").
3. Không xóa gì ngoài những thứ chính bạn tạo. Xóa sheet nháp để cuối và hỏi tôi trước.
4. Tool nào lỗi thì ghi lại và đi tiếp. Không bịa kết quả. Chỉ báo "đã làm" khi kết quả tool xác nhận. Nếu không chắc, ghi "KHÔNG XÁC MINH".
5. Nếu thiếu một loại tool, ghi SKIPPED kèm lý do, không dùng cách vòng.

# Bước 0. Bạn có những tool nào?
Liệt kê TÊN từng tool bạn thực sự có (không đoán), chia nhóm: đọc, ghi, định dạng, cấu trúc sheet, biểu đồ/pivot, chạy mã (Office.js), khác. Ghi ngắn tham số chính của 3 tool đọc/ghi quan trọng nhất. Nếu bạn không thấy tool nào cả, dừng và báo điều đó.
Cũng nói rõ: bạn nghĩ mình là model nào (nói theo chữ bạn tự nhận biết, không bịa), và nhớ giúp tôi mã sau để kiểm tra bộ nhớ hội thoại ở cuối: GW-7421-KIWI. Chỉ trả lời "đã nhớ" cho mã này, không lặp lại nó trong các bước giữa.

# Bước 1. Một vòng đọc (round-trip đơn)
Dùng tool đọc để cho tôi biết: sheet nào đang hoạt động, vùng đang chọn, và vùng dữ liệu của sheet có nhiều dữ liệu nhất. Với mỗi tool bạn gọi, nói rõ đã gửi tham số gì và nhận lại gì. Giá trị bạn nêu phải khớp kết quả tool trả về.

# Bước 2. Tool song song
Bằng MỘT lượt suy nghĩ (nếu có thể, gọi nhiều tool cùng lúc), đọc đồng thời ba vùng khác nhau của sheet dữ liệu (ví dụ ô đầu, một khối 5x5, và một cột). Cho tôi biết bạn đã gọi mấy tool trong cùng một lượt và mỗi kết quả có khớp đúng với vùng của nó không (không lẫn kết quả giữa các lời gọi).

# Bước 3. Ghi và đọc lại (round-trip có thay đổi)
Tạo sheet "_GW_TEST". Ghi một bảng nhỏ 5 dòng x 3 cột (có tiêu đề), cột thứ 3 là công thức nhân cột 1 và cột 2. Đọc lại để xác minh giá trị công thức. Nếu client trả kết quả tính công thức, đối chiếu nó với phép tính tay của bạn.

# Bước 4. Đường lỗi của tool
a. Cố tình đọc một sheet không tồn tại ("KhongTonTai_123"). Tool phải trả lỗi; bạn phải báo lỗi đó đúng như nhận được, không đoán.
b. Cố tình ghi đè một ô đã có dữ liệu trong "_GW_TEST" mà KHÔNG cho phép ghi đè (nếu tool có cờ allowOverwrite thì không bật nó). Nếu tool chặn, hãy hỏi tôi trước khi làm lại. Nếu tool không chặn và ghi thành công, ghi rõ điều đó là kết quả quan sát được.
c. Nếu có tool chạy mã (Office.js), chạy một đoạn gọi API không tồn tại và báo lỗi thật.

# Bước 5. Kết quả lớn và ảnh
a. Đọc một vùng dữ liệu lớn (tối đa vài trăm ô thật có sẵn). Tóm tắt thay vì in hết, và nói kết quả có bị cắt không (nếu tool có báo hasMore, truncated, nextRanges...).
b. Nếu có tool chụp/hiển thị ảnh vùng, dùng nó trên vùng dữ liệu và mô tả đúng điều bạn thấy trong ảnh. Nếu không có, SKIPPED.

# Bước 6. Hội thoại nhiều lượt
Không gọi tool nào. Trả lời: (1) tên sheet nháp bạn tạo ở Bước 3 và giá trị công thức ở dòng 2 của nó, (2) mã tôi nhờ nhớ ở Bước 0. Cả hai phải đúng, không bịa.

# Bước 7. Văn bản dài và tiếng Việt
Viết một đoạn 300 từ tiếng Việt có dấu về lợi ích của việc kiểm thử công thức Excel. Không gọi tool.

# Bước 8. Báo cáo
Chỉ viết báo cáo MỘT lần, ở cuối cùng. Các bước 0 đến 7 trước đó chỉ trả lời ngắn (tối đa vài dòng, không dùng tiêu đề lớn, không dựng bảng). Báo cáo cuối KHÔNG được lặp lại nguyên văn danh sách tool của Bước 0 hay các câu trả lời của Bước 6 và 7: chỉ dẫn chiếu "xem Bước N" hoặc tóm tắt một dòng.
Một bảng: Bước | Tool đã gọi (tên + số lần) | PASS / FAIL / SKIPPED / KHÔNG XÁC MINH | Bằng chứng ngắn.
Sau bảng nêu riêng: (1) lỗi tool kèm nội dung lỗi, (2) chỗ bạn phải gọi lại tool nhiều lần, (3) chỗ mô tả tool khó hiểu, (4) chỗ kết quả tool khác mong đợi.
Cuối cùng hỏi tôi có xóa sheet nháp không. Chỉ xóa khi tôi xác nhận; sau khi xóa, đọc lại danh sách sheet để chứng minh workbook về như ban đầu.

Lưu ý: nếu hệ thống gửi cho bạn một <system-reminder> về tool lỗi, đó là do các ca lỗi cố ý ở Bước 4. Chỉ trả lời MỘT câu ngắn xác nhận các lỗi đã được báo, không viết lại báo cáo.

Bắt đầu từ Bước 0.
````

## Việc bạn tự làm (model không làm được)

Các ca sau liên quan tới client và gateway nên phải bấm tay. Ghi ✅ hoặc ❌ cho từng ca.

| # | Việc làm | Mong đợi |
| --- | --- | --- |
| M1 | Trong lúc model đang chạy một bước dài, bấm **Stop** | Dừng nhanh. Tin nhắn kế tiếp vẫn chạy bình thường |
| M2 | Đổi model giữa chừng trong cùng một chat (ví dụ Copilot → Claude) rồi hỏi "mã tôi nhờ nhớ là gì" | Hoặc trả lời đúng, hoặc báo rõ không còn bối cảnh. Không bịa |
| M3 | Chạy `excel-partner restart gateway` giữa lúc chat đang rảnh, rồi nhắn tiếp | Tin nhắn mới chạy được (có thể mất ngữ cảnh native nhưng không lỗi trắng) |
| M4 | Chạy `excel-partner restart gateway` **khi model đang chờ tool** | Lỗi rõ ràng, không replay lệnh ghi. Ghi lại thông báo |
| M5 | Mở hai chat song song trong Claude for Excel, hỏi mã ở chat kia | Không lẫn bối cảnh giữa hai chat |
| M6 | Chạy 8 chat khác nhau liên tiếp (giới hạn `maxSessions`) rồi mở chat thứ 9 | Báo hết chỗ rõ ràng, không treo |
| M7 | Đóng Excel khi model đang chạy | Gateway không treo; `excel-partner status` vẫn khỏe |
| M8 | `excel-partner doctor gateway` sau bài test | Tất cả `[ ok ]` |

## Nếu lỗi

| Triệu chứng | Kiểm tra |
| --- | --- |
| Client báo không đăng nhập được | `excel-partner info`, URL phải là `https://localhost:8766`. Chạy `excel-partner gateway warmup` rồi thử lại |
| Model không thấy tool | Chat mới. Gateway chỉ chuyển tiếp tool mà client gửi. Nếu client không gửi tool thì không có gì để gọi |
| Báo hết chỗ (capacity) | `excel-partner status` xem số phiên; chờ 30 phút (idle timeout) hoặc `excel-partner restart gateway` |
| Lỗi `502 api_error` | Đọc thông báo: thường là hạn mức hoặc đăng nhập của provider (ví dụ OpenCode hết hạn mức miễn phí) |
| Lỗi khác | `%LOCALAPPDATA%\Agent Gateway\gateway.err.log` và `gateway.log` (không chứa nội dung prompt) |

## Mẫu báo lỗi gửi cho dev

```
Bước:        (ví dụ 2 hoặc M4)
Model chọn:  (tên hiển thị trong menu)
Client:      Claude for Excel / khác, phiên bản
Hiện tượng:  (copy thông báo lỗi nguyên văn)
Workbook:    đã đổi gì? (có / không / một phần)
Log:         vài dòng cuối gateway.err.log
```
