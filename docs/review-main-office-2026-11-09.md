# Rà soát main-office — 09/11/2026

Đã sửa các lỗi quyền được tái hiện, bổ sung bảo vệ vùng nhìn khi chạy mã Excel, và chỉnh Settings/Testing. Toàn bộ test hiện có chạy qua, nhưng chưa đủ bằng chứng để kết luận toàn bộ Office desktop đã được kiểm thử đầy đủ.

## Các lỗi đã xác nhận và sửa

| Mức độ | Vấn đề trước khi sửa | Thay đổi |
| --- | --- | --- |
| Cao | PowerPoint: model gửi `confirm_delete: true` có thể xóa slide trong chế độ Ask mà không hỏi người dùng. Auto không tự đặt cờ xác nhận. | Gỡ cờ do model tự xác nhận trong Ask; chỉ bật sau khi người dùng chấp thuận. Auto tự xác nhận. Test chạy qua executor PowerPoint giả lập và kiểm tra số slide thực tế. |
| Cao | Word/PowerPoint: nhiều thao tác ghi chạy thẳng trong Ask. Đã tái hiện thay toàn bộ câu trong Word với số lần hiện xác nhận bằng 0. `execute_office_js` cũng chưa được hỏi riêng. | Ask hỏi trước khi chạy các thao tác ghi và mã Office.js; từ chối thì executor không chạy. Thẻ `propose_doc_edits` tiếp tục lấy chấp thuận qua Apply. Auto vẫn tự thực hiện. |
| Trung bình | Mã trong `execute_office_js` có thể gọi `activate()`/`select()` và kéo người dùng về vùng AI đang sửa. Có thể liên tục khóa cập nhật màn hình. | Bọc các đối tượng truyền cho mã Excel để bỏ qua lệnh điều hướng và khóa màn hình, đồng thời trả thông tin các lệnh đã bỏ qua. Không sửa phương thức toàn cục nên người dùng vẫn điều hướng được khi mã đang chờ sync. |
| Trung bình | Nếu bỏ qua `activate()` rồi mã dùng active sheet/selected range để ghi, có nguy cơ ghi nhầm vào vùng người dùng đang xem. | Từ chối kiểu truy cập ngầm này sau một lệnh điều hướng bị bỏ qua, yêu cầu dùng worksheet/range tường minh. Test xác nhận cả hai sheet không bị ghi trong trường hợp đó. |
| Thấp | Đổi provider trong Settings đẩy người dùng về chat; Back từ màn hình con về chat; mất vị trí cuộn. | Giữ Settings khi đổi provider; Back về Settings và giữ bản nháp, vị trí cuộn cùng focus. |
| Thấp | Testing dùng input mặc định và nội dung kéo hết chiều rộng trình duyệt. | Form có nhãn, trường nhập/nút đồng bộ; Settings và màn hình con căn giữa, nội dung tối đa 672 px; hỗ trợ khung hẹp và hai theme. |

Bảo vệ vùng nhìn là cơ chế cải thiện hành vi của mã AI, không phải sandbox JavaScript. Không nên hiểu nó là cách cô lập mã không đáng tin. Chưa kiểm chứng cảm giác cuộn/chuyển sheet trên Excel Windows thật; test hiện dùng mô hình Office và các tình huống đổi sheet trong lúc chờ sync.

Microsoft cũng cảnh báo gọi `suspendScreenUpdatingUntilNextSync()` lặp lại có thể gây nhấp nháy: [Excel.Application](https://learn.microsoft.com/en-us/javascript/api/excel/excel.application?view=excel-js-preview).

## Kết quả kiểm tra

- `npm run typecheck`: pass, gồm server và add-in.
- `npm run build`: pass, gồm server, add-in và gateway.
- `npm test`: 129 file; 1.690 test, trong đó **1.688 pass, 2 skipped, 0 fail**.
- Hai test skipped: phân giải lệnh Copilot trên Windows PATH; chuyển provider qua nhiều model thật.
- Chromium chạy với bridge và bundle giao diện thật, provider giả lập để có kết quả ổn định: **9 nhóm kiểm tra pass**, không có uncaught page error hoặc request thất bại.
- Đã kiểm tra `/`, `/index.html`, chat, lưu instructions, đổi provider/theme, History sau reload, Back/bản nháp/vị trí cuộn; Settings và Testing ở 1200/380/320 px, sáng/tối.
- Không thực hiện sign-in lên hệ thống ngoài trong browser smoke test.

Chạy lại kiểm tra trình duyệt sau khi đã có Chromium tương ứng với Playwright:

```bash
npm run test:browser
# Nếu dùng Chromium đã cài ở đường dẫn khác:
BROWSER_EXECUTABLE=/path/to/chrome npm run test:browser
```

Script: `scripts/test-browser.mjs`. Settings của test nằm trong thư mục tạm và được dọn sau khi chạy. Kết quả và ảnh được ghi ở `output/playwright/office-review/` (được Git bỏ qua).

## Cảnh báo dependency còn tồn tại

`npm audit` báo 11 package entries: 2 critical, 4 high, 5 moderate. Khi bỏ dev dependencies: **5 entries production, gồm 3 high và 2 moderate; không có critical**.

| Dependency production | Mức audit | Ghi chú |
| --- | --- | --- |
| `zod` | High | Cảnh báo cạn bộ nhớ khi xác thực collection quá lớn. Audit đề nghị bản major mới; cần kiểm tra tương thích schema trước khi nâng. |
| `@github/copilot-sdk` | High | Entry kế thừa rủi ro từ dependency Zod. Không phải một exploit độc lập đã được tái hiện. |
| `body-parser` | High | Cảnh báo DoS khi xử lý compressed request body. Bridge chính dùng HTTP của Node; cần đánh giá đường đi thực tế của dependency trước khi quy thành khả năng khai thác bridge. |
| `fast-uri` | Moderate | Các cảnh báo chuẩn hóa URI/host/scheme. |
| `qs` | Moderate | Cảnh báo DoS trong `stringify`. |

Các entry critical thuộc công cụ test (`vitest`/`tinypool`); một số advisory cần điều kiện như Vitest UI server hoặc prototype pollution từ nguồn khác. Audit không phải bằng chứng rằng bridge hiện có remote code execution.

Chưa nâng dependency trong lượt này. Cần xử lý cả dependency trực tiếp và các bản lồng nhau, rồi chạy lại test thay vì dùng `npm audit fix --force` một cách tự động. Dữ liệu audit được giữ trong artifacts để đối chiếu advisory và phiên bản.

## Những phần chưa được bao phủ đầy đủ

1. Word, PowerPoint và Excel desktop Windows thật: Office.js runtime, API requirement sets, đồng bộ batch, thao tác người dùng đồng thời và khôi phục file.
2. Trọn ma trận provider/model thật. `providers.live.test.ts` là opt-in qua `LIVE_PROVIDERS`; số test pass mặc định không chứng minh các lượt gọi model thật đều hoạt động.
3. Không có đo coverage phần trăm trong cấu hình Vitest hiện tại. Nhiều bài tên `e2e` kiểm tra HTTP/WebSocket với Office/provider giả lập; browser smoke mới bổ sung kiểm tra trình duyệt thật.
4. Backup trước thao tác không thể Undo là best effort: khi tạo bản sao thất bại, journal vẫn cho thao tác tiếp tục và ghi `backupError`. Đây là hành vi đã được test rõ trong `tests/undo.test.ts`, chưa thay đổi trong lượt này. Với Word/PowerPoint cần lưu file và kiểm tra bản sao trước thay đổi quan trọng.

## Xem bản mới

Bản chạy từ WSL: `http://localhost:18766/index.html`. Đã xác nhận từ Windows rằng địa chỉ này phục vụ HTML mới. Cổng `8765` trên Windows là bridge khác và không tự nhận các sửa đổi trong workspace WSL.
