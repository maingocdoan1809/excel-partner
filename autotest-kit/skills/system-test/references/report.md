# Kết quả, so sánh, báo cáo (`result [<run_id>] [full|excel] [base=<run_id>]`)

`result` (`view=summary`, mặc định) = kết quả của run + so sánh với lần trước trong một câu trả lời. `view=full` = thêm Báo cáo A–F. `view=excel` = build lại Excel (mọi ngôn ngữ + bản tổng hợp).

Nguồn: `test_execution_output/runs/<run_id>/` (SUMMARY.md, plan.json, data/results_<suite>.json, logs/autorun_*.log, reports/, excel/, evidence/). Không sửa results của một run đã xong; chạy lại = run mới.

## Kết quả của một run
1. Đọc `<run>/SUMMARY.md` (thiếu / cũ → `python -m qakit.tools.runs summary <run>`), rồi `python -m qakit.tools.runs show <run>` → bảng ID · trạng thái · **gợi ý nguyên nhân** · remark; NOT AUTOMATED của plan.
2. Với mỗi case không PASS, agent xác nhận nguyên nhân (rules.md): mở ảnh `_annotated` của bước lỗi + log; FAIL chỉ kết luận `product` khi ảnh / dữ liệu đọc lại chứng minh ứng dụng sai so với testcase; selector / timing / assertion sai → `automation` (đề xuất `automate` + `retry`); đăng nhập / quyền / adapter / slot → `environment/permission`; thiếu tài khoản / dữ liệu → `test data`.
3. Đầu ra: số theo trạng thái, bảng lỗi theo nguyên nhân (đã xác nhận / chỉ là gợi ý), việc cần làm tiếp.

## So sánh với lần trước (`base=<run_id>` để chọn mốc)
`python -m qakit.tools.runs compare <run> [--base …]` → ghi `<run>/reports/compare.md`: **New FAIL** (trước PASS), **Fixed**, **Still FAIL**, đổi trạng thái khác, lần đầu chạy, không đổi. Mặc định = kết quả gần nhất trước đó của TỪNG ID ở mọi run. Agent bình luận mục mới FAIL / đổi trạng thái (do ứng dụng đổi, môi trường, hay sửa script) dựa trên ảnh của cả hai run.

## Excel evidence (giai đoạn 6 của `run`, hoặc `result view=excel`)
- autorun build một file cho mỗi ngôn ngữ trong `language.excel`: `excel/<spec>_evidence_<lang>.xlsx` (case đã chạy). Lần chạy một phần có thêm `excel/<spec>_evidence_<lang>_tong_hop.xlsx` (`python -m qakit.tools.runs consolidate <run>`): mọi case của spec hiện tại, kết quả của run này hoặc run `finished` gần nhất của cùng suite (dòng "Nguồn kết quả"; tô cam nếu đặc tả đã đổi sau run đó; chưa từng chạy = NOT EXECUTED).
- **Ngôn ngữ:** caption viết theo ngôn ngữ testcase (`language.testcase`). Ngôn ngữ Excel khác ngôn ngữ testcase → tạo `suites/<suite>/<lang>_text.py` (SPEC / CLS / TYP / COVER / SHEETS viết tay; CAP / ACT / NOTE / REM: dịch từ results — có thể viết `make_<lang>_runtime.py` sinh tự động và thoát 1 khi còn chuỗi chưa dịch). Build: `QA_OUT=<run> python -m qakit.reporting.excel_builder --suite <suite> --spec <spec> --only-run --lang <lang> --tr suites/<suite>/<lang>_text.py`. Giữ nguyên nhãn UI / thông điệp trong 「」; email, ID, số không dịch. Sau khi sửa bản dịch → chạy lại `consolidate`.
- Bố cục: sheet tổng hợp đầu tiên; giữ nguyên các sheet nguồn; mỗi testcase một sheet tên theo ID; mỗi sheet: kết quả (màu: PASS xanh / FAIL đỏ / BLOCKED vàng / PENDING cam / NOT EXECUTED xám), thời điểm, môi trường, URL, phân loại, vai trò, tiền điều kiện, bước, expected, từng bước (Status · caption · thao tác · thực tế), ảnh nhúng rộng 900px với nhãn `Step n-k: …`, thời gian theo slot.
- **Nghiệm thu (bắt buộc):** mở / render mọi bản (Excel thật nếu có: COM → PDF → ảnh; hoặc LibreOffice) và xem: sheet tổng hợp, ít nhất một sheet có ảnh, mọi sheet FAIL / BLOCKED / PENDING; khung đỏ đúng phần tử, chú thích đúng ngôn ngữ, không ảnh bị cắt / lệch bước; sheet nguồn giống file gốc. Không render được → ghi rõ phần chưa xác minh. Xóa file render tạm (danh sách cụ thể).

## Báo cáo A–F (`result view=full`; giai đoạn 7 của `run`) → `<run>/reports/report.md`
- **A. Kết quả:** số file / case trong phạm vi; PASS / FAIL / BLOCKED / PENDING / NOT EXECUTED / NOT AUTOMATED; case lỗi + nguyên nhân (đã xác nhận / gợi ý) + **điều kiện gỡ**; đối chiếu nhóm dự kiến ↔ thực tế.
- **B. File đầu ra:** Excel, ảnh, log, results; kiểm tra đầy đủ.
- **C. Chất lượng evidence:** kết quả nghiệm thu; lỗi đã tự sửa; tiêu chí chưa xác minh.
- **D. Khả năng tự động hóa:** blocker còn lại: mô tả, bằng chứng, case bị ảnh hưởng, nguyên nhân gốc, đã thử gì, hành động gỡ, quyền / người cần.
- **E. Toàn vẹn dữ liệu:** thay đổi đã làm (kể cả vai trò / môi trường) và khôi phục; dữ liệu test còn sót; thay đổi ngoài ý muốn; kết quả snapshot diff (hoặc "chưa kiểm" khi snapshot = none).
- **F. Thời gian:** theo giai đoạn và theo case × slot (`timings` trong results); tổng cả lần chạy.
- So sánh với lần trước: tóm tắt `compare.md`. Tóm tắt gửi user = A + E + mới FAIL / PENDING. Tách bạch: **đã xác minh · rủi ro · chưa thể kết luận**. PENDING (việc user phải làm) lên đầu.

### Checklist trước khi nói "hoàn thành"
- [ ] Mọi case trong plan có trạng thái + lý do; mọi bước đã chạy có ảnh thật, khung đỏ đúng phần tử; NOT AUTOMATED được liệt kê, không tính là đã chạy.
- [ ] Trạng thái chỉ từ assertion; BLOCKED có điều kiện gỡ; PENDING có "NEEDS / CẦN: …".
- [ ] Sheet nguồn không đổi; không ghi đè file nguồn.
- [ ] Đã xem mọi bản Excel (mỗi ngôn ngữ) ở các sheet bắt buộc.
- [ ] Dọn dẹp xong; vai trò / môi trường đã khôi phục; snapshot diff sạch (hoặc đã báo).
- [ ] Báo cáo + compare đã ghi vào thư mục run; `runs/index.json` có trạng thái finished.
