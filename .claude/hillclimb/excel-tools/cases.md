# Excel assistant: tool choice and arguments (proposed cases)

**Flow:** one request to the Excel assistant in the pane, which has Excel's tools. **Question the eval answers:** does the edit to the prompt and tool descriptions proposed by the audit (variant `v1`) make the assistant call the right tool with the right arguments, or report worse than today's (`baseline`)?

**How it runs (proposed):** the real bridge, the real provider (Claude), and a small in-memory workbook standing in for Office.js. Each case starts from one fixture below, so the end state can be graded as well as the tool calls. Nothing here touches your real workbooks.

**Where the cases came from:** 32 distinct messages in your real chat history. About 20 are greetings, tests and chit-chat ("xin chào", "1 + 1 bằng mấy"). Eleven are real tasks: six about drawing or rearranging diagrams, five about something else. I kept those five as seeds (marked `real`; emails, paths, URLs and long numbers are already stripped). The rest are marked `synth`: variations I wrote around those seeds and around the one tool the audit changes, `insert_rows`.

**Added at your request:** three more kinds of work, all `synth` and built on the same fixtures: adding data (`add-data`), calculation and analysis (`analysis`), and reports (`report`). Their numbers are fixed in the fixture files, so every answer can be checked against a number computed in code, not against another model.

**Not covered, and why:** the six diagram messages ("sắp xếp lại sơ đồ", "chụp ảnh lại và xem") are over half of your real tasks. An in-memory workbook cannot render shapes, so those need real Excel. Say so if you want them added later; for now they are out.

## Fixtures

| id | workbook | contents |
|---|---|---|
| F1 | Sales.xlsx | sheet `Data`, A1:D11: headers `Product, Qty, Price, Total`; 10 product rows; column D empty |
| F2 | Report.xlsx | sheet `Doanh thu`, A1:C11: headers `Tháng, Doanh thu, Chi phí`; 10 monthly rows; C7 holds `#DIV/0!` |
| F3 | Notes.xlsx | sheet `Ghi chú`, A1:A6: six short Vietnamese phrases, two with typos (`Thág 3`, `Hà Nôi`) |
| F4 | Orders.xlsx | sheet `Orders`, A1:D13: `Order, Product, Qty, Price`; 12 rows with three planted problems: a negative `Qty` (C5), text in `Price` (D7), a duplicated order id (A4 and A11) |

## Cases

| id | tags | source | fixture / selection | what a pass looks like |
|---|---|---|---|---|
| R1 | read | real | F1 + F2 in one workbook | Reads every sheet and lists each sheet name and used range. No write tools. |
| R2 | read | real | F1, `B3:C4` selected | Answers from the selection (names `B3:C4`). No write tools. |
| R3 | read | synth | F2 | Finds `C7` as the error cell with a read tool. No writes. |
| R4 | read | real | F1 | Answers in chat what it can do, without calling a tool or changing anything. |
| P1 | python | real | F1 | Calls `run_python` for 6\*7 and says 42. |
| W1 | write | real | F3, `A1:A6` selected | Fixes only the two typos. Other cells unchanged. |
| W2 | write | synth | F2 | `B12` holds a `SUM` formula over `B2:B11`, not a typed number. No other cell changed. |
| W3 | write | synth | F1 | `D2:D11` are formulas using that row's `B` and `C`, not constants. Header row untouched. |
| W4 | write | synth | F1 | Does not overwrite `C2:C4` silently: workbook unchanged and the reply asks to confirm. |
| W5 | write | synth | F1 | Does not hide or delete anything for a vague complaint; suggests or makes a formatting change. |
| I1 | insert | synth | F1 | `insert_rows` with row 5, count 3, above. The old row 5 is now row 8. |
| I2 | insert | synth | F1 | One blank row directly under the header; data below it shifted down by one. |
| I3 | insert | synth | F1 | Two blank rows below row 10 (`A10` given as a cell reference). |
| I4 | insert | synth | F2 | One blank row above row 3 of sheet `Doanh thu` (sheet named in the request). |
| I5 | insert | synth | F1 | 1500 rows exceed the tool's limit of 1000. Either two calls totalling 1500, or it reports the limit. It never claims 1500 if fewer were added. |
| I6 | insert | synth | F1 | Request has no position or count. Asks which row and how many; no write. |
| I7 | insert | synth | F1 | A request to delete row 4 must not call `insert_rows`. Deletes row 4 or asks to confirm. |
| A1 | add-data | synth | F1 | One new row at row 12 (`Cam`, 5, 15000), directly under the data: no gap, no overwrite, numbers stored as numbers. |
| A2 | add-data | synth | F1 | Three new rows at 12 to 14 that match the columns' types (text, number, number). Existing rows unchanged. |
| A3 | add-data | synth | F1 | New header in `E1` and the same value in `E2:E11`. Nothing else changed. |
| A4 | add-data | synth | F2 | A new row below the last month with `Tháng 11` and the amount as a number, either 52000000 or 52000 (the sheet's unit is not stated, and a reply that explains the unit it chose is fine); never the text `52 triệu`. |
| N1 | analysis | synth | F2 | Reply gives the correct total and average of `Doanh thu`, read from the workbook. No writes. |
| N2 | analysis | synth | F2 | Names the correct highest and lowest month. No writes. |
| N3 | analysis | synth | F2 | `D2:D11` are formulas `B - C` per row. `D7` shows the error from `C7`, and the reply says so instead of reporting success. |
| N4 | analysis | synth | F1 | Reply gives the correct largest Qty × Price among the rows. No writes. |
| N5 | analysis | synth | F4 | Names the three planted problems (`C5`, `D7`, the duplicate order id). No writes. Judged, because a reply may say "row 5, Qty" instead of "C5". |
| B1 | report | synth | F2 | Total, average, highest and lowest month as formulas over the source (on a new sheet or beside the data), not typed numbers. The original cells are unchanged. |
| B2 | report | synth | F2 | A new sheet with one row per quarter, formulas over the source months. The reply notes Q4 holds only one month. |
| B3 | report | synth | F2 | A chart over `A1:B11` (headers in row 1, months in column A). Tool call checked; nothing is rendered. |

29 cases. Tags: `read` 4, `python` 1, `write` 5, `insert` 7, `add-data` 4, `analysis` 5, `report` 3. `tags[0]` is the section header in the report.

## Inputs, one per case

R1
```
Read all the data in this workbook. For each worksheet, list the sheet name, the used range address, and the contents of the used range (values, not formulas).
```

R2
```
đang chọn gì
```

R3
```
Có ô nào đang bị lỗi công thức không?
```

R4
```
bạn làm được gì
```

P1
```
Use the run_python tool to compute 6*7 and tell me the result.
```

W1
```
sửa lỗi chính ta đi
```

W2
```
Tính tổng cột B vào ô B12
```

W3
```
Điền cột D = Qty × Price cho tất cả các dòng
```

W4
```
Ghi chữ "n/a" vào C2:C4
```

W5
```
Cột C nhìn rối quá
```

I1
```
Chèn 3 dòng trống phía trên dòng 5
```

I2
```
Insert one blank row right below the header row.
```

I3
```
Add 2 blank rows below A10.
```

I4
```
Thêm 1 dòng trống trên dòng 3 của sheet Doanh thu
```

I5
```
Thêm 1500 dòng trống dưới dòng 2
```

I6
```
Chèn thêm vài dòng
```

I7
```
Xóa dòng 4
```

A1
```
Thêm sản phẩm mới: Cam, số lượng 5, giá 15000 vào cuối bảng
```

A2
```
Add three more rows of sample data in the same format as the existing ones.
```

A3
```
Thêm cột Ghi chú ở cuối bảng, ghi "Còn hàng" cho mọi dòng
```

A4
```
Nhập doanh thu tháng 11 là 52 triệu
```

N1
```
Tổng và trung bình doanh thu cả năm là bao nhiêu?
```

N2
```
Tháng nào doanh thu cao nhất, tháng nào thấp nhất?
```

N3
```
Thêm cột Lợi nhuận = Doanh thu - Chi phí
```

N4
```
Đơn hàng có giá trị lớn nhất là bao nhiêu (số lượng × giá)?
```

N5
```
Kiểm tra dữ liệu giúp tôi, có gì bất thường không?
```

B1
```
Lập báo cáo tóm tắt doanh thu: tổng, trung bình, tháng cao nhất và thấp nhất.
```

B2
```
Tạo báo cáo doanh thu theo quý ở một sheet mới
```

B3
```
Vẽ biểu đồ cột doanh thu theo tháng
```

## Honest limits to read before approving

- **Noise.** 29 cases at 3 reps each is about ±11 points on the overall pass rate (1/√87). The `insert` group alone is still 7 cases, about ±22 points (1/√21); the other groups are 3 to 5 cases, so none of them can show a small change on its own. The audit's change to `insert_rows` can only show up as a large swing in that group (for example from 3 of 7 to 6 of 7). A small gain will look like noise. More `insert` cases or more reps would narrow this; each costs more model calls.
- **Mostly synthesized.** 5 of 29 cases are real. The `insert`, `write`, `add-data`, `analysis` and `report` cases are my guess at what you ask for; your history has no real example of adding data, analysing or reporting.
- **Reply wording is graded loosely.** For `analysis` and `report`, the numbers and cell references in the reply are checked in code. The prose around them is not. If real Excel edits look different from these, say so now.
- **Prompt scope.** In Excel the add-in's own prompt is empty by default (only the optional "recommended" preset has text). So the baseline here has no add-in system prompt, and `v1` differs from it only by the `insert_rows` description. The wording edits to the Word and PowerPoint prompts are not measured by this eval.

## Changes to the grading after the first full run

Three checks were too strict and were corrected, then those cases were run again for both variants: `I5` now counts rows inserted through `modify_structure` too; `A4` accepts 52000 as well as 52000000; `B1` no longer requires a new sheet. `N5` now relies on the judge alone, because the code check demanded the literal text `C5` and `D7`, and a correct reply that says "row 5" failed it. The oracle (29 of 29) and null (0 of 29) checks were rerun after each change.
