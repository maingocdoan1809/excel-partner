# Đánh giá cô lập runtime, fallback và công cụ vận hành

Ngày điều tra: 10/11/2026. Đây là bản đánh giá lịch sử trước triển khai, không phải
mô tả trạng thái hiện tại. Cơ chế trên Office Partner hiện được mô tả tại
[diagnostics.md](diagnostics.md). Xem nghiên cứu nền tảng tại
[provider-runtime-isolation.md](provider-runtime-isolation.md).

## Kết luận

Có thể phát hành mỗi bản add-in với một bộ phiên bản provider riêng. Để triển
khai cho nhiều người dùng, cần đồng thời quản lý bộ phát hành, chuyển đổi người
dùng cũ, khôi phục khi cập nhật lỗi và công cụ chẩn đoán chạy độc lập với bridge.
Chỉ khóa CLI hoặc thêm fallback sang PATH chưa đáp ứng yêu cầu này.

Chưa đủ bằng chứng để chọn một bộ CLI cụ thể làm bản production. Chưa tái hiện
sự cố Copilot người dùng báo và chưa chạy xác nhận live SDK/runtime/Office.
Không thể cam kết loại bỏ lỗi do dịch vụ, đăng nhập hoặc chính sách phía provider.

## Phạm vi nhánh đã kiểm tra

Đã đọc mã ở checkout và so sánh với `origin/main`, không checkout/merge/cherry-pick.
Đã đối chiếu HEAD remote bằng `git ls-remote`; kết quả khớp các ref đang đọc:

| Nhánh | Commit | Package |
| --- | --- | --- |
| `main-office` đang làm việc | `9b8705201b1388686bd159438e35c4bd2e856ce2` | `office-partner@0.1.22` |
| `origin/main` | `e9e8dc586a1b3d7b0b34ade5cd673eede3b1ab5e` | `excel-partner@0.1.22` |
| `main` local | `9151a3a91408ee8f8492989153d5344a990c1428` | Cũ hơn remote 5 commit |

Hai bản `0.1.22` khác package và mã nguồn. ID bộ phát hành phải gồm tên sản
phẩm/package, phiên bản, định danh build/profile và nền tảng; không chỉ dùng semver.

## Các phát hiện cần xử lý trước phát hành

| Phát hiện | Bằng chứng trong code | Hệ quả |
| --- | --- | --- |
| Cập nhật thay package tại chỗ sau khi dừng dịch vụ | `bin/update.mjs:105`, gọi `stop` trước `npm install --global` | Tải/cài lỗi làm gián đoạn cả ứng dụng đang dùng. |
| Recovery chưa phục hồi một bản cũ độc lập | Cùng file, nhánh install lỗi chỉ gọi `start` tại `installedRoot`; setup lỗi sau install thì throw | Không thể gọi đây là rollback được bảo đảm. Thư mục có thể đã bị thay một phần. |
| Xóa npm leftovers trước cài lại | `bin/update.mjs:82` | Không dùng thư mục npm tạm làm bản dự phòng. |
| Update từ pane chỉ kiểm tra bridge đang bận | `bridge/src/server.ts:423` | Gateway có phiên chờ tool result hoặc inference riêng vẫn có thể bị dừng. CLI update cũng chưa dùng cùng cơ chế admission/drain. |
| Setup hiện chưa chọn/cài provider | `bin/office-partner.mjs`, case setup | Không thêm prompt bắt buộc vào mọi lần setup vì updater gọi setup trong quá trình cập nhật không tương tác. |
| Setup không biết chắc provider người dùng đang chọn | `addin/taskpane/ui/provider-picker.ts:7`, lựa chọn trong localStorage | Không thể hứa tự đọc lựa chọn này từ `config.json` rồi cài đúng provider. |
| Mỗi nơi tự chọn executable | Các adapter đọc `*_BIN`/PATH; `bridge/src/provider-login.ts:18` chọn riêng | Chỉ sửa nơi mở chat có thể khiến login, model list và chat dùng các CLI khác nhau. |
| Chỉ có kiểm tra phiên bản tối thiểu | `providers/src/base.ts:60` | Chưa chứng minh tương thích với bản mới; trạng thái ready chưa đồng nghĩa tool round trip hoạt động. |
| Copilot bỏ qua runtime mặc định của SDK | `providers/src/copilot-native.ts:39`, `:51`, `:69` | SDK khóa nhưng CLI vẫn thay đổi theo máy người dùng. |
| SDK Copilot và minimum adapter chưa trùng | SDK `1.0.18` ghi runtime `1.0.94`; adapter minimum `1.0.95` | Không chuyển sang bundled runtime mà chưa kiểm tra các API/flags cần dùng. |
| Chat chưa ghi profile tạo phiên | `addin/taskpane/ui/sessions.ts:50`, `providers/src/session-manager.ts`, `gateway/src/conversations.ts:27` | Native ID một mình không đủ để quyết định resume bằng runtime nào. |
| State provider có thể dùng chung | Copilot title/usage đọc `COPILOT_HOME` hoặc thư mục mặc định; OpenCode merge config user | CLI global vẫn có thể gián tiếp ảnh hưởng qua cấu hình hoặc migration dữ liệu. |
| Doctor add-in còn giới hạn | `bin/office-partner.mjs:250` | Kiểm tra Node, certificate, startup và HTTP; chưa chẩn đoán đầy đủ profile, tool schemas hay execution. |
| Log chưa có lớp lọc chung | `bridge/src/logger.ts:16`, `:37` serialize extra; adapter có thể log stderr | Không đưa nguyên log/environment vào gói hỗ trợ. |
| Dependency còn có khoảng semver | `package.json`; chưa có npm-shrinkwrap | Cùng phiên bản package cài ở thời điểm khác nhau có thể nhận dependency khác. |
| Điều kiện Node chưa đồng bộ | Product khai báo `>=20`, SDK đang cài yêu cầu `^20.19.0 || >=22.12.0` | Bộ manifest và doctor phải kiểm tra giao của các yêu cầu runtime. |
| Uninstall hiện chủ yếu gỡ đăng ký | CLI uninstall giữ dataDir; gateway cũng giữ dữ liệu | Cần bổ sung quản lý quyền sở hữu runtime, không xóa toàn bộ dataDir. |

Gateway đã có trace theo phase và checkpoint `idle/inflight/blocked`. Shared session
manager giữ native identity và không tự thay resume thất bại bằng chat mới. Đây là
nền tảng cần giữ khi bổ sung recovery, không được làm mất vì thêm retry tổng quát.

## Bộ phát hành và bộ khôi phục

Mỗi release khai báo manifest cố định: package/build, phiên bản protocol, Node hỗ
trợ, adapter revision, SDK và CLI chính xác, assets theo OS/architecture, integrity,
khả năng transport, phiên bản schema dữ liệu và các tổ hợp resume/fallback được
kiểm chứng. Bản mới có thể giữ nguyên phiên bản CLI nếu đó là bộ đã kiểm chứng.

Giữ đầy đủ chương trình và dependency tree đã cài cho bản dự phòng; chỉ giữ số
version hoặc tarball thiếu dependencies sẽ không bảo đảm rollback offline.
Npm-shrinkwrap hoặc artifact đã dựng đầy đủ là các phương án cần lựa chọn khi
đóng gói. Copilot SDK hiện là dependency trực tiếp có platform runtime đi kèm:
việc chọn provider ở setup chưa tự động làm package tải đúng những provider được
chọn. Nếu yêu cầu tải tối thiểu, cần tách gói provider tương ứng.

Đề xuất cấu trúc logic:

```text
<application-data>/
  recovery/                 # launcher và công cụ cứu hộ độc lập
  releases/<release-id>/    # chương trình và dependency tree cố định
  runtimes/<profile-id>/    # đầy đủ executable, loader, assets của provider
  deployment/               # active, previous, ownership, journal, locks
  diagnostics/              # log có giới hạn, gói hỗ trợ cục bộ
```

Không đặt bản dự phòng duy nhất trong thư mục package mà npm sẽ thay/xóa. Launcher
và đường phục hồi cần hoạt động khi package global thiếu hoặc bridge không lên.
Phải thiết kế luôn cách bootstrap cập nhật, giữ tương thích và khôi phục chính nó.

Rollback toàn bộ release là đường mặc định: add-in assets, bridge, gateway,
adapter, SDK và CLI đi cùng nhau. Chỉ đổi CLI trong bản app hiện tại khi manifest
đã xác nhận tương thích. SDK JS nạp trong tiến trình cũ không tự đổi theo một
biến đường dẫn; đổi profile cần đóng worker/client cũ và tạo lại đúng bộ.

## Quy tắc fallback

| Tình huống | Hành vi đề xuất |
| --- | --- |
| Tải, integrity, giải nén hoặc preflight bản mới lỗi | Giữ release/profile đang hoạt động, ghi trạng thái thất bại, cho thử lại. Chưa dừng bản cũ. |
| Khởi động bản mới lỗi sau chuyển active | Recovery launcher trả về bộ previous đã kiểm chứng và kiểm tra dịch vụ lại; chỉ thực hiện trong số lần giới hạn. |
| Runtime thiếu file hoặc hỏng | Sửa đúng profile từ artifact được xác minh; trong lúc đó dùng bộ previous tương thích nếu có. Không lấy latest. |
| Đăng nhập hết hạn, policy từ chối, model không được phép | Phân loại lỗi và đưa đúng bước xử lý. Không tự đổi provider, account hay đường thực thi. |
| CLI sai protocol/tool schemas | Chặn profile đó cho phiên mới; về profile tương thích trước inference nếu có bằng chứng an toàn. |
| Lỗi xảy ra sau khi gửi prompt hoặc trả tool_use | Giữ bằng chứng và conversation ID; không tự phát lại prompt hoặc tool call. Trạng thái chưa rõ phải được báo rõ. |
| Resume chat cũ lỗi | Giữ lịch sử và ID gốc; thử profile cũ chỉ nếu state còn tương thích và chưa phát inference. Không âm thầm tạo chat thay thế. |
| Mất mạng/dịch vụ provider lỗi | Giữ cài đặt; retry có giới hạn nếu thao tác đủ điều kiện. Không coi mọi lỗi mạng là CLI regression. |
| Không còn bộ dự phòng hợp lệ | Báo degraded với mã lỗi và đường sửa cài đặt. Không báo đã rollback thành công. |

Thứ tự quyết định phụ thuộc định danh conversation: chat có profile dùng chính
profile đó; chat mới dùng active đã kiểm chứng. Với người dùng cũ chưa chuyển đổi,
giữ đường CLI ngoài đã cấu hình trước đó làm lựa chọn tạm thời có ghi nhận. Chỉ
cho fallback khi tổ hợp adapter/runtime/state được xác minh tương thích. Không
mặc định chạy bất kỳ `copilot`/`codex` mới nào vừa xuất hiện trên PATH.

Một executable tìm thấy trên máy không phải bằng chứng nó đã chạy tốt. Ghi rõ
nguồn: phát hiện, qua preflight, hay đã hoàn thành turn. Nếu file CLI ngoài bị
thay đổi, vô hiệu kết luận tương thích cũ. Không tự sao chép CLI global thành
runtime riêng nếu chưa biết đầy đủ distribution và cơ chế cập nhật của nó.

Giới hạn số lần fallback, ghi lý do, cách chọn bộ đích và kết quả; không tạo vòng
lặp hai phiên bản. Bộ bị đánh dấu lỗi không tự được chọn lại khi khởi động sau.

## Quy trình cập nhật và rollback

```text
Đang dùng A
  -> tải/dựng B riêng -> xác minh B -> preflight B
  -> ngăn nhận tác vụ mới -> đợi bridge và gateway rảnh
  -> checkpoint dữ liệu/phục hồi -> dừng dịch vụ A
  -> chuyển active sang B -> khởi động và xác minh B
  -> xác nhận B hoặc quay lại A có kiểm tra
```

Lưu journal bền vững qua từng bước và cập nhật pointer bằng thao tác thay file
được kiểm tra trên Windows. Mất điện ở bước nào cũng phải xác định được active,
previous và phần cài còn dang dở ở lần khởi động tiếp theo. Có khóa liên tiến
trình và kiểm tra lại trước chuyển đổi; kiểm tra busy một lần tại HTTP request
không ngăn tác vụ mới đến sau đó.

Gateway chờ tool_result cũng là đang bận. Không đổi release khi có continuation
đang chờ. Những cuộc hội thoại inflight không được coi là an toàn để replay sau
restart. Luôn giữ ranh giới thực thi Office của official client.

Nếu người dùng yêu cầu rollback, hiển thị phiên bản đích, provider profiles và
tình trạng tương thích dữ liệu. Bản đã tải và kiểm chứng phải phục hồi offline.
Hỗ trợ chọn bản cũ khác bằng phiên bản chính xác, kiểm tra tính hợp lệ trước khi
dừng bản hiện tại. Sau rollback, giữ lựa chọn đích để tránh vòng lặp tự cập nhật
ngay về bản vừa lỗi. Rollback phần mềm không undo thao tác đã sửa tài liệu.

Giữ tối thiểu active và previous hoàn chỉnh. Profile còn được conversation giữ
tham chiếu không được dọn tự động. Việc xóa/di chuyển phải kiểm tra đường tuyệt
đối nằm trong thư mục ứng dụng và không còn worker đang dùng; xử lý Windows file
locks. Không sửa quyền hoặc xóa CLI global, thông tin đăng nhập hay thư mục user.

## Chuyển đổi người dùng cũ

1. Thêm state schema mở rộng tương thích trước, ở chế độ giữ hành vi cũ. Ghi
   package/build, executable thực tế và phiên bản khi tạo/tiếp tục phiên; không
   suy đoán phiên bản của những chat đã lưu trước đó.
2. Tách setup tương tác và reconcile không tương tác. Setup lần đầu cho chọn
   nhiều provider; update/autostart không hỏi qua stdin. Lưu lựa chọn dùng chung
   có schema, đồng thời giữ khóa localStorage và cấu hình cũ.
3. Với lựa chọn chưa biết, tiếp tục hỗ trợ provider cũ và cho task pane xác nhận
   lựa chọn sau khi kết nối. Gateway có thể bổ sung inventory từ enabledProviders.
   Cài nền không đồng nghĩa đã kích hoạt profile mới cho chat cũ.
4. Dựng runtime riêng mà không ghi đè CLI cũ. Kiểm tra login bằng cơ chế chính
   thức; không sao chép token hoặc yêu cầu API key. Xác minh session/cache/config
   separation theo provider, tránh migration state làm hỏng đường quay lại.
5. Profile mới được chọn sau các kiểm tra cần thiết. Phiên đã có profile tiếp
   tục đúng bộ; chat cũ chưa có mapping giữ trạng thái chưa rõ cho đến khi có đủ
   bằng chứng. Provider tùy chỉnh qua `*_BIN` được giữ dưới chế độ explicit external.

Cần xử lý riêng lần nâng cấp đầu tiên từ updater cũ: mã cũ vẫn dừng dịch vụ và
ghi đè package trước khi setup bản mới được chạy. Bản mới không thể retroactively
backup bản cũ từ trước bước đó. Muốn bảo vệ giao dịch nâng cấp này cần helper
migration độc lập chạy trước, hoặc quy trình triển khai bên ngoài có snapshot.
Một release chuyển tiếp chỉ bảo vệ được các lần cập nhật sau khi nó đã cài xong.
Không được quảng bá lần nâng cấp đầu tiên là đã có đầy đủ rollback nếu chưa có
đường bootstrap này.

Gỡ add-in chỉ giải phóng owner add-in; gateway hoặc bản dự phòng còn cần runtime
thì giữ nó. Gỡ toàn bộ ứng dụng dọn các runtime/release thuộc sở hữu ứng dụng,
giữ dữ liệu người dùng mặc định. Npm hiện không có uninstall lifecycle hook:
`npm uninstall -g` không bảo đảm dọn runtime ở dataDir. Cần lệnh uninstall của
ứng dụng hoặc bộ cài Windows quản lý đầy đủ vòng đời; tài liệu phải nói rõ đường
gỡ được hỗ trợ.

## Công cụ vận hành cần có

Các tên lệnh dưới đây là giao diện đề xuất, chưa tồn tại đầy đủ trong sản phẩm.
Cả `excel-partner` trên main và `office-partner` dùng chung hành vi tương ứng.

| Công cụ | Yêu cầu |
| --- | --- |
| `doctor --json` | Báo cáo có schema/version và exit code ổn định; package/build, OS/arch, Node, ownership, active/previous, artifact integrity, executable được chọn và lý do, transport, login, protocol, catalog, bridge/gateway, TLS, cổng, proxy/network. Mỗi lỗi có bước xử lý cụ thể. |
| `doctor --provider <id>` | Kiểm tra từng provider độc lập, có timeout, phân biệt không cài/chưa chọn/hỏng/chưa đăng nhập/không tương thích. Provider không chọn không làm cả cài đặt báo lỗi. |
| `diagnostics export` | Xuất ZIP cục bộ gồm metadata cho phép, sự kiện có cấu trúc và kết quả doctor. Có preview, giới hạn kích thước/thời gian và thông tin đã loại; không tự gửi ra ngoài. |
| `repair` | Dựng lại đúng artifact của release/profile đã chọn, khôi phục registration cần thiết sau khi kiểm tra; lặp lại an toàn. Không tự nâng lên latest hoặc reset cấu hình/credentials/chat. |
| `versions` / `rollback <version>` | Xem active/previous/cached, kiểm tra rollback trước khi chạy, hỗ trợ phục hồi offline và báo kết quả thật. |
| `providers list/install/remove` | Chọn provider sau setup, cài phiên bản theo manifest, ghi nhận ownership, không dọn runtime còn dùng. |
| `debug on/off` | Bật log chi tiết có thời hạn, tự hết hạn, rotation và giới hạn dung lượng; không ghi bí mật hoặc dữ liệu tài liệu. |
| Recovery launcher | Doctor, repair và rollback vẫn chạy khi bridge/gateway chết hoặc package global hỏng. Có đường gọi được lưu rõ lúc setup. |

Settings cần có kiểm tra hệ thống, xuất chẩn đoán, sửa cài đặt và danh sách bản
có thể quay lại. Chức năng tương tự trong launcher dùng được khi task pane mất
kết nối; không phụ thuộc người dùng mở DevTools hay tự ghép log từ nhiều nơi.

Mã lỗi có cấu trúc cần tách các phase: download, verify, unpack, spawn,
handshake, auth, policy, model access, catalog registration, invocation,
result conversion, native resume và Office execution. Ghi correlation ID xuyên
bridge/provider/MCP/gateway, các mốc thời gian và exit/error code, kèm profile.
Không chỉ lưu thông báo chung `provider_error` rồi yêu cầu cập nhật CLI global.

Gói hỗ trợ dùng allowlist; mặc định không chứa environment dump, credentials,
token, MCP config có token, prompt, tool arguments/results, workbook/document,
session title, đường dẫn cá nhân hoặc account identifier. Chuẩn hóa các đường
dẫn thành nhãn và băm các ID/tên MCP tùy chỉnh khi cần liên kết sự kiện. Lọc cả
stderr và thông báo exception, không chỉ trường JSON có tên token. Raw CLI logs
không mặc nhiên an toàn để đưa vào ZIP. Retention và log rotation cần hữu hạn.

## Đưa cơ chế sang main

`providers/src/base.ts`, `registry.ts`, `session-manager.ts`, `codex.ts`,
`opencode.ts`, `bridge/src/provider-login.ts` và `scripts/package-windows.ps1`
hiện giống nhau giữa hai ref đã đối chiếu. Đây là các điểm tích hợp thuận lợi.

Đặt resolver/profile và hợp đồng runtime ở lớp provider độc lập Office host;
đặt installer/recovery/diagnostics trong module dùng chung của CLI. Không buộc
main phải nhận toàn bộ `runtime/src` và công cụ đa host để dùng version manager.
Mỗi package khai báo entrypoint và release identity qua metadata; không hardcode
tên `office-partner` vào core. Build phải đóng gói đầy đủ các file dùng chung.

Main còn đường cài source riêng: `install.ps1`/`scripts/install.ps1` có chọn
provider và chế độ native/WSL; `scripts/install-windows.sh` bỏ qua CLI đã có hoặc
cài npm global khi thiếu. Các script update source chạy git pull, npm ci/build
trước restart. Khi đưa cơ chế mới sang main, cần chuyển các đường này về cùng
manifest/resolver hoặc giữ chế độ external có thông báo rõ. Không chỉ sửa CLI
npm rồi coi mọi người dùng main đã được chuyển. Windows và WSL cần profile,
thư mục state và artifact đúng môi trường; không tự chuyển credential giữa hai
môi trường. Các kiểm thử main phải bao gồm đường source/native/WSL còn hỗ trợ.

Đề xuất phát triển các commit hạ tầng nhỏ từ baseline main mới nhất, mang sang
main-office bằng cherry-pick có kiểm tra adapter/UI từng nhánh. Không merge toàn
bộ main-office vào main chỉ để có runtime isolation. Chạy kiểm tra độc lập trên
cả hai nhánh và artifact npm của từng package trước phát hành.

Đặc biệt cần giữ các sửa lỗi mới ở `e9e8dc5` của main:

- Process selector có ranh giới đường dẫn, xác định đúng package/prefix và
  target của updater; bảo vệ ancestors và Excel. Main-office còn selector rộng
  hơn dựa trên substring/cwd. Áp dụng lại phần bảo vệ cho cả Word/PowerPoint và
  owner runtime khi tích hợp, không đưa selector cũ từ main-office đè lên main.
- Update runner dừng bridge tường minh, kiểm tra gateway restart và thử phục
  hồi gateway trên đường lỗi. Hai test file riêng cho selector/runner có trên
  main mà không có trong main-office; cần giữ các kịch bản này.
- Copilot images có thay đổi riêng ở main. Dùng diff nhỏ khi chèn resolver,
  không thay toàn bộ adapter bằng bản ở nhánh còn lại.

Main-office còn thêm `--system-prompt-snapshot off` cho Claude và danh mục tool
theo Office host. Các yêu cầu này phải nằm trong kiểm chứng profile của nhánh
tương ứng. Profile chỉ tương thích với main chưa tự chứng minh tương thích với
main-office.

Thư mục `legacy/excel-partner` ở main-office là gói chuyển tên sang
`office-partner@latest`. Chưa kiểm tra trạng thái publish trên npm. Nếu tiếp tục
duy trì main/excel-partner, cần chính sách phát hành rõ để shim không vô tình
chuyển người dùng main sang dòng sản phẩm khác. Hai package còn dùng chung dataDir,
startup name và cổng bridge mặc định; không giả định chúng chạy song song độc lập.

## Kiểm chứng đã thực hiện

Trên main-office `9b87052`, trước khi sửa mã chức năng:

- `npm run typecheck`: qua.
- `npm run build`: qua, gồm server, add-in và gateway.
- `npm test -- --exclude '**/*.live.test.ts'`: 136 file; 134 qua, 2 lỗi.
  Tổng 1.740 ca: 1.738 qua, 2 lỗi. Đã loại các file live để không gọi provider thật.
- Chạy lại riêng hai file lỗi: cả hai lỗi đều tái hiện, 11 ca còn lại qua.
- `tests/legacy-shim.test.ts:86`: expected đường dẫn thiếu drive so với kết quả
  `path.resolve` trên Windows. Cần sửa cách tạo kỳ vọng và kiểm chứng riêng hành vi
  phân giải đường dẫn; đây chưa phải bằng chứng migration production bị hỏng.
- `tests/mcp-host.stdio.test.ts:31`: `spawnSync(tsx.cmd, ...)` không qua shell trả
  `status=null`, `error.code=EINVAL`. Kiểm tra chẩn đoán riêng cùng CLI qua
  `node node_modules/tsx/dist/cli.mjs` trả `status=2` và thông báo OFFICE_HOST như
  mong đợi. Không sửa assertion hoặc code trong đợt điều tra này.
- Vitest ban đầu bị sandbox chặn khi đọc config; đã chạy lại ngoài sandbox.
  Đây là lỗi môi trường kiểm tra riêng với hai lỗi baseline nêu trên.

Không chạy tests trên checkout main; đánh giá main ở trên là đọc mã và diff.
Không có cơ sở kết luận main đã qua kiểm thử chỉ từ kết quả main-office. Không
thay nhánh, phát hành, cài/gỡ CLI, đăng nhập provider hoặc chạy suy luận live.

## Điều kiện chấp nhận trước triển khai rộng

1. Cả hai nhánh qua typecheck/build/mock tests, gồm hai lỗi baseline và các test
   updater của main. Kiểm tra artifact đã đóng gói, không chỉ source checkout.
2. Fault injection ở mọi bước cài/update/rollback: offline, proxy/TLS, hết đĩa,
   integrity sai, quyền ghi, antivirus/file lock, mất điện, update đồng thời,
   khởi động ở Windows sign-in và lỗi gateway. Bản cũ phải khởi động lại được
   khi bridge/package mới hỏng, bằng công cụ cứu hộ độc lập.
3. Tái hiện migration từ package cũ đang phát hành, đường source install cũ,
   CLI tùy chỉnh, người dùng chỉ gateway và người dùng chỉ add-in; không làm mất
   lựa chọn, lịch sử hoặc native identity. Kiểm tra riêng lần nâng cấp đầu tiên.
4. Chứng minh thay CLI/PATH global không đổi runtime managed; resume có profile;
   catalog MCP bật/tắt cập nhật schemas đúng và giữ conversation; cancellation
   không để lại worker; không tự phát lại tool trong mọi nhánh recovery.
5. Live checks opt-in cho từng cặp SDK/CLI và từng host được hỗ trợ: login,
   discovery, một tool round trip, lỗi tool, resume và gateway continuation.
   Kiểm tra ghi dùng tài liệu thử nghiệm. Chính sách auth/state chưa rõ phải
   được giải quyết trước khi chuyển người dùng sang managed mode mặc định.
6. Kiểm chứng lọc gói diagnostics bằng dữ liệu mồi có token, paths, prompts và
   nội dung tài liệu; công cụ repair không xóa data; uninstall không tác động
   CLI user/gateway còn hoạt động; cleanup không xóa bộ cần cho rollback.
7. Pilot nhóm nhỏ, giữ bản dự phòng, rồi mở rộng theo đợt có thể dừng. Dùng release
   channel/phiên bản được duyệt và cơ chế giữ bản sau rollback. Thu thập bằng gói
   hỗ trợ do người dùng chủ động xuất; không tự bổ sung telemetry/upload.

## Tài liệu chính thức bổ sung

- [npm lifecycle scripts](https://docs.npmjs.com/cli/v11/using-npm/scripts/#a-note-on-a-lack-of-npm-uninstall-scripts): npm hiện không chạy uninstall lifecycle hooks.
- [npm-shrinkwrap](https://docs.npmjs.com/cli/v11/configuring-npm/npm-shrinkwrap-json/): lockfile có thể phát hành cùng ứng dụng CLI; package-lock của repo không thay thế vai trò này khi publish.
- [Node child_process trên Windows](https://nodejs.org/api/child_process.html#spawning-bat-and-cmd-files-on-windows): cách khởi chạy `.cmd`/`.bat` khác executable thực.

Ưu tiên thực hiện sau đánh giá: củng cố recovery/diagnostics và lưu metadata tương
thích ở cả hai nhánh; sau đó mới triển khai cài runtime riêng và chuyển người dùng.
