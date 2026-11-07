# TÀI LIỆU HƯỚNG DẪN TỔNG HỢP TOÀN BỘ CÁC LỆNH & HÀM
## BỘ CÔNG CỤ AVTOOL & THIẾT KẾ ÁO ĐƯỜNG KCAD PRO 2026
*(Tích hợp TCCS 38:2022/TCĐBVN, 22 TCN 211-06 & Xử lý Dữ liệu Kỹ thuật Chuyên sâu)*

---

### PHẦN I. CÁC LỆNH TRÊN THANH RIBBON (GIAO DIỆN ĐỒ HỌA)

#### 1. TAB [AVTool] - TIỆN ÍCH DỮ LIỆU & DỰ TOÁN
| Nhóm chức năng | Tên nút trên Ribbon | Phím tắt / Lệnh | Mô tả chi tiết |
| :--- | :--- | :---: | :--- |
| **Chuyển đổi Font** | **Đổi Font** | `Ctrl + Q` | Mở hộp thoại trực quan hỗ trợ quét và đổi font TCVN3 (.VnTime), VNI sang Unicode. Tự động bảo vệ công thức và hình ảnh. |
| | **Auto -> Uni** | | Tự nhận diện mã nguồn từng ô và chuyển sang Unicode. |
| | **TCVN3 -> Uni** | | Chuyển nhanh vùng chọn từ TCVN3 (.VnTime) sang Unicode. |
| | **VNI -> Uni** | | Chuyển nhanh vùng chọn từ VNI-Times sang Unicode. |
| | **Uni -> TCVN3** | | Chuyển ngược từ Unicode về TCVN3. |
| | **Uni -> VNI** | | Chuyển ngược từ Unicode về VNI. |
| **Đổi Kiểu Chữ** | **CHỮ HOA** | | Chuyển toàn bộ văn bản sang IN HOA (hỗ trợ tiếng Việt đầy đủ). |
| | **chữ thường** | | Chuyển toàn bộ văn bản sang chữ thường. |
| | **Ký Tự Đầu** | | Viết Hoa Ký Tự Đầu Mỗi Từ (Proper / Title Case). |
| **Tiện Ích Khối Lượng** | **Chèn DValue Vào File** | | Tự động nhúng hàm `DValue` vào file Excel (.xlsm) để gửi cho CĐT/thẩm tra mở trên máy khác không cần cài Add-in. |
| | **Sửa Lỗi Diễn Giải** | | Chuẩn hóa các dấu ngoặc hoặc ký tự sai trong diễn giải dự toán. |
| | **Đổi Dấu Chấm/Phẩy** | | Mở form tùy chọn đổi dấu phân cách thập phân `.` thành `,` hoặc ngược lại. |
| | **Mở Dòng Ẩn & Bôi Màu** | | Quét toàn bộ sheet hoặc cả file, mở tất cả dòng bị ẩn và tô màu vàng nhạt (RGB 255, 255, 153) để dễ kiểm soát sai sót/chi phí giấu. |
| | **In Hàng Loạt** | | Form chọn nhiều sheet, sắp xếp thứ tự in, Print Preview hoặc in hàng loạt. |
| | **Cập Nhật Online** | | Kết nối máy chủ GitHub kiểm tra và tự tải bản cập nhật mới nhất. |
| **Nội Suy & Bảng Tra** | **Nội Suy Ma Trận 2D** | | Chèn cú pháp và hướng dẫn nội suy 2 chiều song tuyến tính (Bilinear). |
| | **Nội Suy Dãy Số 1D** | | Chèn cú pháp nội suy dãy số 1 chiều. |
| | **Nội Suy 2 Chiều** | | Chèn hàm nội suy 2 mốc kẹp. |
| | **Tạo Bảng Mẫu** | | Tự sinh ra 1 Sheet mẫu chứa đầy đủ bảng tra 1D, ma trận 2D đã lập sẵn công thức. |
| | **Hướng Dẫn Nội Suy** | | Mở hộp thoại tóm tắt cú pháp các hàm nội suy. |

---

#### 2. TAB [KCAD] - THIẾT KẾ KẾT CẤU ÁO ĐƯỜNG MỀM (TCCS 38:2022 & 22 TCN 211-06)
| Nhóm chức năng | Tên nút trên Ribbon | Cú pháp chèn vào ô | Tiêu chuẩn / Toán đồ |
| :--- | :--- | :---: | :--- |
| **Mẫu Bảng In Chuẩn** | **MĐ Làm Mới (TCCS 38:2022)** | *(Tạo Workbook mới)* | Mở trọn bộ file mẫu 4 sheet (`ÁO ĐƯỜNG MỀM-38.2022`, `Bảng tra`, `Toán đồ Kogan`, `Toán đồ TCCS 38-2022`) đã căn lề in A4 chuẩn, tên công trình để trống. |
| | **Tăng Cường Mặt Đường Cũ** | *(Tạo Workbook mới)* | Mở trọn bộ file mẫu 11 sheet (các kết cấu Loại 1, 2A, 2B, 3, 4, các toán đồ trượt, uốn, võng). |
| **Kiểm Toán Cường Độ** | **Tính Eyc** | `=TinhEyc()` | Tính mô đun đàn hồi yêu cầu $E_{yc}$ (MPa) theo số trục xe tích lũy $N_e$, cấp mặt đường và độ tin cậy. |
| | **Tra Beta** | `=TraBeta()` | Tra hệ số hiệu chỉnh ứng suất kéo uốn $\beta$ theo Bảng 3.6 TCCS 38:2022. |
| | **Tra Ech/E0** | `=TraEch()` | Tra mô đun đàn hồi chung hệ 2 tầng theo Toán đồ Kogan (Hình 3.1). |
| | **Hệ Số Kcđ** | `=TraKcd()` | Tra hệ số cường độ yêu cầu $K_{cđ}$ theo độ tin cậy $P$ (0.85, 0.90, 0.95, 0.98). |
| **Kiểm Toán Chống Trượt** | **Tra Tax/p2** | `=TraTaxHD2()` | Tra ứng suất cắt trượt tác dụng $\tau_{ax}/p$ với trường hợp $H/D = 2$ (Toán đồ 3.2). |
| | **Tra Tax/p4** | `=TraTaxHD4()` | Tra ứng suất cắt trượt tác dụng $\tau_{ax}/p$ với trường hợp $H/D = 4$ (Toán đồ 3.3, 3.4). |
| | **Tra Tav** | `=TraTav()` | Tra ứng suất cắt trượt do trọng lượng bản thân $\tau_{av}/p$. |
| **Kiểm Toán Kéo Uốn** | **Ứng Suất Tầng Mặt** | `=TinhUS_tangmat()` | Tra ứng suất kéo uốn đáy tầng mặt bê tông nhựa $\sigma_{ku}/p$ (Toán đồ 3.5). |
| | **Ứng Suất Tầng Móng** | `=TinhUS_tangmong()` | Tra ứng suất kéo uốn đáy tầng móng gia cố xi măng $\sigma_{ku}/p$ (Toán đồ 3.6). |
| **Trợ Giúp** | **Hướng Dẫn KCAD** | | Mở popup tóm tắt công thức, ý nghĩa tham số hiển thị tiếng Việt Unicode có dấu 100%. |

---

### PHẦN II. CÁC HÀM "LỆNH ẨN" GÕ TRỰC TIẾP VÀO Ô TÍNH (UDF)
*(Bạn chỉ cần đặt chuột vào ô bất kỳ và gõ dấu `=` cùng tên hàm, hệ thống sẽ tự động tính toán)*

#### 1. HÀM TÍNH KHỐI LƯỢNG DỰ TOÁN TỪ DIỄN GIẢI CHỮ
* **Cú pháp:** `=DValue(chuỗi_diễn_giải)`
* **Ý nghĩa:** Tự động nhận diện các phép tính số học, biểu thức ngoặc và loại bỏ toàn bộ từ ngữ thuyết minh tiếng Việt, đơn vị đo.
* **Đặc tính hỗ trợ:**
  - Hỗ trợ các loại ngoặc lồng nhau: `()`, `[]`, `{}`.
  - Hỗ trợ mọi dấu nhân chia: `*`, `x`, `X`, `/`, `:`, `^`, `%`.
  - Tự động bỏ qua đơn vị đo: `m`, `m2`, `m3`, `kg`, `tấn`, `md`, `cái`, `bộ`, `khoang`, `đoạn`, `bên`, `cánh`, `trụ`, `mố`, `dầm`, `cột`...
* **Ví dụ thực tế:**
  - `=DValue("Đoạn 1: (2.5m + 3.2m) * 2 bên * 1.5m cao")` $\rightarrow$ **Kết quả:** `17.1`
  - `=DValue("Mố M1: 2*3.5*1.2m3 + [4.5 + 2.5]*2kg")` $\rightarrow$ **Kết quả:** `22.4`

---

#### 2. CÁC HÀM THIẾT KẾ KẾT CẤU ÁO ĐƯỜNG (KCAD)
##### 2.1. Tính Mô Đun Đàn Hồi Yêu Cầu ($E_{yc}$)
* **Cú pháp:** `=TinhEyc(Ne, [CapDuong], [DoTinCay])` hoặc `=TraEyc(...)`
* **Tham số:**
  - `Ne`: Số trục xe tiêu chuẩn 100 kN tích lũy trong thời hạn thiết kế (trục).
  - `CapDuong` *(tùy chọn)*:
    + `1` hoặc `"Cap cao A1"` *(mặc định)*: Đường cao tốc, cấp I, II, III.
    + `2` hoặc `"Cap cao A2"`: Đường cấp IV, V.
    + `3` hoặc `"Cap thap B1"`: Đường cấp VI, đường giao thông nông thôn.
  - `DoTinCay` *(tùy chọn)*: Độ tin cậy thiết kế ($0.98, 0.95, 0.90, 0.85$).
* **Công thức chuẩn TCCS 38:2022:**
  $$E_{yc} = 98.65 \cdot \left[\lg(N_e) - 1.79\right]$$
* **Ví dụ:** `=TinhEyc(1500000)` $\rightarrow$ **147.24 MPa**.

##### 2.2. Tra Hệ Số Hiệu Chỉnh Ứng Suất Kéo Uốn ($\beta$)
* **Cú pháp:** `=TraBeta(H_D)`
* **Tham số:** `H_D` = Tỷ số giữa tổng chiều dày các tầng bên trên đáy lớp kiểm toán với đường kính vệt bánh xe ($D = 33\text{ cm}$).
* **Ví dụ:** `=TraBeta(0.5)` $\rightarrow$ **1.034** (nội suy chính xác theo Bảng 3.6).

##### 2.3. Tra Toán Đồ Kogan Mô Đun Chung Hệ 2 Tầng ($E_{ch}/E_0$)
* **Cú pháp:** `=TraEch(H_D, E0_Etb)`
* **Tham số:**
  - `H_D`: Tỷ số chiều dày tương đương $h/D$.
  - `E0_Etb`: Tỷ số mô đun giữa lớp dưới và lớp trên ($E_0 / E_{tb}$ hoặc $E_m / E_{ch}$).
* **Ví dụ:** `=TraEch(1.2, 0.25)` $\rightarrow$ Tự động nội suy song tuyến tính trên toán đồ Hình 3.1.

##### 2.4. Tra Hệ Số Cường Độ Yêu Cầu ($K_{cđ}$)
* **Cú pháp:** `=TraKcd(DoTinCay, [CapDuong])`
* **Bảng tra chuẩn:**
  - $P = 0.98$ (98%) $\rightarrow$ $K_{cđ} = 1.15$
  - $P = 0.95$ (95%) $\rightarrow$ $K_{cđ} = 1.10$
  - $P = 0.90$ (90%) $\rightarrow$ $K_{cđ} = 1.05$
  - $P = 0.85$ (85%) $\rightarrow$ $K_{cđ} = 1.00$
* **Ví dụ:** `=TraKcd(0.95)` $\rightarrow$ **1.10**; `=TraKcd(98)` $\rightarrow$ **1.15**.

##### 2.5. Các Hàm Kiểm Toán Chống Trượt & Kéo Uốn
* `=TraTaxHD2(H1_D, E1_Echm)`: Tra ứng suất cắt trượt hoạt tính $\tau_{ax}/p$ với $H/D = 2$ (Toán đồ 3.2).
* `=TraTaxHD4(H1_D, E1_Echm)`: Tra ứng suất cắt trượt hoạt tính $\tau_{ax}/p$ với $H/D = 4$ (Toán đồ 3.3, 3.4).
* `=TraTav(H1_D, E1_Echm)`: Tra ứng suất cắt trượt do trọng lượng bản thân các lớp áo đường $\tau_{av}/p$.
* `=TinhUS_tangmat(H1_D, E1_Echm)`: Tra ứng suất kéo uốn đáy tầng mặt BTN $\sigma_{ku}/p$ (Toán đồ 3.5).
* `=TinhUS_tangmong(Hm_D, Em_E0)`: Tra ứng suất kéo uốn đáy tầng móng gia cố chất liên kết vô cơ $\sigma_{ku}/p$ (Toán đồ 3.6).
* `=tlookup(TableRange, Xu, Yu)`: Tra bảng ma trận 2 chiều trực tiếp trên các sheet toán đồ.
* `=interpolate(col_idx, x_val, data_range)`: Nội suy tuyến tính cột giá trị trong bảng dữ liệu.

---

#### 3. CÁC HÀM NỘI SUY TOÁN HỌC & KỸ THUẬT (Mod_NoiSuy)
##### 3.1. Nội Suy Tuyến Tính 2 Điểm Kẹp
* **Cú pháp:** `=Noisuy2b(X1, Y1, X2, Y2, X_can_tra)`
* **Công thức:**
  $$Y = Y_1 + \frac{Y_2 - Y_1}{X_2 - X_1} \cdot (X - X_1)$$
* **Ví dụ:** `=Noisuy2b(10, 100, 20, 250, 14)` $\rightarrow$ **160**.

##### 3.2. Nội Suy Dãy Số 1 Chiều (Tự Động Sắp Xếp)
* **Cú pháp:** `=NoisuyDay(Day_X, Day_Y, X_can_tra)`
* **Ưu điểm:** `Day_X` và `Day_Y` có thể là hàng ngang hoặc cột dọc, không cần sắp xếp trước, tự động tìm khoảng kẹp và nội suy/ngoại suy.
* **Ví dụ:** `=NoisuyDay(A2:A10, B2:B10, 15.5)`.

##### 3.3. Nội Suy Ma Trận 2 Chiều Song Tuyến Tính (Bilinear)
* **Cú pháp:** `=NoisuyMatran(Hang_X, Cot_Y, MaTran_GT, Xu, Yu)`
* **Tham số:**
  - `Hang_X`: Dãy mốc trục hoành (trục ngang).
  - `Cot_Y`: Dãy mốc trục tung (trục dọc).
  - `MaTran_GT`: Vùng dữ liệu các ô bên trong bảng.
  - `Xu`, `Yu`: Tọa độ 2 trục cần tra.
* **Ví dụ:** `=NoisuyMatran(B1:F1, A2:A8, B2:F8, 3.5, 12.8)`.

##### 3.4. Hàm Tính Diện Tích Tiết Diện Thép Tròn ($mm \rightarrow m^2$)
* **Cú pháp:** `=Dt(d)`
* **Công thức:** $Dt = \frac{\pi \cdot d^2}{4 \cdot 10^6}$ ($m^2$).
* **Ví dụ:** `=Dt(18)` $\rightarrow$ **0.0002545 $m^2$** (diện tích cây thép D18).

##### 3.5. Hàm Lấy Tên Vùng Ô
* **Cú pháp:** `=Name(Range)`
* **Ví dụ:** `=Name(B2:D10)` $\rightarrow$ Trả về chuỗi `"B2xD10"`.

---

#### 4. CÁC HÀM CHUYỂN ĐỔI BẢNG MÃ & ĐỔI KIỂU CHỮ TRONG Ô TÍNH
* `=TCVN3ToUnicode(chuỗi)`: Chuyển đổi chuỗi từ TCVN3 (.VnTime) sang Unicode dựng sẵn.
* `=VNIToUnicode(chuỗi)`: Chuyển đổi chuỗi từ VNI sang Unicode dựng sẵn.
* `=UnicodeToTCVN3(chuỗi)`: Chuyển chuỗi từ Unicode sang TCVN3.
* `=UnicodeToVNI(chuỗi)`: Chuyển chuỗi từ Unicode sang VNI.
* `=AutoToUnicode(chuỗi, [fontName])`: Tự động dò mã và chuyển sang Unicode.
* `=ChangeTextCase(chuỗi, caseType)`:
  - `caseType = 1`: Chuyển sang **IN HOA**.
  - `caseType = 2`: Chuyển sang **in thường**.
  - `caseType = 3`: Chuyển sang **Viết Hoa Đầu Từ**.

---

### PHẦN III. TÍNH NĂNG TỰ ĐỘNG & TIỆN ÍCH NỔI BẬT

1. **Gợi ý tham số thông minh (Excel-DNA IntelliSense):**
   - Khi bạn gõ `=TraBeta(` hoặc `=TinhEyc(`, Excel tự động hiện khung gợi ý tiếng Việt có dấu giải thích từng đối số ngay dưới con trỏ.
2. **Chèn lệnh rỗng `()` tiện dụng:**
   - Các nút tra cứu trên thanh Ribbon khi bấm sẽ chỉ chèn lệnh dạng rỗng `=TraBeta()` hoặc `=TraEch()`, để người dùng tự chọn ô tham chiếu, không điền sẵn số mẫu gây nhầm lẫn.
3. **Mẫu bảng tính đầy đủ sheet:**
   - Nút mở mẫu tạo ra file Excel mới chứa 100% các sheet (kết cấu, bảng tra, toán đồ, các phương án).
   - Phần tên công trình và họ tên người lập/kiểm tra được để trống để người dùng nhập thông tin dự án mới.
4. **An toàn 100% cho công thức & hình ảnh:**
   - Khi quét chuyển đổi font toàn bảng tính, hệ thống bỏ qua các ô chứa công thức (`HasFormula = True`) và hình ảnh/chữ ký scan.
