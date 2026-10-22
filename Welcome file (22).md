## PHẦN 1: Hoàn thiện tài liệu SRS bằng Heading Styles và Auto TOC

### 1. Dàn ý cấu trúc tiêu đề SRS 3 cấp

-   **1. Giới thiệu dự án**  `(Heading 1)`
    
    -   **1.1. Mục tiêu**  `(Heading 2)`
        
    -   **1.2. Đối tượng sử dụng**  `(Heading 2)`
        
-   **2. Chức năng hệ thống**  `(Heading 1)`
    
    -   **2.1. Quản lý mượn sách**  `(Heading 2)`
        
        -   **2.1.1. Luồng mượn sách qua thẻ sinh viên**  `(Heading 3)`
            
    -   **2.2. Quản lý trả sách**  `(Heading 2)`
        
-   **3. Yêu cầu bảo mật**  `(Heading 1)`
    

### 2. Thao tác gán nhãn Heading Styles

1.  Bôi đen từng dòng tiêu đề tương ứng.
    
2.  Sử dụng phím tắt nhanh:
    
    -   **Heading 1**: Nhấn `Ctrl + Alt + 1` cho các mục:  `1.`,  `2.`,  `3.`
        
    -   **Heading 2**: Nhấn `Ctrl + Alt + 2` cho các mục:  `1.1.`,  `1.2.`,  `2.1.`,  `2.2.`
        
    -   **Heading 3**: Nhấn `Ctrl + Alt + 3` cho mục:  `2.1.1.`
        

### 3. Nội dung mô tả yêu cầu (Nội dung mẫu chèn vào SRS)

#### Dưới mục `2. Chức năng hệ thống` (Nêu 2 Yêu cầu chức năng - Functional Requirements):

1.  Hệ thống phải cho phép thủ thư tra cứu thông tin sách, kiểm tra tình trạng khả dụng và lập phiếu mượn cho sinh viên theo mã thẻ.
    
2.  Hệ thống phải tự động tính toán số ngày mượn, cảnh báo quá hạn và cập nhật trạng thái phiếu mượn ngay khi sách được hoàn trả.
    

#### Dưới mục `3. Yêu cầu bảo mật` (Nêu 2 Yêu cầu phi chức năng - Non-functional Requirements):

1.  Tất cả mật khẩu tài khoản người dùng và thông tin giao dịch phải được mã hóa chuẩn AES-256 trước khi lưu trữ vào CSDL.
    
2.  Hệ thống phải phân quyền truy cập chặt chẽ (RBAC) theo từng vai trò (Sinh viên, Thủ thư, Quản trị viên) và tự động đăng xuất sau 15 phút không thao tác.
    

### 4. Quy trình chèn Auto TOC và Cập nhật số trang khi lệch trang

#### Bước 1: Chèn Mục lục tự động

-   Đặt con trỏ ở trang đầu tiên của tài liệu.
    
-   Vào thẻ **References**  $\rightarrow$ chọn **Table of Contents**  $\rightarrow$ chọn mẫu **Automatic Table 1** (hoặc **Automatic Table 2**).
    

#### Bước 2: Mô tả thao tác Cập nhật số trang (khi đẩy Mục 3 sang trang sau)

1.  Sau khi chèn thêm văn bản làm mục **3. Yêu cầu bảo mật** bị đẩy từ trang 2 sang trang 3 (hoặc trang tiếp theo).
    
2.  Click chuột vào bảng Mục lục vừa tạo.
    
3.  Nhấn phím **`F9`** (hoặc chọn nút **Update Table...** ở góc trên bảng mục lục).
    
4.  Chọn tùy chọn **Update entire table** (Cập nhật toàn bộ bảng) $\rightarrow$ Nhấn **OK**.
    

**Kết quả minh họa bảng Auto TOC sau khi Update:**

Plaintext

```
MỤC LỤC
1. Giới thiệu dự án ..................................................... 2
   1.1. Mục tiêu ........................................................ 2
   1.2. Đối tượng sử dụng ............................................... 2
2. Chức năng hệ thống ................................................... 2
   2.1. Quản lý mượn sách ............................................... 2
       2.1.1. Luồng mượn sách qua thẻ sinh viên ........................ 2
   2.2. Quản lý trả sách ................................................ 2
3. Yêu cầu bảo mật ...................................................... 3

```

## PHẦN 2: Chuẩn hóa dữ liệu Excel và AutoFilter

### 1. Phân tích 3 lỗi vi phạm 5 Quy tắc Vàng dữ liệu

**Lỗi**

**Vị trí (Hàng/Ô)**

**Quy tắc bị vi phạm**

**1**

Hàng 1 (Ô `A1:E1`)

**Merge ô (Gộp ô):** Tiêu đề gộp ô nằm đè lên vùng dữ liệu làm hỏng cấu trúc cột và gây lỗi khi lọc/sắp xếp.

**2**

Hàng 5

**Hàng trống:** Đoạn ngắt giữa dữ liệu làm đứt gãy vùng dữ liệu liên tục (`CurrentRegion`), làm Excel không nhận diện đúng toàn bộ bảng.

**3**

Hàng 6

**Dòng tổng cộng nằm trong vùng dữ liệu:** Dòng tính tổng/thống kê chèn ngang giữa các dòng chi tiết phá vỡ cấu trúc bản ghi và làm sai lệch phép toán tổng.

### 2. Bảng dữ liệu chuẩn hóa (Đúng tư duy CSDL)

Bảng dữ liệu sau khi loại bỏ tiêu đề gộp ô, xóa hàng trống và chuyển dòng Tổng cộng ra ngoài:

**MSSV**

**Họ Tên**

**Mã Sách**

**Ngày Mượn**

**Tình Trạng**

SV001

Nguyễn Văn An

S001

01/10/2026

Đang mượn

SV002

Trần Thị Bình

S002

03/10/2026

Quá hạn

SV003

Lê Văn Cường

S003

10/10/2026

Đã trả

SV004

Phạm Thu Hà

S099

12/10/2026

Đang mượn

_(Ghi chú: Tiêu đề "DANH SÁCH MƯỢN SÁCH THÁNG 10" được đưa lên làm tên Sheet hoặc đặt ở dòng riêng biệt phía trên có khoảng cách hẳn với bảng)._

### 3. Phân biệt Record và Field

-   **Record (Bản ghi / Hàng):** Là một tập hợp các thông tin cụ thể của một đối tượng duy nhất theo hàng ngang.
    
    -   _Ví dụ:_ Dòng chứa thông tin của sinh viên `SV001` | `Nguyễn Văn An` | `S001` | `01/10/2026` | `Đang mượn` cấu thành 1 Record.
        
-   **Field (Trường dữ liệu / Cột):** Là thuộc tính chứa một loại thông tin thống nhất theo chiều dọc.
    
    -   _Ví dụ:_ Cột **Mã Sách** (chứa các giá trị `S001`,  `S002`,  `S003`,  `S099`) là 1 Field.
        

### 4. Thao tác bật AutoFilter và lọc sinh viên "Quá hạn"

1.  Đặt con trỏ chuột vào một ô bất kỳ trong bảng dữ liệu (ví dụ ô `A1`).
    
2.  Nhấn tổ hợp phím tắt **`Ctrl + Shift + L`** để bật tính năng AutoFilter (xuất hiện mũi tên góc tại từng ô tiêu đề cột).
    
3.  Click vào mũi tên tại tiêu đề cột **Tình Trạng** (Cột `E`).
    
4.  Bỏ chọn _(Select All)_, tích chọn duy nhất vào ô **Quá hạn**.
    
5.  Nhấn **OK**. (Kết quả màn hình sẽ chỉ hiển thị bản ghi của sinh viên `SV002 - Trần Thị Bình`).
    

## PHẦN 3: Nối dữ liệu bằng VLOOKUP

### 1. Thêm cột "Tên Sách" và viết công thức VLOOKUP

Giả sử bảng `PhieuMuon` bắt đầu tiêu đề từ ô `A1:F1`, cột Mã Sách nằm ở cột `C` (từ `C2`). Sheet `DanhMucSach` có bảng dữ liệu tra cứu nằm ở vùng `A2:C4` (Cột A: Mã Sách, Cột B: Tên Sách, Cột C: Thể Loạn).

Công thức VLOOKUP cơ bản điền tại ô `E2` (cột Tên Sách) của Sheet `PhieuMuon`:

Excel

```
=VLOOKUP(C2, DanhMucSach!$A$2:$C$100, 2, FALSE)

```

### 2. Phân tích lỗi đối với Mã Sách `S099` (Hàng 8 trong dữ liệu gốc)

-   **Kết quả trả về:** Lỗi **`#N/A`** (Not Available).
    
-   **Nguyên nhân:** Hàm VLOOKUP thực hiện tìm kiếm chính xác tuyệt đối (do tham số 4 là `FALSE`), tuy nhiên giá trị mã `S099` hoàn toàn **không tồn tại** trong cột Mã Sách của sheet `DanhMucSach`.
    

### 3. Công thức hoàn chỉnh kết hợp `IFERROR` và `TRIM`

Để xử lý triệt để lỗi `#N/A` và phòng tránh bẫy khoảng trắng thừa trong mã sách, công thức chuẩn hóa đầy đủ là:

Excel

```
=IFERROR(VLOOKUP(TRIM(C2), DanhMucSach!$A$2:$C$100, 2, FALSE), "Không tìm thấy sách")

```

**Bảng kết quả tra cứu hoàn chỉnh:**

**MSSV**

**Họ Tên**

**Mã Sách**

**Tên Sách (Cột mới)**

**Ngày Mượn**

**Tình Trạng**

SV001

Nguyễn Văn An

S001

Lập Trình C Căn Bản

01/10/2026

Đang mượn

SV002

Trần Thị Bình

S002

Giải Tích 1

03/10/2026

Quá hạn

SV003

Lê Văn Cường

S003

Kỹ Năng Mềm

10/10/2026

Đã trả

SV004

Phạm Thu Hà

S099

**Không tìm thấy sách**

12/10/2026

Đang mượn

## PHẦN 4: Slide Báo cáo Tiến độ (Rule 10/20/30)

### 1. Bố cục nội dung 1 Slide báo cáo (Slide Content Design)

#### A. Header (Assertion Title - Tiêu đề kết luận):

> **"Tỷ lệ phiếu mượn đến hạn đạt 75% nhưng phát hiện 25% dữ liệu mã sách lỗi cần xử lý"**

#### B. Content Area (Khu vực nội dung chính):

-   **Key Metric (Con số nổi bật - Callout):**
    
    -   **25%**  _(1/4 phiếu mượn bị Quá hạn)_
        
    -   **1/4**  _(Mã sách không tồn tại trong danh mục tra cứu: `S099`)_
        
-   **Biểu đồ lựa chọn:**  **Biểu đồ Tròn / Bánh đai ốc (Donut Chart)**
    
    -   **Thống kê phân bổ tình trạng phiếu mượn:**
        
        -   _Đang mượn:_ 50% (2 phiếu)
            
        -   _Đã trả:_ 25% (1 phiếu)
            
        -   _Quá hạn:_ 25% (1 phiếu)
            
    -   **Lý do chọn Donut Chart:** Dữ liệu tình trạng là các phần cấu thành một tổng thể (100% danh sách phiếu mượn tháng 10) với số lượng phân loại ít (3 trạng thái). Biểu đồ Donut giúp thể hiện rõ nét tỷ lệ phần trăm tương quan, đồng thời khoảng trống ở giữa biểu đồ dùng để đặt con số nhấn mạnh (ví dụ: Tổng 4 phiếu).
        

### 2. Áp dụng quy tắc 10/20/30 (Guy Kawasaki)

-   **Tổng thể quy tắc:** 10 Slide – 20 Phút thuyết trình – Cỡ chữ tối thiểu 30pt.
    
-   **Định vị slide trong tổng thể bài báo cáo:**
    
    -   Slide báo cáo tiến độ này đóng vai trò là **1 slide** nằm trong tổng số **10 slide** của bài thuyết trình báo cáo giữa kỳ/tiến độ dự án.
        
    -   Mức thời gian phân bổ dự kiến cho slide này là **2 phút** (trên tổng số **20 phút** trình bày toàn bộ bài báo cáo).
        

## PHẦN 5: Hướng dẫn đóng gói và Quy định nộp bài

1.  **Cấu trúc thư mục nộp bài:** Tạo thư mục trên máy tính/Google Drive theo đúng cú pháp:`[D21CQCN01]_[ThietKeHeThong]_[Session01]_[BaiTapLon]`
    
2.  **Các file thành phần bên trong thư mục:**
    
    -   `SRS_ThuVien.docx` (Chứa tài liệu SRS đã có Auto TOC và định dạng Heading).
        
    -   `PhieuMuon_ThuVien.xlsx` (Chứa 2 sheet `PhieuMuon` đã xóa lỗi/dùng VLOOKUP+IFERROR và sheet `DanhMucSach`).
        
    -   `SlideBaoCao.pptx` (Chứa slide báo cáo chuẩn 1 Slide - 1 Thông điệp).
        
3.  **Hình thức nộp:** Nộp đường dẫn chia sẻ (Link Public/View) của thư mục Google Drive lên Portal học tập trước deadline.
