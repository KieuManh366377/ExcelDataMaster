# ⚡ **ExcelDataMaster - Thư Viện Siêu Tốc Truy Xuất Dữ Liệu Excel Qua Bộ Nhớ**

> “Không cần mở Excel, không cần OLE Automation. Truy xuất dữ liệu từ nhiều file Excel chỉ trong ít giây – với sức mạnh của Memory DB.”

---

## 📌 **Tổng Quan**

**ExcelDataMaster** là thư viện cấp hệ thống (Windows DLL), được thiết kế chuyên biệt để **truy xuất, quản lý và phân tích dữ liệu từ file Excel (.xlsx)** một cách **nhanh chóng, an toàn và hoàn toàn độc lập với giao diện Excel**.

Thư viện sử dụng **cơ chế bộ nhớ RAM (Memory Database)** để tải và xử lý dữ liệu — giúp bạn thực hiện các truy vấn SQL trên dữ liệu Excel nhanh hơn gấp nhiều lần so với cách truyền thống.

---

## 🚀 **Điểm Nổi Bật**

### ✅ **1. Siêu tốc với Memory DB**

* Dữ liệu từ Excel được load vào RAM như một bảng trong cơ sở dữ liệu SQLite hoặc FireDAC MemTable.
* Truy vấn SQL ngay lập tức mà không cần mở Excel.

### ✅ **2. Dễ dàng tích hợp**

* Giao tiếp qua DLL chuẩn (`stdcall`) tương thích VBA, VB.NET, C#, Python, Delphi,...
* Dùng được trong macro Excel hoặc ứng dụng desktop.

### ✅ **3. Hệ thống quản lý Session mạnh mẽ**

* Mỗi truy vấn tạo một session độc lập: dễ theo dõi, dọn dẹp và tái sử dụng.
* Hỗ trợ session timeout, dọn dẹp nền (AutoCleaner), tracking SQL, quản lý theo handle.

### ✅ **4. Truy vấn song song (Parallel Query)**

* Load đồng thời hàng chục file Excel khác nhau chỉ với một hàm `ExcelLoadBatchData(...)`.

### ✅ **5. Đầy đủ API nâng cao**

* Xem danh sách session đang hoạt động
* Lọc session lỗi, session hết hạn
* Lấy thống kê, nhật ký, thời gian truy cập gần nhất
* Trích tiêu đề cột (column names)
* Kết xuất JSON danh sách SQL theo từng file

---

## 🛠️ **Cách Sử Dụng Trong VBA**

```vb
' Khởi tạo session từ Excel
Dim h As LongPtr
h = MemSessionOpenFromExcel("C:\Data.xlsx", "SELECT * FROM [Sheet1$]")

' Lấy tiêu đề cột
Dim cols As Variant
cols = MemSessionGetColumnNames(h)

' Lọc dữ liệu theo từ khóa
Dim results As Variant
results = ExcelLoadBatchData(Array("C:\Data.xlsx"), Array("SELECT * FROM [Sheet1$] WHERE [Name] LIKE '%John%'"), Array(True))
```

---

## 📈 **Ứng Dụng Thực Tế**

* ✅ Xây dựng **form lọc động trong Excel** (Auto Filter Form)
* ✅ Tạo **báo cáo tổng hợp từ nhiều file Excel** mà không cần mở file
* ✅ Phân tích dữ liệu trong ứng dụng C# hoặc Python
* ✅ Tự động kiểm tra dữ liệu lỗi, session expired, log lại truy vấn

---

## 🔒 **An Toàn - Không can thiệp file gốc**

* Chỉ đọc dữ liệu – không ghi, không làm hỏng Excel.
* Xử lý Unicode tiếng Việt đầy đủ (UTF-16 / UTF-8)

## 🌟 **Tải Về Và Trải Nghiệm**

