# QLVT Tool V2

QLVT Tool là ứng dụng quản lý vật tư đơn giản, cho phép bạn:
- Import danh sách vật tư từ file Excel
- Tìm kiếm vật tư theo mã hoặc tên
- Sao chép mã vật tư nhanh chóng
- Chỉnh sửa thông tin vật tư
- Thay đổi thứ tự vật tư bằng kéo thả

## Cài đặt

### 1. Chuẩn bị môi trường

#### 1.1 Cài đặt Python
Tải và cài đặt Python phiên bản 3.8 hoặc cao hơn từ [trang chủ Python](https://www.python.org/downloads/).

Khi cài đặt, nhớ đánh dấu tùy chọn "Add Python to PATH".

#### 1.2 Tạo môi trường ảo (venv)

Mở Command Prompt hoặc PowerShell và chạy các lệnh sau:

```
cd "d:\Python\QLVT Tool V2"
python -m venv venv
venv\Scripts\activate
```

#### 1.3 Cài đặt các thư viện cần thiết

Sau khi kích hoạt môi trường ảo, cài đặt các thư viện:

```
pip install -r requirements.txt
```

### 2. Chạy ứng dụng

```
python main.py
```

## Hướng dẫn sử dụng

### Import dữ liệu từ Excel
1. Nhấn nút "Import Excel"
2. Chọn file Excel có chứa cột "Mã vật tư" và "Tên vật tư"
3. Dữ liệu sẽ được tải và hiển thị trong ứng dụng

### Tìm kiếm vật tư
- Nhập từ khóa vào ô tìm kiếm
- Ứng dụng sẽ lọc và hiển thị các vật tư có mã hoặc tên chứa từ khóa

### Sao chép mã vật tư
- Nhấn nút "Copy" bên cạnh vật tư muốn sao chép
- Mã vật tư sẽ được sao chép vào clipboard

### Chỉnh sửa vật tư
- Double-click lên tên vật tư để mở hộp thoại chỉnh sửa
- Nhập thông tin mới và nhấn "Lưu"

### Thay đổi thứ tự vật tư
- Kéo và thả (drag & drop) vật tư để thay đổi vị trí

## Cấu trúc file Excel
File Excel cần có ít nhất 2 cột:
1. "Mã VT" - chứa mã của vật tư
2. "Tên VT" - chứa tên của vật tư

Ví dụ:

| Mã VT    | Tên VT                 |
|-----------|------------------------|
| VT001     | Thép tấm 2mm           |
| VT002     | Ống đồng phi 15        |
| VT003     | Dây điện 2x1.5mm       |
