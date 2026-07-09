# Sổ tay vận hành Shopee & TikTok Shop

Dự án này là bộ công cụ tích hợp Landing Page **Cẩm nang Vận hành Bán hàng Shopee & TikTok Shop** (bao gồm tài liệu hướng dẫn, các bảng tính lợi nhuận tự động và trợ lý tranh chấp).

## 📄 File chạy chính
*   **`index.html`** (Trang chủ chính của Landing Page).

## 🚀 Cách publish lên GitHub Pages
1.  Giải nén tệp tin `github_pages_publish_sotayvanhanh.zip`.
2.  Tải toàn bộ các tệp tin và thư mục bên trong lên một kho chứa (Repository) trên GitHub của bạn.
    *   *Đảm bảo file `index.html` nằm ngay tại thư mục gốc (root) của Repository.*
3.  Truy cập vào phần **Settings** > **Pages** trên GitHub của Repository đó.
4.  Tại mục **Build and deployment** > **Source**, chọn **Deploy from a branch**.
5.  Chọn branch chính (ví dụ: `main` hoặc `master`) và chọn thư mục gốc `/ (root)`. Sau đó nhấn **Save**.
6.  Chờ 1-2 phút, GitHub sẽ cung cấp một đường dẫn web chạy trực tiếp (ví dụ: `https://<username>.github.io/<repo-name>/`).

## ⚠️ Lưu ý quan trọng
*   **Cấu trúc thư mục:** Không được thay đổi cấu trúc thư mục của các asset hình ảnh (`shopee_flash_sale_images/` và `scratch/gdoc_html/images/`), vì trang `index.html` sử dụng đường dẫn tương đối (relative paths) để trỏ đến các ảnh này.
*   **`.nojekyll`:** Tệp `.nojekyll` được đặt tại thư mục gốc giúp GitHub Pages không xử lý sai hoặc bỏ qua các tệp tin tĩnh/thư mục đặc biệt.
*   **Biên dịch lại:** Nếu muốn tự cập nhật nội dung từ các file Markdown trong thư mục `sources/` hoặc dữ liệu Excel, bạn chỉnh sửa nguồn và chạy lệnh:
    ```bash
    python build_unified_handbook_new.py
    ```
