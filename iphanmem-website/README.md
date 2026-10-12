# iPhanMem website

Website tĩnh (HTML/CSS/JS thuần, không cần build) cho iphanmem.com.

## Triển khai lên Cloudflare
1. Đưa thư mục này lên một repository GitHub riêng (ví dụ `iphanmem-website`).
2. Cloudflare → Workers & Pages → Create → kết nối GitHub → chọn repo.
3. Build command: để trống. Output directory: `/` (hoặc `iphanmem-website` nếu giữ nguyên cấu trúc thư mục).
4. Gắn tên miền `iphanmem.com` trong phần Custom domains của dự án.

## Việc cần làm trước khi công bố
- [ ] Bật email `lienhe@iphanmem.com` (Cloudflare Email Routing) hoặc đổi `EMAIL` trong nội dung.
- [ ] Điền số Zalo (tìm `[cập nhật số Zalo]`).
- [ ] Xác nhận giá, mức giảm theo số license, chính sách vĩnh viễn và phí cập nhật 20%/năm.
- [ ] Xác nhận yêu cầu hệ thống với đội phát triển, thay khối "video/ảnh/benchmark" bằng nội dung thật.
- [ ] Bổ sung thông tin pháp nhân, điều khoản, chính sách bảo mật.
- [ ] Rà lại các câu về offline, thiết bị, dữ liệu cho đúng với phần mềm thật.

## Cấu trúc
`index, loading, cutting, erp, pricing, download, contact .html` · `style.css` · `script.js` · `sitemap.xml` · `robots.txt` · `_headers`

Form liên hệ hiện mở ứng dụng email của khách (mailto). Khi cần lưu yêu cầu trên máy chủ, thêm Pages Functions hoặc dịch vụ form.
