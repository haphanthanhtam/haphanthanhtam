# Iphanmem website

Website tĩnh (HTML/CSS/JS thuần, không cần build) cho iphanmem.com. Trọng tâm là trang **Tải xuống**; không hiển thị giá (khách nhắn Zalo để được tư vấn).

## Triển khai lên Cloudflare
1. Đưa thư mục này lên một repository GitHub riêng (ví dụ `iphanmem-website`).
2. Cloudflare → Workers & Pages → Create → kết nối GitHub → chọn repo.
3. Build command: để trống. Output directory: `/` (hoặc `iphanmem-website` nếu giữ nguyên cấu trúc thư mục).
4. Gắn tên miền `iphanmem.com` trong phần Custom domains của dự án.

## Đưa bộ cài lên (quan trọng)
Bộ cài Windows thường lớn hơn giới hạn dung lượng mỗi file của Cloudflare Pages, và không nên commit vào Git.
Hãy đặt bộ cài trên Cloudflare R2 (bucket riêng, gắn tên miền `dl.iphanmem.com`), rồi sửa `releases.json`:
- `sample`: đổi thành `false` khi dùng dữ liệu thật.
- `baseUrl`: địa chỉ bucket, ví dụ `https://dl.iphanmem.com`.
- `products.<id>`: `version`, `date`, `size`, `file` (đường dẫn trong bucket), `sha256` (lấy bằng `certutil -hashfile <file> SHA256`), `notes`.

Trang Tải xuống tự đọc file này, không cần sửa HTML mỗi lần phát hành. Nên ký số (code signing) bộ cài trước khi phát hành rộng.

## Việc cần làm trước khi công bố
- [ ] Xác nhận yêu cầu hệ thống với đội phát triển; thay khối "video/ảnh/benchmark" bằng nội dung thật.
- [ ] Bổ sung thông tin pháp nhân, điều khoản, chính sách bảo mật.
- [ ] Rà lại các câu về offline, thiết bị, dữ liệu, dùng thử cho đúng với phần mềm thật.
- [ ] Bật `Auto renew` cho tên miền.

## Ghi chú nội bộ về bán hàng
- Website không hiển thị giá. Khách liên hệ qua Zalo 0335 156 602 hoặc email haphanthanhtam.work@gmail.com.
- Chính sách license (1 license = 1 máy) chỉ trao đổi khi tư vấn, không ghi trên web.
- Giá và hình thức mua (theo năm, vĩnh viễn) tự trao đổi với từng khách qua Zalo.

## Cấu trúc
`index, loading, cutting, erp, download, contact .html` · `releases.json` · `style.css` · `script.js` · `sitemap.xml` · `robots.txt` · `_headers`
