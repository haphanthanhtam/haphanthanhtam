# Iphanmem website

Website tĩnh (HTML/CSS/JS thuần, không cần build) cho iphanmem.com. Trọng tâm là trang **Tải xuống**; giá hiển thị dạng "Giá: IB" (khách nhắn Zalo/email để nhận báo giá).

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
- [ ] Bật email `lienhe@iphanmem.com` (Cloudflare Email Routing) hoặc đổi `EMAIL` trong `gen.py`/nội dung.
- [ ] Điền số Zalo (tìm `[cập nhật số Zalo]`).
- [ ] Xác nhận yêu cầu hệ thống với đội phát triển; thay khối "video/ảnh/benchmark" bằng nội dung thật.
- [ ] Bổ sung thông tin pháp nhân, điều khoản, chính sách bảo mật.
- [ ] Rà lại các câu về offline, thiết bị, dữ liệu, dùng thử cho đúng với phần mềm thật.
- [ ] Bật `Auto renew` cho tên miền.

## Bảng giá nội bộ (đề xuất, KHÔNG đăng công khai)
Dùng khi trả lời IB. Chưa phải giá đã chốt, chưa đối chiếu với đối thủ hay chi phí.

| | Theo năm | Vĩnh viễn |
|---|---|---|
| Loading 3D | 14.900.000 đ | 39.900.000 đ |
| Cutting | 9.900.000 đ | 24.900.000 đ |
| Combo cả hai | giảm 15% | giảm 15% |
| ERP | báo giá theo dự án | |

- 1 license = 1 người dùng, tối đa 2 thiết bị. Chưa gồm VAT 10%.
- Từ 5 license giảm 10%, từ 10 license giảm 20%.
- Vĩnh viễn: kèm cập nhật và hỗ trợ 12 tháng đầu; từ năm 2 gói cập nhật tùy chọn 20% giá vĩnh viễn mỗi năm.

## Cấu trúc
`index, loading, cutting, erp, pricing, download, contact .html` · `releases.json` · `style.css` · `script.js` · `sitemap.xml` · `robots.txt` · `_headers`

Form liên hệ hiện mở ứng dụng email của khách (mailto). Khi cần lưu yêu cầu trên máy chủ, thêm Pages Functions hoặc dịch vụ form.
