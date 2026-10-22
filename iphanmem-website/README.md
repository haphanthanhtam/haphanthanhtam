# Iphanmem website (iphanmem.com)

Website giới thiệu và nhận bản dùng thử cho iLoading, iCutting và ERP.
HTML/CSS/JS thuần, **không cần build**, không có framework hay thư viện ngoài (chỉ phông chữ Google Fonts).
Không chứa mã nguồn hay thuật toán của các phần mềm Loading, Cutting, ERP.

## Cấu trúc
| File | Nội dung |
|---|---|
| `index.html` | Trang chủ |
| `loading.html`, `cutting.html`, `erp.html` | Trang sản phẩm |
| `download.html` + `releases.json` | Trang tải về, đọc thông tin phiên bản từ `releases.json` |
| `contact.html` | Liên hệ qua Zalo / email |
| `404.html` | Trang lỗi 404 (Cloudflare Pages tự dùng) |
| `style.css`, `script.js` | Giao diện và mã JS (menu di động, trang tải về, nút chép) |
| `_headers`, `robots.txt`, `sitemap.xml`, `favicon.svg` | Cấu hình Cloudflare, SEO, biểu tượng |

Header và footer lặp lại ở từng trang HTML. Khi đổi menu, số Zalo hay email, sửa đồng loạt ở mọi file `.html`
(tìm `0335 156 602` và `info@iphanmem.com`).

## Chạy thử trên máy
```
cd iphanmem-website
python3 -m http.server 8000      # rồi mở http://localhost:8000
```
(Phải chạy qua máy chủ, không mở thẳng file, vì trang tải về dùng `fetch('releases.json')`.)

## Triển khai và cập nhật tự động (Cloudflare Pages + GitHub)
Cấu hình:
- Project: `iphanmem` (Workers & Pages → Create → Pages → Connect to Git)
- Repository: `haphanthanhtam/haphanthanhtam`
- Production branch: `claude/software-marketplace-ecosystem-5eh2me` (có thể đổi sang `main` sau khi gộp)
- Root directory: `iphanmem-website`
- Build command: để trống. Build output directory: `/` (nếu giao diện không nhận, thử để trống hoặc `.`)

Sau khi kết nối, mỗi lần push lên production branch, Cloudflare tự đăng lại website. Các nhánh khác tạo bản xem trước (preview) riêng.

### Tên miền
1. Dự án Pages → Custom domains → Set up a domain → `iphanmem.com` (domain đã ở Cloudflare nên DNS tự thêm).
2. Thêm `www.iphanmem.com` cũng ở Custom domains, rồi tạo **Redirect Rule** (Rules → Redirect Rules): `www.iphanmem.com/*` → `https://iphanmem.com/${1}` (301). Cách này không chạm tới bản ghi email/MX/TXT.
3. Chờ chứng chỉ SSL chuyển sang Active, rồi kiểm tra `https://iphanmem.com` và `https://www.iphanmem.com`.

### Rollback khi bản mới bị lỗi
- Cách nhanh: Dự án Pages → Deployments → chọn bản đăng tốt trước đó → Rollback to this deployment.
- Cách trong Git: `git revert <commit>` rồi push (không dùng force push).

## Bộ cài (khi đã có)
Trang tải về hiện hiển thị "Sắp phát hành" vì chưa có bộ cài thật. Khi có:
1. Ký số bộ cài (code signing) để Windows ít cảnh báo hơn.
2. Tạo bucket Cloudflare R2, gắn tên miền `dl.iphanmem.com`, tải bộ cài lên (không commit file .exe vào Git; Pages cũng giới hạn dung lượng mỗi file).
3. Lấy SHA-256: `certutil -hashfile <file> SHA256`.
4. Sửa `releases.json`:
```json
{
  "baseUrl": "https://dl.iphanmem.com",
  "products": {
    "loading": { "available": true, "version": "1.0.0", "date": "dd/mm/yyyy", "size": "xxx MB",
                 "channel": "Bản ổn định", "file": "loading/Ten-File-Setup.exe",
                 "sha256": "<mã sha256 thật>", "notes": ["Nội dung thật của bản phát hành"] },
    "cutting": { "available": false }
  }
}
```
Chỉ đặt `available: true` khi file thật đã tồn tại và tải được. Chưa kiểm tra giá, giới hạn miễn phí của R2: xem trang giá hiện hành của Cloudflare trước khi bật.

## Việc cần làm
- [ ] Tạo email theo tên miền (Cloudflare Email Routing) nếu không muốn dùng Gmail công khai.
- [ ] Xác nhận tính năng, yêu cầu cấu hình tối thiểu, video/ảnh giao diện thật cho từng trang sản phẩm.
- [ ] Bổ sung thông tin pháp nhân, điều khoản, chính sách bảo mật (cần nội dung do chủ sở hữu/luật sư xác nhận).
- [ ] Bật Auto renew cho tên miền trong Cloudflare → Domains.
- [ ] Thêm ảnh chia sẻ (og:image) và Google Search Console / Analytics nếu cần.

## Ghi chú nội bộ
Website không hiển thị giá. Khách liên hệ qua Zalo. Chính sách license (1 license = 1 máy) chỉ trao đổi khi tư vấn.
