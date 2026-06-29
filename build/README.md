# Hướng dẫn: Tự động lọc data mỗi khi có patch mới

## Vấn đề
File JSON export từ game (Skill.json, SkillEffect.json...) chứa rất nhiều field
debug/engine không liên quan tới việc hiển thị skill, khiến người dùng phải tải
~160MB mỗi lần mở trang. Giải pháp: lọc bớt còn ~22MB (giảm 86%).

## Cách hoạt động (đã setup sẵn, bạn không cần làm gì thêm sau khi merge)

1. Bạn copy patch mới của game vào `data/<tên_patch>/` như bình thường,
   thêm entry vào `data/versions.json`, rồi `git push` lên nhánh `main`.
2. GitHub Action (`.github/workflows/build-data-min.yml`) tự động:
   - Phát hiện có thay đổi trong `data/**`
   - Chạy `build/build_all.py` để lọc TẤT CẢ các patch tìm thấy trong `data/`
   - Tự commit kết quả vào `data_min/` và push ngược lại vào `main`
3. Vercel tự deploy lại bản mới nhất (vì nó theo dõi nhánh `main`).

Bạn không cần chạy lệnh gì cả — chỉ cần push patch game mới như bình thường.

## Cấu trúc file liên quan

```
.github/workflows/build-data-min.yml   <- Định nghĩa GitHub Action
build/minify_data.py                   <- Lọc 1 thư mục patch (input, output)
build/build_all.py                     <- Tự quét toàn bộ data/, gọi minify_data.py cho từng patch
data/                                   <- Dữ liệu GỐC từ game (giữ nguyên, đừng xoá)
data_min/                               <- Dữ liệu ĐÃ LỌC (script.js/compare.js đọc từ đây)
```

## Nếu muốn chạy thủ công (không bắt buộc)

Trên máy của bạn (cần Python 3):
```bash
python3 build/build_all.py
```
Lệnh này tự quét toàn bộ `data/`, build lại toàn bộ `data_min/`.

Hoặc chạy chạy tay trên GitHub: vào tab **Actions** trong repo, chọn workflow
"Tự động lọc data game", bấm **Run workflow**.

## Lưu ý quan trọng

- Action có cơ chế chống lặp vô hạn: commit do bot tạo ra sẽ có
  `[skip-data-build]` trong message, nên không tự trigger lại Action.
- Nếu bạn thêm field mới vào `script.js`/`compare.js` mà cần đọc thêm dữ
  liệu chưa có trong `data_min/`, phải cập nhật danh sách `KEEP_FIELDS` ở
  đầu file `build/minify_data.py` rồi push lại (hoặc chạy `workflow_dispatch`
  để build lại ngay không cần đổi file `data/`).
