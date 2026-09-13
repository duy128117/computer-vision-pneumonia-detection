# Kiểm tra thực tế

- Python: 3.11.9.
- PyTorch: 2.6.0+cu124; torchvision: 0.21.0+cu124.
- CUDA nhận NVIDIA GeForce RTX 2050, 4 GB.
- `python -m pytest tests -q --basetemp=.tmp/pytest`: **8 passed in 12.77s**.
- Sau bổ sung mixed precision: **8 passed in 5.34s**.
- Streamlit AppTest: trạng thái chưa có checkpoint hiển thị warning, không exception.
- Kiểm tra cú pháp Python và các code cell notebook: đạt.
- Dataset đã tải và trích xuất: 5.856 JPEG; CRC32 và kích thước từng member ZIP được xác minh; SHA256 ảnh lưu trong `data/source.json`.
- Phân bố gốc: train NORMAL 1.341 / PNEUMONIA 3.875; val 8 / 8; test 234 / 390.
- Split dùng thực nghiệm: train 1.078 / 2.737; validation 270 / 668; test 231 / 387. Tổng 5.371 ảnh, loại 485 ảnh theo kiểm tra trùng và bệnh nhân; mọi nhóm bệnh nhân/hash pixel đều tách biệt giữa các split.
- Lý do loại: 457 ảnh development trùng bệnh nhân với test, 22 ảnh trùng pixel trong development, 6 ảnh trùng pixel trong test (`results/excluded_images.csv`).

Các mục trên là kiểm tra kỹ thuật và số đếm dữ liệu, không phải metric đánh giá mô hình. Kết quả train/test chỉ có giá trị khi pipeline sinh checkpoint, history và prediction tương ứng. Xem `logs/run_*.json` và `results/report.md` để xác nhận trạng thái thực nghiệm.
