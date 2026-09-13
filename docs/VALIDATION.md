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

## Kết quả hoàn tất ngày 2026-09-13

- ResNet50, checkpoint epoch 8: Accuracy 0.9061; Precision 0.8730; Sensitivity 0.9948; Specificity 0.7576; F1 0.9300; AUC 0.9679; TN/FP/FN/TP = 175/56/2/385.
- DenseNet121, checkpoint epoch 8: Accuracy 0.8657; Precision 0.8248; Sensitivity 0.9974; Specificity 0.6450; F1 0.9029; AUC 0.9578; TN/FP/FN/TP = 149/82/1/386.
- Chọn ResNet50 theo AUC validation của checkpoint, sau đó sensitivity và specificity; test không tham gia chọn model.
- Kiểm toán cuối: metric được tính lại từ toàn bộ 618 dòng dự đoán test/model và trùng với JSON; checkpoint khớp manifest/history; 27 slide hợp lệ; tất cả ảnh tham chiếu trong báo cáo tồn tại.
- Kiểm thử cuối: 8 passed in 6.82s; inference + Grad-CAM DenseNet121 trên ảnh validation đạt; Streamlit nhận cả hai checkpoint và chuyển model không lỗi.
