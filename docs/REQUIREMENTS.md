# Đối chiếu đặc tả

| Nhóm yêu cầu | Nơi triển khai | Bằng chứng sau khi chạy |
|---|---|---|
| Dataset thật, thống kê và imbalance | `scripts/download_data.py`, `src/dataset.py` | source.json, manifest.csv, dataset_counts.csv, dataset_distribution.png |
| Chống leakage, patient grouping | `patient_id`, `validate_manifest`, `prepare` | dataset_audit.json, hash và patient trong manifest |
| Grayscale → 3 kênh, resize, normalize | `src/transforms.py` | Cùng transform cho evaluate/inference |
| Augmentation train only | `XrayDataset(train=True)` | Transform xác định cho val/test |
| ResNet50, DenseNet121 ImageNet | `src/models.py` | pretrained metadata trong checkpoint |
| Freeze rồi fine-tune | `src/train.py` | stage trong history; backbone/BatchNorm đóng băng ở stage 1 |
| Weighted BCE | `src/train.py` | pos_weight và config trong checkpoint/run JSON |
| AdamW, scheduler, early stopping | `src/train.py` | LR trong CSV; checkpoint epoch tốt nhất |
| Training history đầy đủ | `src/train.py` | logs/history_resnet50.csv, history_densenet121.csv |
| Training plots | `training_plots` | loss, accuracy, precision/recall, F1, AUC PNG |
| Metric, confusion matrix, ROC | `src/metrics.py`, `src/evaluate.py` | predictions CSV, metrics JSON và PNG |
| So sánh, chọn model | `src/report.py` | model_comparison.csv, selected_model.json theo validation |
| Grad-CAM model đã train | `src/gradcam.py`, `load_checkpoint` | Original/heatmap/overlay theo model và case |
| Phân tích FP/FN | `src/evaluate.py`, `src/report.py` | results/errors, gradcam, case counts |
| Demo upload → validation → inference | `app/app.py`, `src/inference.py` | UI chọn 2 model và xác suất 2 lớp |
| Notebook, báo cáo, slide | `notebooks/analysis.ipynb`, `src/report.py` | report.md, presentation.pptx từ artifact thật |
| So sánh classifier / fine-tune | history stage và report | validation AUC tại checkpoint loss tốt nhất mỗi stage |

Mã nguồn đã có không đồng nghĩa thực nghiệm đã hoàn tất. Chỉ xác nhận hoàn tất training khi có `logs/run_*.json` cho cả hai model, và hoàn tất pipeline khi đã có prediction, biểu đồ, báo cáo/slide cùng checkpoint thật. Không dùng unit test hoặc tensor ngẫu nhiên làm kết quả đề tài.
