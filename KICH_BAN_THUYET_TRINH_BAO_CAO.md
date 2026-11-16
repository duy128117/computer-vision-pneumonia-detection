# KỊCH BẢN THUYẾT TRÌNH BÁO CÁO BẢO VỆ ĐỒ ÁN / KHÓA LUẬN
## ĐỀ TÀI: PHÂN LOẠI VIÊM PHỔI TRÊN ẢNH X-QUANG NGỰC SỬ DỤNG MẠNG NƠ-RON TÍCH CHẬP SÂU (PNEUMOGUARD-AI)

- **Học phần:** Thị giác máy tính nâng cao (CV CT2 F403)
- **Giảng viên hướng dẫn:** TS. Huỳnh Hữu Hưng
- **Nhóm thực hiện:** Nhóm 03
- **Thành viên báo cáo:**
  1. **Mai Phước Thịnh** (Thuyết trình Phần 1: Đặt vấn đề, Tiền xử lý dữ liệu chuẩn hóa, Tiến trình huấn luyện ResNet50 & DenseNet121 - Slide 1 đến Slide 6)
  2. **Nguyễn Đình Duy** (Thuyết trình Phần 2: Kiến trúc InceptionV3, Generalization Gap, Đánh giá đa chỉ số, Đánh đổi lâm sàng, Kiểm định McNemar, Giải pháp Cổng chú ý & Tổng kết - Slide 7 đến Slide 13)
- **Thời lượng dự kiến:** 12 – 15 phút trình bày + 5 – 10 phút hỏi đáp phản biện (Q&A).
- **Tài liệu tham chiếu đồng bộ:**
  - Slide trình chiếu: `d:\Downloads\So_Sanh_3_Mo_Hinh.pptx` (13 slide)
  - Báo cáo toàn văn: [baocao_final_viemphoi.docx](file:///d:/CODE/tgmt/baocao_final_viemphoi.docx) (44 trang chuẩn định dạng)
  - Bài báo khoa học: [Paper_Pneumonia_Classification.docx](file:///d:/CODE/tgmt/Paper_Pneumonia_Classification.docx) (Chuẩn IEEE Conference)

---

## MỤC LỤC PHÂN CHIA THỜI GIAN & PHÂN CÔNG BÁO CÁO

| STT | Slide | Tiêu đề Slide | Người trình bày | Thời gian |
|:---:|:---:|:---|:---:|:---:|
| 1 | Slide 1 | Giới thiệu Đề tài & Tác giả | Mai Phước Thịnh | 0:00 - 0:45 |
| 2 | Slide 2 | 01. Giới thiệu đề tài: Mục tiêu bài toán & Lộ trình nghiên cứu | Mai Phước Thịnh | 0:45 - 2:00 |
| 3 | Slide 3 | 02. Tổng quan dữ liệu: Quy trình làm sạch & Phân tập Patient-Disjoint | Mai Phước Thịnh | 2:00 - 3:45 |
| 4 | Slide 4 | 03. Tiến trình huấn luyện: ResNet50 (Đạt đỉnh Epoch 8) | Mai Phước Thịnh | 3:45 - 5:00 |
| 5 | Slide 5 | 03. Tiến trình huấn luyện: DenseNet121 (Đạt đỉnh Epoch 8) | Mai Phước Thịnh | 5:00 - 6:00 |
| 6 | Slide 6 | 04. So sánh mô hình lõi ban đầu: ResNet50 vs DenseNet121 | Mai Phước Thịnh | 6:00 - 7:30 |
| -- | -- | *Chuyển giao người trình bày (Handover to Nguyễn Đình Duy)* | Thịnh $\rightarrow$ Duy | 7:30 - 7:45 |
| 7 | Slide 7 | 05. Bổ sung mô hình InceptionV3: Ba cơ chế trích xuất đặc trưng | Nguyễn Đình Duy | 7:45 - 9:00 |
| 8 | Slide 8 | 06. Độ sụt giảm khái quát hóa (Generalization Gap) | Nguyễn Đình Duy | 9:00 - 10:15 |
| 9 | Slide 9 | 07. Hồ sơ hiệu năng đa chỉ số: Bảng so sánh toàn diện trên Test set | Nguyễn Đình Duy | 10:15 - 11:30 |
| 10 | Slide 10 | 08. Đánh đổi sai số lâm sàng: Chi phí của độ nhạy tuyệt đối | Nguyễn Đình Duy | 11:30 - 12:45 |
| 11 | Slide 11 | 09. Kiểm định thống kê McNemar & Hiệu quả tài nguyên phần cứng | Nguyễn Đình Duy | 12:45 - 13:45 |
| 12 | Slide 12 | 10. Tổng kết & Hướng phát triển (Cổng chú ý RoI & Grad-CAM) | Nguyễn Đình Duy | 13:45 - 14:45 |
| 13 | Slide 13 | Lời cảm ơn & Mời Hội đồng phản biện | Mai Phước Thịnh & Nguyễn Đình Duy | 14:45 - 15:00 |

---

# KỊCH BẢN CHI TIẾT TỪNG SLIDE (SLIDE-BY-SLIDE SCRIPT)

---

### SLIDE 1: TRANG TIÊU ĐỀ
- **Nội dung slide:**
  - Tên đề tài: Phân loại Viêm Phổi Trên Ảnh X-Quang Ngực Sử Dụng Mạng Nơ-ron Tích Chập Sâu (PneumoGuard-AI).
  - Sinh viên thực hiện: Mai Phước Thịnh – Nguyễn Đình Duy.
  - Lớp: CV CT2 F403. Giảng viên hướng dẫn: TS. Huỳnh Hữu Hưng.
- **Người trình bày:** Mai Phước Thịnh (0:00 – 0:45)
- **Mục tiêu slide:** Tạo ấn tượng mở đầu trang trọng, tự tin, nêu rõ danh tính và đề tài.
- **Lời thoại trình bày (Spoken Script):**
  > *"Kính thưa Thầy/Cô trong Hội đồng chấm đồ án học phần Thị giác máy tính nâng cao, cùng toàn thể các bạn sinh viên có mặt trong buổi báo cáo hôm nay.*
  > 
  > *Em tên là **Mai Phước Thịnh**, và cùng đồng hành với em là bạn **Nguyễn Đình Duy**, thuộc Nhóm 03. Dưới sự hướng dẫn tận tình của Thầy **TS. Huỳnh Hữu Hưng**, hôm nay nhóm chúng em xin được đại diện báo cáo kết quả nghiên cứu đề tài: **'Phân loại Viêm Phổi Trên Ảnh X-Quang Ngực Sử Dụng Mạng Nơ-ron Tích Chập Sâu - Hệ thống PneumoGuard-AI'**.*
  > 
  > *Nghiên cứu của chúng em tập trung giải quyết bài toán cốt lõi trong ứng dụng trí tuệ nhân tạo y tế: đó là làm sao để mô hình học sâu vừa phát hiện sớm ca bệnh, vừa kiểm soát nghiêm ngặt hiện tượng báo động giả, đồng thời loại trừ triệt để hiện tượng rò rỉ dữ liệu vốn đang tồn tại phổ biến ở các nghiên cứu trước đây. Sau đây, em xin phép được bắt đầu phần trình bày."*

---

### SLIDE 2: 01. GIỚI THIỆU ĐỀ TÀI - MỤC TIÊU BÀI TOÁN & LỘ TRÌNH NGHIÊN CỨU
- **Nội dung slide:**
  - Tiêu đề: 01. Giới thiệu đề tài - Mục tiêu bài toán và lộ trình nghiên cứu.
  - Trọng tâm nghiên cứu: Đánh giá sự đánh đổi giữa bỏ sót bệnh (False Negative) và cảnh báo giả (False Positive) trên tập kiểm thử độc lập.
  - 4 trụ cột lộ trình:
    1. Bài toán lâm sàng: Phát hiện sớm viêm phổi, giảm rủi ro tử vong do bỏ sót.
    2. Đối sánh nền tảng: So sánh mạng kết nối tắt ResNet50 và kết nối dày DenseNet121.
    3. Mở rộng đa tỷ lệ: Bổ sung mạng InceptionV3 để kiểm chứng nhánh đa tỷ lệ.
    4. Lựa chọn mô hình: Lựa chọn ResNet50 nhờ sự cân bằng tối ưu giữa 2 loại sai số.
- **Người trình bày:** Mai Phước Thịnh (0:45 – 2:00)
- **Mục tiêu slide:** Làm rõ bối cảnh y tế, định nghĩa 2 hiểm họa sai số (FN vs FP) và giới thiệu 4 bước tiếp cận khoa học của nhóm.
- **Lời thoại trình bày (Spoken Script):**
  > *"Kính thưa Hội đồng, theo Tổ chức Y tế Thế giới (WHO), viêm phổi vẫn là nguyên nhân truyền nhiễm gây tử vong hàng đầu ở trẻ em dưới 5 tuổi, cướp đi sinh mạng của hơn 800.000 trẻ mỗi năm. Chụp X-quang ngực (CXR) là tiêu chuẩn vàng để chẩn đoán, nhưng việc đọc phim đối mặt với 2 thách thức khốc liệt:*
  > 
  > * *Thứ nhất là sự quá tải trầm trọng tại các bệnh viện tuyến cơ sở, nơi thiếu vắng các bác sĩ chẩn đoán hình ảnh chuyên khoa.*
  > * *Thứ hai là cấu trúc lồng ngực nhi khoa chưa cốt hóa hoàn toàn, các tổn thương thâm nhiễm kính mờ thường bị che lấp bởi bóng tim và khung xương sườn.*
  > 
  > *Trong bối cảnh đó, một mô hình AI triển khai lâm sàng không thể chỉ nhìn vào một con số Accuracy chung chung. Nhóm chúng em đặt trọng tâm nghiên cứu vào **sự đánh đổi giữa hai loại hiểm họa sai số lâm sàng**:*
  > * *Bỏ sót ca bệnh (Âm tính giả - False Negative): Dẫn đến nguy cơ bệnh nhi bị suy hô hấp, nhiễm trùng huyết và tử vong.*
  > * *Cảnh báo nhầm (Dương tính giả - False Positive): Gây ra tình trạng quá tải chỉ định xét nghiệm, bức xạ không đáng có và lạm dụng kháng sinh gây kháng thuốc.*
  > 
  > *Để giải quyết bài toán, nhóm thiết lập lộ trình nghiên cứu 4 giai đoạn:*
  > * *Bước 1: Chuẩn hóa dữ liệu sạch và xây dựng tập benchmark cách ly bệnh nhân tuyệt đối.*
  > * *Bước 2: Xây dựng và đối sánh 2 kiến trúc kinh điển là ResNet50 (kết nối tắt) và DenseNet121 (kết nối dày đặc).*
  > * *Bước 3: Mở rộng thực nghiệm sang InceptionV3 để phân tích năng lực trích xuất đa tỷ lệ.*
  > * *Và Bước 4: Kiểm định thống kê nghiêm ngặt để chọn ra mô hình tối ưu nhất cho môi trường bệnh viện."*

---

### SLIDE 3: 02. TỔNG QUAN DỮ LIỆU - QUY TRÌNH LÀM SẠCH & PHÂN TẬP DỮ LIỆU
- **Nội dung slide:**
  - Bảng số liệu:
    * Tập Huấn luyện (Train): 3.815 ảnh (71,03%) - Cập nhật trọng số mạng kèm Data Augmentation.
    * Tập Xác thực (Validation): 938 ảnh (17,46%) - Theo dõi hội tụ, điều chỉnh Early Stopping.
    * Tập Kiểm thử (Test): 618 ảnh (11,51%) - Đánh giá khách quan độc lập (231 Bình thường, 387 Viêm phổi).
  - 4 thông số mấu chốt:
    * 5.371 ảnh X-quang ngực sạch được đưa vào sử dụng.
    * 485 ảnh lỗi bị loại bỏ (419 ảnh trùng lặp mã băm SHA-256 + 66 ảnh lỗi/hỏng cấu trúc DICOM-JPEG).
    * 70,6% tỷ lệ nhãn viêm phổi (mất cân bằng tự nhiên).
    * 618 ảnh kiểm thử độc lập Patient-Disjoint.
- **Người trình bày:** Mai Phước Thịnh (2:00 – 3:45)
- **Mục tiêu slide:** Khẳng định đóng góp phương pháp luận về việc làm sạch dữ liệu và giải thích tại sao kết quả của nhóm là trung thực, loại bỏ hoàn toàn rò rỉ dữ liệu (Data Leakage).
- **Lời thoại trình bày (Spoken Script):**
  > *"Điểm cốt lõi đầu tiên tạo nên sự khác biệt và tính trung thực học thuật của đề tài nằm ở khâu **Kiểm toán và Chuẩn hóa Dữ liệu** trên Slide số 3.*
  > 
  > *Hầu hết các nghiên cứu hiện nay trên Kaggle hay các bài báo khi tiếp cận bộ dữ liệu gốc Kermany thường sử dụng trực tiếp mà không kiểm tra tính toàn vẹn. Nhóm chúng em đã xây dựng một pipeline kiểm toán độc lập bằng mã băm mật mã học **SHA-256** và kiểm tra độ toàn vẹn pixel.*
  > 
  > *Kết quả kiểm toán đã phát hiện và **thanh lọc triệt để 485 bức ảnh bất thường**:*
  > * *Gồm 419 bức ảnh bị trùng lặp hoàn toàn nội dung nhị phân (Exact Bitwise Duplicates).*
  > * *Và 66 tệp ảnh bị hỏng tiêu đề, nhiễu nén nghiêm trọng hoặc không thể giải mã.*
  > 
  > *Sau khi sàng lọc, chúng em thu được **5.371 ảnh X-quang chuẩn hóa**. Toàn bộ dữ liệu được phân chia theo nguyên tắc nghiêm ngặt: **Cô lập định danh bệnh nhân (Patient-Disjoint Split)**. Điều này đảm bảo tuyệt đối rằng không có bất kỳ bệnh nhân nào xuất hiện đồng thời ở cả tập huấn luyện và tập kiểm thử.*
  > 
  > *Cụ thể: Tập Train gồm 3.815 ảnh (71,03%), tập Validation gồm 938 ảnh (17,46%), và tập Test độc lập gồm 618 ảnh (11,51%) với tỷ lệ chuẩn 231 ca bình thường và 387 ca viêm phổi.*
  > 
  > *Bên cạnh đó, tập dữ liệu có sự mất cân bằng với 70,6% ca viêm phổi. Thay vì áp dụng phương pháp Oversampling nhân bản ảnh dễ gây Overfitting, nhóm chúng em đã áp dụng hàm mất mát **Weighted Binary Cross-Entropy**, gán trọng số $w_{pos} = 0,706$ và $w_{neg} = 0,294$ để cân bằng độ dốc gradient một cách tự nhiên."*

---

### SLIDE 4: 03. TIẾN TRÌNH HUẤN LUYỆN - RESNET50 ĐẠT ĐỈNH TẠI EPOCH 8
- **Nội dung slide:**
  - Tiêu đề: 03. Tiến trình huấn luyện: ResNet50 đạt đỉnh tại Epoch 8.
  - Điểm mốc: Checkpoint lưu tại Epoch 8 với Validation Accuracy đạt 97,87%.
  - Cơ chế kiểm soát: Sau Epoch 8, Validation Loss bắt đầu dao động và không cải thiện. Cơ chế Early Stopping với patience = 4 đã kích hoạt và dừng huấn luyện ở Epoch 12 để ngăn chặn hiện tượng học vẹt (Overfitting).
- **Người trình bày:** Mai Phước Thịnh (3:45 – 5:00)
- **Mục tiêu slide:** Trình bày chi tiết động học hội tụ (Training Dynamics) của ResNet50, chứng minh mô hình được kiểm soát tránh Overfitting.
- **Lời thoại trình bày (Spoken Script):**
  > *"Tiếp theo, tại Slide số 4 là tiến trình huấn luyện của mô hình nền tảng đầu tiên: **ResNet50**.*
  > 
  > *Chúng em thiết lập mô hình với độ phân giải đầu vào $224 \times 224$, tối ưu hóa bằng thuật toán Adam với tốc độ học khởi tạo $\alpha = 10^{-4}$ kết hợp cơ chế giảm tốc độ học ReduceLROnPlateau.*
  > 
  > *Quan sát đường cong huấn luyện, Hội đồng có thể thấy:*
  > * *Trong 5 epoch đầu tiên, nhờ các khối Residual Block với kết nối tắt (Skip Connections), gradient được lan truyền thông suốt, Training Loss giảm nhanh từ 0,55 xuống dưới 0,15.*
  > * *Mô hình đạt trạng thái tối ưu toàn cục tại **Epoch thứ 8**, với Validation Accuracy đạt đỉnh **97,87%** và Validation Loss đạt mức cực tiểu.*
  > * *Từ Epoch 9 đến Epoch 12, mặc dù Training Loss tiếp tục giảm nhẹ, nhưng Validation Loss bắt đầu có hiện tượng đi ngang và dao động. Cơ chế **Early Stopping** với ngưỡng kiên nhẫn (patience) bằng 4 chu kỳ đã can thiệp chính xác tại Epoch 12, tự động ngắt huấn luyện và khôi phục lại bộ trọng số tốt nhất tại Epoch 8.*
  > 
  > *Điều này khẳng định mô hình đã được bảo vệ chặt chẽ, không bị suy thoái do hiện tượng ghi nhớ mẫu dữ liệu."*

---

### SLIDE 5: 03. TIẾN TRÌNH HUẤN LUYỆN - DENSENET121 ĐẠT ĐỈNH TẠI EPOCH 8
- **Nội dung slide:**
  - Tiêu đề: 03. Tiến trình huấn luyện: DenseNet121 đạt đỉnh tại Epoch 8.
  - Điểm mốc: Checkpoint lưu tại Epoch 8, độ chính xác xác thực đạt 97,87% tương đương ResNet50.
  - Cơ chế kiểm soát: Quá trình hội tụ ổn định, kích hoạt Early Stopping và dừng ở Epoch 12.
- **Người trình bày:** Mai Phước Thịnh (5:00 – 6:00)
- **Mục tiêu slide:** Giới thiệu quá trình hội tụ của mô hình đối chứng DenseNet121, nhấn mạnh sự tương đồng trên tập xác thực.
- **Lời thoại trình bày (Spoken Script):**
  > *"Tại Slide số 5 là tiến trình huấn luyện của mô hình đối chứng: **DenseNet121**.*
  > 
  > *DenseNet121 sử dụng cơ chế ghép nối dày đặc (Dense Connectivity), nơi mỗi lớp đều nhận bản đồ đặc trưng từ tất cả các lớp phía trước. Điều này giúp mạng tái sử dụng đặc trưng tối đa và giảm số lượng tham số xuống chỉ còn 6,95 triệu.*
  > 
  > *Về mặt động học hội tụ trên tập xác thực nội bộ:*
  > * *DenseNet121 thể hiện sự hội tụ rất tương đồng với ResNet50, cũng đạt đỉnh xác thực cao nhất tại **Epoch 8** với độ chính xác **97,87%**.*
  > * *Và thuật toán cũng kích hoạt Early Stopping dừng huấn luyện tại Epoch 12 để bảo toàn trọng số tối ưu.*
  > 
  > *Như vậy, nếu chỉ dừng lại ở tập Validation, hai mô hình gần như 'bất phân thắng bại'. Nhưng sự khác biệt bản chất chỉ thực sự lộ diện khi chúng em đưa cả hai vào bài kiểm tra khắc nghiệt trên tập kiểm thử độc lập."*

---

### SLIDE 6: 04. SO SÁNH MÔ HÌNH LÕI BAN ĐẦU - RESNET50 VÀ DENSENET121
- **Nội dung slide:**
  - Tiêu đề: 04. So sánh mô hình lõi ban đầu: Đối sánh hiệu năng ResNet50 và DenseNet121.
  - Kết luận bước 1: Hai mô hình hội tụ ngang nhau ở chu kỳ 8 (97,87%), nhưng ResNet50 giảm được 26 ca báo động giả trên tập kiểm thử.
  - Số liệu so sánh trên 618 ảnh test:
    * **ResNet50 (Mô hình được chọn):** Accuracy 90,61% | Specificity 75,76% | F1 93,00% | AUC 96,79% | False Positive: 56 ca.
    * **DenseNet121 (Mô hình đối chứng):** Accuracy 86,57% | Specificity 64,50% | F1 90,29% | AUC 95,78% | False Positive: 82 ca.
- **Người trình bày:** Mai Phước Thịnh (6:00 – 7:30)
- **Mục tiêu slide:** Chỉ ra sự vượt trội rõ rệt của ResNet50 so với DenseNet121 trên tập Test, đặc biệt là giảm thiểu 26 ca dương tính giả.
- **Lời thoại trình bày (Spoken Script):**
  > *"Kính thưa Thầy/Cô, Slide số 6 thể hiện bước so sánh then chốt đầu tiên của đề tài.*
  > 
  > *Khi đánh giá trên 618 ảnh của tập kiểm thử độc lập, chúng em nhận thấy sự phân hóa rất rõ ràng:*
  > * *Về độ chính xác tổng thể (Accuracy): **ResNet50 đạt 90,61%**, vượt trội hơn hẳn so với mức **86,57%** của DenseNet121 (chênh lệch hơn 4,04 điểm phần trăm).*
  > * *Nhưng chỉ số mang tính quyết định lâm sàng nằm ở **Độ đặc hiệu (Specificity)**: ResNet50 đạt **75,76%**, trong khi DenseNet121 rơi tự do xuống chỉ còn **64,50%**.*
  > 
  > *Quy đổi sang số ca bệnh thực tế: Số ca dương tính giả (False Positive) của DenseNet121 vọt lên tới 82 ca, trong khi ResNet50 chặn lại ở mức 56 ca.*
  > 
  > *Nói cách khác, **ResNet50 đã giúp bệnh viện loại bỏ được 26 ca cảnh báo giả**. Việc ghép nối các đặc trưng cấp thấp xuyên suốt trong DenseNet121 vô tình khiến mô hình bị nhạy cảm quá mức với các bóng mờ sinh lý của phế huyết quản bình thường, dẫn đến chẩn đoán nhầm thành tổn thương đông đặc.*
  > 
  > *Từ kết quả này, ResNet50 đã khẳng định vị thế là mô hình lõi vượt trội. Để kiểm chứng sâu hơn về khả năng phân giải các tổn thương thâm nhiễm đa kích thước, nhóm chúng em tiếp tục mở rộng thực nghiệm sang mô hình thứ ba: InceptionV3. Phần phân tích chuyên sâu tiếp theo xin kính mời bạn **Nguyễn Đình Duy** trình bày."*

---

### SLIDE 7: 05. BỔ SUNG MÔ HÌNH INCEPTIONV3 - BA CƠ CHẾ TRÍCH XUẤT ĐẶC TRƯNG
- **Nội dung slide:**
  - Tiêu đề: 05. Bổ sung mô hình InceptionV3: Ba cơ chế trích xuất đặc trưng trên ảnh X-quang.
  - 3 cơ chế kiến trúc:
    1. **ResNet50 (Kết nối tắt):** Truyền đạo hàm ổn định qua mạng sâu, bảo toàn thông tin gốc qua phép cộng tắt (Identity Addition).
    2. **DenseNet121 (Kết nối dày):** Tái sử dụng đặc trưng liên tục qua phép ghép kênh (Channel Concatenation), tiết kiệm tham số.
    3. **InceptionV3 (Đa tỷ lệ):** Sử dụng các bộ lọc đa kích thước ($1 \times 1, 3 \times 3, 5 \times 5$ phân rã thành $1 \times 7, 7 \times 1$) để học đồng thời các vùng mờ nhỏ lẫn tổn thương lớn.
- **Người trình bày:** Nguyễn Đình Duy (7:45 – 9:00)
- **Mục tiêu slide:** Tiếp nhận phần trình bày mượt mà, phân tích lý do khoa học chọn InceptionV3 và so sánh nguyên lý trích xuất đặc trưng của 3 họ kiến trúc.
- **Lời thoại trình bày (Spoken Script):**
  > *"Em xin cảm ơn bạn Thịnh. Kính thưa Hội đồng, em tên là **Nguyễn Đình Duy**, sau đây em xin phép được tiếp nối phần trình bày của nhóm.*
  > 
  > *Như Hội đồng có thể quan sát trên Slide số 7, tổn thương viêm phổi trên ảnh X-quang có đặc tính bệnh học vô cùng đa dạng:*
  > * *Ở giai đoạn sớm hoặc do virus, tổn thương chỉ là những nốt mờ kẽ vi thể rải rác.*
  > * *Nhưng ở giai đoạn viêm phổi thùy do vi khuẩn, tổn thương lan rộng tạo thành các khối đông đặc che phủ toàn bộ thùy phổi.*
  > 
  > *Một bộ lọc tích chập có kích thước cố định như $3 \times 3$ trong ResNet hay DenseNet có thể gặp khó khăn khi phải vừa bắt các nốt mờ nhỏ, vừa cảm nhận được ranh giới tổn thương diện rộng. Đó chính là lý do chúng em đưa **InceptionV3** vào nghiên cứu.*
  > 
  > *Slide 7 so sánh trực quan triết lý thiết kế của ba trường phái kiến trúc:*
  > * *ResNet50: Dùng kết nối tắt $x + F(x)$ để chống tiêu biến đạo hàm khi mạng đi sâu.*
  > * *DenseNet121: Nối chồng các lớp đặc trưng $[x_0, x_1, ...]$ để tái sử dụng triệt để thông tin mức thấp.*
  > * *Còn InceptionV3: Trong mỗi khối Inception, ảnh được quét đồng thời qua nhiều nhánh tích chập với các trường tiếp nhận (receptive fields) khác nhau từ $1 \times 1, 3 \times 3$ cho tới các cặp tích chập bất đối xứng $1 \times 7$ và $7 \times 1$. Nhờ đó, mô hình có khả năng trích xuất đặc trưng đa tỷ lệ một cách tự nhiên."*

---

### SLIDE 8: 06. ĐỘ SỤT GIẢM KHÁI QUÁT HÓA (GENERALIZATION GAP)
- **Nội dung slide:**
  - Tiêu đề: 06. Độ sụt giảm khái quát hóa: So sánh độ chính xác xác thực và kiểm thử độc lập.
  - Hiện tượng mấu chốt: InceptionV3 đạt đỉnh xác thực cao nhất tại chu kỳ 12 (98,51%) nhưng lại sụt giảm mạnh nhất khi đánh giá trên tập kiểm thử độc lập (xuống 86,08%, độ sụt giảm lên tới 12,43%).
  - ResNet50 có độ sụt giảm thấp hơn và ổn định nhất (từ 97,87% xuống 90,61%, gap 7,26%).
- **Người trình bày:** Nguyễn Đình Duy (9:00 – 10:15)
- **Mục tiêu slide:** Chỉ ra hiện tượng ảo giác độ chính xác nội bộ (Internal Validation Illusion) và minh chứng tại sao cần đánh giá trên tập kiểm thử độc lập.
- **Lời thoại trình bày (Spoken Script):**
  > *"Bước sang Slide số 8, chúng em xin trình bày một phát hiện thực nghiệm mang tính cảnh báo sâu sắc về phương pháp luận: **Độ sụt giảm khái quát hóa (Generalization Gap)**.*
  > 
  > *Khi theo dõi quá trình huấn luyện:*
  > * *InceptionV3 thể hiện năng lực 'học' cực kỳ ấn tượng, đạt đỉnh Validation Accuracy lên tới **98,51%** tại Epoch 12, vượt qua cả ResNet50 (97,87%).*
  > * *Nếu một nhóm nghiên cứu thiếu kinh nghiệm vội vàng công bố kết quả dựa trên tập Validation này, họ sẽ kết luận InceptionV3 là mô hình số một.*
  > 
  > *Tuy nhiên, bức tranh thực tế đảo chiều hoàn toàn khi đưa mô hình sang tập kiểm thử độc lập 618 ảnh:*
  > * *Độ chính xác của InceptionV3 **lao dốc từ 98,51% xuống chỉ còn 86,08%**, tạo ra khoảng cách sụt giảm khổng lồ lên tới **12,43 điểm phần trăm**!*
  > * *Trong khi đó, ResNet50 duy trì sự ổn định bền vững hơn rất nhiều, với mức rơi từ 97,87% xuống 90,61% (khoảng cách chỉ là 7,26 điểm phần trăm).*
  > 
  > *Hiện tượng này giải thích cho việc: Mạng InceptionV3 với cấu trúc phân nhánh phức tạp đã vô tình học thuộc (overfit) vào các đặc trưng kỹ thuật chụp của tập huấn luyện nội bộ. Khi đối mặt với ảnh chụp từ các máy X-quang và bệnh nhi hoàn toàn mới ở tập Test, mạng bị bối rối và mất tính khái quát hóa."*

---

### SLIDE 9: 07. HỒ SƠ HIỆU NĂNG ĐA CHỈ SỐ - BẢNG SO SÁNH TOÀN DIỆN TRÊN TẬP KIỂM THỬ
- **Nội dung slide:**
  - Bảng so sánh 3 mô hình trên 618 ảnh test độc lập:

| Chỉ số đánh giá | ResNet50 (Mô hình chính) | DenseNet121 | InceptionV3 |
|:---|:---:|:---:|:---:|
| **Độ chính xác (Accuracy)** | **90,61%** | 86,57% | 86,08% |
| **Độ chuẩn xác (Precision)** | **87,30%** | 82,48% | 81,82% |
| **Độ nhạy (Sensitivity)** | 99,48% | 99,74% | **100,00%** |
| **Độ đặc hiệu (Specificity)** | **75,76%** | 64,50% | 62,77% |
| **Điểm cân bằng F1** | **93,00%** | 90,29% | 90,00% |
| **Diện tích dưới đường cong ROC (AUC)** | **96,79%** | 95,78% | 96,29% |

- **Người trình bày:** Nguyễn Đình Duy (10:15 – 11:30)
- **Mục tiêu slide:** Phân tích toàn diện 6 chỉ số đánh giá, nhấn mạnh sự dẫn đầu toàn diện của ResNet50 trên 5/6 thước đo, và nhận diện hiện tượng Sensitivity 100% của InceptionV3.
- **Lời thoại trình bày (Spoken Script):**
  > *"Slide số 9 là **Bảng tổng hợp hồ sơ hiệu năng đa chỉ số** của cả 3 mô hình trên tập kiểm thử độc lập.*
  > 
  > *Nhìn vào bảng số liệu, Hội đồng có thể nhận thấy ngay:*
  > * *Về **Độ chính xác tổng thể (Accuracy)**: ResNet50 dẫn đầu tuyệt đối với **90,61%**, cao hơn DenseNet121 (+4,04%) và InceptionV3 (+4,53%).*
  > * *Về **Độ chuẩn xác (Precision)**: ResNet50 đạt **87,30%**, vượt xa mức ~82% của hai mô hình còn lại.*
  > * *Về **Điểm F1-score và AUC**: ResNet50 lần lượt đạt **93,00%** và **96,79%**, chứng minh khả năng phân tách ranh giới xác suất giữa 2 lớp là tối ưu nhất.*
  > 
  > *Tuy nhiên, có một điểm số liệu vô cùng thú vị mà chúng em muốn kính xin Hội đồng lưu tâm:*
  > * *Đó là ở chỉ số **Độ nhạy (Sensitivity)**: InceptionV3 đạt con số tuyệt đối **100,00%**, trong khi ResNet50 là **99,48%**.*
  > * *Nhưng ngược lại, ở chỉ số **Độ đặc hiệu (Specificity)**: InceptionV3 lại sụt giảm nghiêm trọng xuống chỉ còn **62,77%**, kém ResNet50 tới gần 13 điểm phần trăm (75,76%).*
  > 
  > *Tại sao lại có hiện tượng một mô hình đạt độ nhạy 100% nhưng độ chính xác chung lại thấp nhất? Câu trả lời nằm ở bài toán đánh đổi sai số lâm sàng ở Slide số 10."*

---

### SLIDE 10: 08. ĐÁNH ĐỔI SAI SỐ LÂM SÀNG - CHI PHÍ CỦA ĐỘ NHẠY TUYỆT ĐỐI
- **Nội dung slide:**
  - Bảng ma trận nhầm lẫn (Confusion Matrix):
    * Âm tính thật (True Negative - Đoán đúng khỏe): ResNet50 = 175 ca vs InceptionV3 = 145 ca.
    * Dương tính giả (False Positive - Báo nhầm bệnh): ResNet50 = 56 ca vs InceptionV3 = **86 ca (+30 ca)**.
    * Âm tính giả (False Negative - Bỏ sót bệnh): ResNet50 = 2 ca vs InceptionV3 = **0 ca (-2 ca)**.
    * Dương tính thật (True Positive - Đoán đúng bệnh): ResNet50 = 385 ca vs InceptionV3 = 387 ca.
  - Bản chất đánh đổi: **-2 Bỏ sót $\leftrightarrow$ +30 Báo giả**.
  - Đánh giá: Độ nhạy tăng 0,52% nhưng Độ đặc hiệu giảm tới 12,81%. Báo động giả tăng 53,6% gây quá tải hệ thống y tế.
- **Người trình bày:** Nguyễn Đình Duy (11:30 – 12:45)
- **Mục tiêu slide:** Đây là slide đắt giá nhất về mặt tư duy ứng dụng y tế, phân tích sâu sắc bài toán kinh tế - y tế giữa số ca bỏ sót và số ca báo nhầm.
- **Lời thoại trình bày (Spoken Script):**
  > *"Kính thưa Thầy/Cô, đây là nội dung mà nhóm chúng em tâm đắc nhất trong toàn bộ công trình nghiên cứu: **Sự đánh đổi sai số lâm sàng** tại Slide số 10.*
  > 
  > *Nhiều người khi nhìn vào InceptionV3 đạt Sensitivity 100% sẽ nghĩ ngay rằng đây là mô hình hoàn hảo vì 'không bỏ sót bất kỳ ca bệnh nào' (Âm tính giả = 0 ca, so với 2 ca của ResNet50).*
  > 
  > *Nhưng trong y tế, không có bữa trưa nào là miễn phí! Để đổi lấy việc giảm được **2 ca bỏ sót**, InceptionV3 đã làm bùng nổ số ca cảnh báo giả từ 56 ca lên tới **86 ca** — tức là phát sinh thêm **30 ca dương tính giả**!*
  > 
  > *Hãy hình dung tác động lâm sàng tại một bệnh viện nhi:*
  > * *30 ca trẻ em hoàn toàn khỏe mạnh bị hệ thống AI dán nhãn là 'Viêm phổi'.*
  > * *Hậu quả kéo theo là 30 gia đình rơi vào hoang mang; 30 đứa trẻ phải chịu thêm các xét nghiệm xâm lấn, chụp CT liều cao hoặc bị bác sĩ chỉ định dùng kháng sinh dự phòng oan uổng, gia tăng nguy cơ kháng kháng sinh cộng đồng.*
  > * *Đồng thời, hiện tượng báo động giả tràn lan sẽ gây ra hội chứng 'Kiệt quệ cảnh báo' (Alert Fatigue) ở bác sĩ, khiến họ mất niềm tin và bỏ qua cảnh báo của AI.*
  > 
  > *Đối với ResNet50, mô hình chỉ để lọt 2 ca bỏ sót trong tổng số 387 ca bệnh (đạt độ nhạy cực cao 99,48%), nhưng lại bảo vệ thành công 175 trẻ khỏe mạnh và giữ số ca báo giả ở mức thấp nhất. Đây là điểm cân bằng lâm sàng hoàn hảo và có trách nhiệm nhất."*

---

### SLIDE 11: 09. KIỂM ĐỊNH THỐNG KÊ & HIỆU QUẢ TÀI NGUYÊN
- **Nội dung slide:**
  - Kết quả kiểm định phi tham số McNemar Test:
    * ResNet50 vs DenseNet121: $p = 0,000070 < 0,001$ $\rightarrow$ Khác biệt có ý nghĩa thống kê cao.
    * ResNet50 vs InceptionV3: $p = 0,000025 < 0,001$ $\rightarrow$ Khác biệt có ý nghĩa thống kê cao.
    * DenseNet121 vs InceptionV3: $p = 0,742829 > 0,05$ $\rightarrow$ Không có sự khác biệt có ý nghĩa thống kê.
  - Thời gian huấn luyện thực tế (trên GPU NVIDIA RTX 3060 6GB, 12 epochs):
    * **ResNet50:** 15 phút 15 giây (23,51M tham số) - Huấn luyện nhanh nhất.
    * **DenseNet121:** 18 phút 46 giây (6,95M tham số) - Nhẹ tham số nhưng tốn bộ nhớ ghép tensor.
    * **InceptionV3:** 23 phút 41 giây (21,79M tham số) - Tốn thời gian nhất do nhiều nhánh song song.
- **Người trình bày:** Nguyễn Đình Duy (12:45 – 13:45)
- **Mục tiêu slide:** Cung cấp bằng chứng toán học chứng minh kết quả không ngẫu nhiên (McNemar test) và so sánh chi phí tính toán thực tế.
- **Lời thoại trình bày (Spoken Script):**
  > *"Để các kết luận trên đạt chuẩn mực của một bài báo khoa học quốc tế, chúng em không chỉ dừng lại ở các con số thống kê mô tả, mà đã tiến hành **Kiểm định giả thuyết thống kê phi tham số McNemar** trên Slide số 11.*
  > 
  > *Kiểm định McNemar tập trung phân tích ma trận bất đồng — tức là những mẫu ảnh mà mô hình này đoán đúng nhưng mô hình kia đoán sai:*
  > * *Khi so sánh ResNet50 với DenseNet121, giá trị $p = 0,000070$, nhỏ hơn rất nhiều so với ngưỡng $\alpha = 0,001$.*
  > * *Khi so sánh ResNet50 với InceptionV3, giá trị $p = 0,000025$, cũng đạt mức ý nghĩa cực cao.*
  > * *Ngược lại, giữa DenseNet121 và InceptionV3 thì $p = 0,743 > 0,05$.*
  > 
  > *Điều này cung cấp bằng chứng toán học đanh thép khẳng định rằng: **Sự vượt trội của ResNet50 là có ý nghĩa thống kê thực sự**, hoàn toàn không phải do may rủi hay ngẫu nhiên trong quá trình bốc thăm dữ liệu.*
  > 
  > *Xét về mặt chi phí tài nguyên phần cứng trên GPU thực nghiệm RTX 3060:*
  > * *ResNet50 hoàn thành huấn luyện chỉ sau **15 phút 15 giây**, nhanh nhất trong cả 3 mô hình.*
  > * *DenseNet121 tuy chỉ có gần 7 triệu tham số nhưng mất tới 18 phút 46 giây do chi phí quản lý bộ nhớ và ghép nối tensor liên tục.*
  > * *Và InceptionV3 chậm nhất với 23 phút 41 giây do tính toán đa nhánh phức tạp."*

---

### SLIDE 12: 10. TỔNG KẾT & HƯỚNG PHÁT TRIỂN
- **Nội dung slide:**
  - Tiêu đề: 10. Tổng kết & Hướng phát triển: Lý do khẳng định ResNet50 là mô hình chính.
  - 3 điểm then chốt:
    * **90,61% Hiệu năng vượt trội:** Dẫn đầu toàn diện các chỉ số kiểm thử độc lập.
    * **75,76% Kiểm soát báo giả:** Đạt điểm cân bằng lâm sàng tối ưu giữa độ nhạy và độ đặc hiệu.
    * **Giải pháp mở rộng (Grad-CAM & Attention Gate):** Tích hợp kỹ thuật giải thích trực quan Grad-CAM và Cổng chú ý không gian (Spatial Lung-RoI Attention Gate) để loại bỏ Shortcut Learning, tăng +13,54% mật độ năng lượng tập trung vào nhu mô phổi.
- **Người trình bày:** Nguyễn Đình Duy (13:45 – 14:45)
- **Mục tiêu slide:** Chốt hạ thông điệp khoa học của đề tài và trình bày giải pháp nâng cao đã nghiên cứu trong báo cáo toàn văn và bài báo khoa học.
- **Lời thoại trình bày (Spoken Script):**
  > *"Kính thưa Thầy/Cô trong Hội đồng, từ toàn bộ chuỗi thực nghiệm trên Slide số 12, nhóm chúng em đi đến kết luận khẳng định: **ResNet50 chính là kiến trúc xương sống tối ưu nhất** cho hệ thống hỗ trợ chẩn đoán viêm phổi PneumoGuard-AI.*
  > 
  > *Mô hình hội tụ xuất sắc 3 ưu điểm vượt trội: Hiệu năng tổng thể cao nhất (90,61%), khả năng kiểm soát báo động giả tốt nhất (75,76% độ đặc hiệu), và tốc độ huấn luyện cũng như suy luận nhanh nhất.*
  > 
  > *Về hướng phát triển mở rộng đã được nhóm nghiên cứu sâu trong báo cáo toàn văn và bài báo gửi kèm:*
  > * *Để biến mô hình từ một 'Hộp đen' (Black-box) thành một công cụ minh bạch hỗ trợ bác sĩ, nhóm đã tích hợp bản đồ nhiệt **Grad-CAM** để khoanh vùng chính xác các ổ tổn thương thâm nhiễm.*
  > * *Đặc biệt, để khắc phục triệt để hiện tượng mạng nơ-ron bị đánh lừa bởi các ký tự kỹ thuật viên hay khung xương ngoài phổi (Shortcut Learning), nhóm đã đề xuất thành công khối **Cổng chú ý không gian (Spatial Lung-RoI Guided Attention Gate)**.*
  > * *Khối chú ý này chỉ bổ sung vỏn vẹn **98 tham số** nhưng giúp tăng mật độ năng lượng kích hoạt tập trung vào đúng nhu mô phổi lên thêm **+13,54%**, mang lại khả năng giải thích tổn thương trực quan và an toàn tối đa cho môi trường lâm sàng."*

---

### SLIDE 13: LỜI CẢM ƠN & KẾT THÚC
- **Nội dung slide:**
  - Tiêu đề: Trân trọng cảm ơn Thầy Cô và các bạn đã lắng nghe.
  - Thông tin liên hệ: Nhóm 03 - Sinh viên Mai Phước Thịnh & Nguyễn Đình Duy - Lớp CV CT2 F403.
- **Người trình bày:** Cả hai thành viên (14:45 – 15:00)
- **Mục tiêu slide:** Kết thúc bài thuyết trình đúng thời gian, lịch sự, chuyên nghiệp và chuyển sang phiên hỏi đáp phản biện.
- **Lời thoại trình bày (Spoken Script):**
  > *(Mai Phước Thịnh & Nguyễn Đình Duy cùng đứng trang nghiêm)*
  > 
  > **Mai Phước Thịnh:** *"Kính thưa quý Thầy/Cô trong Hội đồng, trên đây là toàn bộ nội dung báo cáo kết quả nghiên cứu đồ án của Nhóm 03."*
  > 
  > **Nguyễn Đình Duy:** *"Chúng em xin trân trọng gửi lời cảm ơn sâu sắc đến Thầy TS. Huỳnh Hữu Hưng đã tận tình định hướng phương pháp luận khoa học, cùng quý Thầy Cô trong Hội đồng đã dành thời gian lắng nghe.*
  > 
  > *Nhóm chúng em rất mong nhận được những nhận xét, góp ý quý báu từ quý Thầy Cô để công trình nghiên cứu được hoàn thiện hơn nữa. Sau đây, nhóm chúng em xin sẵn sàng lắng nghe các câu hỏi phản biện từ Hội đồng. Chúng em xin trân trọng cảm ơn!"*

---

# BỘ CÂU HỎI PHẢN BIỆN DỰ PHÒNG & HƯỚNG DẪN TRẢ LỜI (Q&A DEFENSE GUIDE)

Dưới đây là 8 câu hỏi phản biện hóc búa nhất mà Hội đồng chấm đồ án/bài báo thường đặt ra, kèm theo câu trả lời chuẩn xác, gắn chặt số liệu từ báo cáo 44 trang và bài báo IEEE:

---

### CÂU HỎI 1: "Tại sao độ chính xác trên tập Test của các em là 90,61%, trong khi nhiều bài báo và các kernel trên Kaggle đạt 98% đến 99%?"
- **Người trả lời ưu tiên:** Mai Phước Thịnh.
- **Chiến lược trả lời:** Khẳng định tính trung thực khoa học, bóc trần hiện tượng Data Leakage (Rò rỉ dữ liệu).
- **Nội dung trả lời:**
  > *"Em xin cảm ơn câu hỏi rất thực tế và sâu sắc của Thầy/Cô. Đây cũng chính là trăn trở lớn nhất của nhóm khi bắt đầu nghiên cứu.*
  > 
  > *Các nghiên cứu đạt 98-99% trên bộ dữ liệu Kermany thường mắc phải một sai lầm nghiêm trọng về phương pháp luận: đó là **Rò rỉ định danh bệnh nhân (Patient Identity Leakage)**.*
  > 
  > *Bộ dữ liệu gốc có nhiều ảnh chụp của cùng một bệnh nhi ở các góc chụp hoặc thời điểm khác nhau. Khi chia tập ngẫu nhiên theo từng file ảnh (Random Image Split), có tới **51,3% bệnh nhân ở tập test xuất hiện đồng thời trong tập train**. Khi đó, mạng nơ-ron không học đặc trưng bệnh lý viêm phổi, mà nó 'học vẹt' hình dạng xương sườn, tư thế nằm và góc chụp riêng của từng đứa trẻ.*
  > 
  > *Trong nghiên cứu này, nhóm chúng em kiên quyết làm sạch dữ liệu: loại bỏ 419 ảnh trùng lặp bitwise bằng SHA-256 và áp dụng cơ chế chia tập **cô lập bệnh nhân tuyệt đối (Patient-Disjoint)**. Con số **90,61%** là độ chính xác thực tế trên những bệnh nhân hoàn toàn xa lạ với mô hình. Nhóm chúng em lựa chọn sự trung thực khoa học và giá trị ứng dụng lâm sàng thực chất thay vì một con số ảo giác do rò rỉ dữ liệu."*

---

### CÂU HỎI 2: "Tại sao nhóm cấm phép tăng cường dữ liệu lật ảnh theo chiều ngang (Horizontal Flip) khi huấn luyện?"
- **Người trả lời ưu tiên:** Nguyễn Đình Duy.
- **Chiến lược trả lời:** Dựa trên giải phẫu học X-quang lồng ngực (CXR Anatomy).
- **Nội dung trả lời:**
  > *"Dạ em xin cảm ơn Thầy/Cô. Trong xử lý ảnh thông thường, lật ngang là phép tăng cường dữ liệu rất phổ biến. Tuy nhiên, trong ảnh X-quang lồng ngực người, lật ngang là một điều **tối kỵ về mặt giải phẫu học**:*
  > 
  > *1. Cơ thể người có tính bất đối xứng nội tạng nghiêm ngặt: Trục tim và mỏm tim bình thường nằm lệch về bên lồng ngực trái, quai động mạch chủ nằm bên trái, trong khi vòm hoành phải luôn cao hơn vòm hoành trái khoảng 1 đến 2 cm do có khối gan nâng đỡ.*
  > *2. Nếu lật ảnh ngang, mô hình sẽ bị ép học những ca giải phẫu 'đảo ngược phủ tạng' (Situs Inversus / Dextrocardia) nhân tạo, gây nhiễu loạn phân phối không gian của các cấu trúc giải phẫu chuẩn.*
  > 
  > *Do đó, nhóm chỉ sử dụng các phép tăng cường bảo toàn tính toàn vẹn giải phẫu như: xoay nhẹ trong phạm vi $\pm 7^\circ$, tịnh tiến nhẹ $5\%$, điều chỉnh độ sáng tương phản và biến dạng đàn hồi affine nhỏ."*

---

### CÂU HỎI 3: "Tại sao nhóm dùng Weighted BCE Loss mà không dùng kỹ thuật Oversampling hay SMOTE để giải quyết mất cân bằng dữ liệu?"
- **Người trả lời ưu tiên:** Mai Phước Thịnh.
- **Chiến lược trả lời:** Phân tích nhược điểm của Oversampling và SMOTE trên không gian pixel ảnh y tế.
- **Nội dung trả lời:**
  > *"Dạ em xin cảm ơn Thầy/Cô. Tập dữ liệu có tỷ lệ viêm phổi chiếm 70,6% và bình thường chiếm 29,4%. Nhóm không dùng Oversampling hay SMOTE vì:*
  > 
  > *1. **Random Oversampling (Nhân bản ảnh bình thường):** Khi nhân bản các ảnh âm tính, các đặc trưng cục bộ của những bức ảnh đó sẽ xuất hiện lặp đi lặp lại nhiều lần trong các batch huấn luyện, khiến mạng nơ-ron rất dễ bị Overfitting vào các đặc trưng cá thể của vài bệnh nhân đó.*
  > *2. **SMOTE (Nội suy lân cận):** SMOTE hoạt động tốt trên dữ liệu bảng (tabular), nhưng trên không gian ma trận pixel ảnh X-quang, phép nội suy tuyến tính giữa 2 bức ảnh sẽ tạo ra những hình ảnh 'bóng ma' (ghosting artifacts), làm biến dạng ranh giới các xương sườn và các nốt mờ giải phẫu thật.*
  > 
  > *Vì vậy, giải pháp tối ưu và an toàn nhất là **Weighted Binary Cross-Entropy**. Chúng em gán trọng số trực tiếp vào hàm mất mát: mẫu bình thường được nhân hệ số $w_{neg} = 0,294$ và mẫu viêm phổi nhân $w_{pos} = 0,706$. Nhờ đó, gradient cập nhật từ mỗi mẫu bình thường có sức nặng lớn hơn tương ứng, giúp cân bằng quá trình tối ưu mà không làm sai lệch phân phối ảnh tự nhiên."*

---

### CÂU HỎI 4: "Kiểm định McNemar có ý nghĩa gì trong bài toán này? Tại sao không dùng Paired t-test?"
- **Người trả lời ưu tiên:** Nguyễn Đình Duy.
- **Chiến lược trả lời:** Nắm vững lý thuyết xác suất thống kê và điều kiện áp dụng kiểm định.
- **Nội dung trả lời:**
  > *"Dạ em xin cảm ơn Thầy/Cô. Lý do nhóm chọn kiểm định McNemar thay vì Paired t-test xuất phát từ bản chất của biến số đo lường:*
  > 
  > *1. **Paired t-test** chỉ áp dụng cho biến số ngẫu nhiên liên tục và yêu cầu hiệu số giữa các cặp phải tuân theo phân phối chuẩn (Gaussian distribution).*
  > *2. Tuy nhiên, kết quả dự đoán trên từng ảnh kiểm thử là biến định danh nhị phân (Đúng = 1 hoặc Sai = 0). Do đó phân phối không thể là phân phối chuẩn.*
  > 
  > * **Kiểm định McNemar** là kiểm định phi tham số chuẩn mực quốc tế dành riêng cho hai bộ phân loại bắt cặp trên cùng một tập kiểm thử. McNemar không so sánh độ chính xác tổng quát, mà tập trung kiểm định ma trận bất đồng: gồm $b$ (số mẫu ResNet50 đoán đúng nhưng mô hình đối chứng đoán sai) và $c$ (số mẫu ResNet50 đoán sai nhưng mô hình đối chứng đoán đúng).*
  > 
  > *Giá trị $p = 0,000070$ khi so với DenseNet và $p = 0,000025$ khi so với InceptionV3 đều nhỏ hơn rất nhiều so với ngưỡng quy ước 0,05. Điều này bác bỏ giả thuyết vô hiệu ($H_0$), chứng minh rằng sự vượt trội của ResNet50 là khác biệt có ý nghĩa thống kê bản chất, chứ không phải ngẫu nhiên."*

---

### CÂU HỎI 5: "Trong y học cứu người là trên hết, InceptionV3 có Độ nhạy (Sensitivity) 100%, tức không bỏ sót ca nào. Tại sao nhóm lại chọn ResNet50?"
- **Người trả lời ưu tiên:** Nguyễn Đình Duy.
- **Chiến lược trả lời:** Sử dụng luận điểm "Đánh đổi sai số lâm sàng" và phân tích chi phí y tế (Health Economics).
- **Nội dung trả lời:**
  > *"Dạ em xin cảm ơn Thầy/Cô, đây là một câu hỏi rất hay chạm đúng bản chất ứng dụng của AI y tế.*
  > 
  > *Đúng là trong sàng lọc, độ nhạy cao là ưu tiên hàng đầu. Tuy nhiên, ResNet50 cũng đã đạt độ nhạy cực kỳ cao là **99,48%**, chỉ lọt đúng 2 ca trong 387 ca bệnh.*
  > 
  > *Nếu chuyển sang InceptionV3 để cứu 2 ca này, chúng ta phải trả một cái giá quá đắt là **tăng thêm 30 ca dương tính giả** (từ 56 ca vọt lên 86 ca, độ đặc hiệu rơi từ 75,76% xuống 62,77%).*
  > 
  > *Trong quy trình lâm sàng:*
  > * 2 ca sót của ResNet50 vẫn nằm trong vòng giám sát y tế, vì bác sĩ không bao giờ chỉ dựa vào một tấm ảnh X-quang đơn lẻ mà luôn kết hợp với nghe phổi bằng ống nghe, đo nồng độ SpO2 và thân nhiệt bệnh nhi. Nếu bệnh nhi có triệu chứng, bác sĩ vẫn sẽ theo dõi.*
  > * Trong khi đó, việc dán nhãn nhầm 86 đứa trẻ khỏe mạnh là viêm phổi sẽ làm tê liệt phòng chụp CT, gây quá tải bác sĩ đọc kết quả, và đẩy 86 đứa trẻ vào liệu trình kháng sinh phổ rộng không cần thiết.*
  > 
  > *Vì vậy, ResNet50 mang lại sự cân bằng hài hòa, an toàn và khả thi nhất cho việc tích hợp vào hệ thống PACS của bệnh viện."*

---

### CÂU HỎI 6: "Mô hình Cổng chú ý không gian (Spatial Attention Gate) trong báo cáo của các em hoạt động như thế nào và đem lại lợi ích gì?"
- **Người trả lời ưu tiên:** Nguyễn Đình Duy (hoặc Mai Phước Thịnh).
- **Chiến lược trả lời:** Trình bày cấu trúc toán học của Spatial Attention và kết quả định lượng Grad-CAM.
- **Nội dung trả lời:**
  > *"Dạ em xin cảm ơn Thầy/Cô. Trong báo cáo toàn văn và bài báo, nhóm đã phát triển một kiến trúc cải tiến gọi là **ResNet50-Attn** tích hợp Cổng chú ý không gian:*
  > 
  > *1. **Nguyên lý hoạt động:** Tại lớp đặc trưng cuối cùng `layer4` của ResNet50 với kích thước $C \times H \times W$ ($2048 \times 7 \times 7$), chúng em áp dụng phép gom kênh (Channel Pooling) bằng cả Average Pooling và Max Pooling, tạo ra 2 bản đồ không gian kích thước $1 \times 7 \times 7$.*
  > *2. Hai bản đồ này được ghép kênh và đưa qua một lớp tích chập duy nhất kích thước $7 \times 7$ với hàm kích hoạt Sigmoid để sinh ra mặt nạ trọng số không gian $M_s \in [0, 1]^{7 \times 7}$.*
  > *3. Mặt nạ này được nhân phần tử (element-wise multiplication) với bản đồ đặc trưng gốc, giúp khuếch đại các vùng chứa nhu mô phổi và dập tắt hoàn toàn các tín hiệu nhiễu từ viền đen bên ngoài, xương đòn hay các chữ cái đánh dấu của kỹ thuật viên.*
  > 
  > *Khối này chỉ thêm vỏn vẹn **98 tham số** nhưng đem lại bước đột phá về tính minh bạch: Mật độ năng lượng Grad-CAM tập trung vào đúng tổn thương trong phổi tăng thêm **+13,54%**, giúp bác sĩ tin tưởng tuyệt đối vào cơ sở chẩn đoán của mô hình."*

---

### CÂU HỎI 7: "Tại sao DenseNet121 có số lượng tham số nhỏ nhất (6,95M so với 23,5M của ResNet50) nhưng thời gian huấn luyện lại lâu hơn ResNet50?"
- **Người trả lời ưu tiên:** Mai Phước Thịnh.
- **Chiến lược trả lời:** Phân biệt giữa "Số lượng tham số" (Parameters) và "Băng thông bộ nhớ / Truy cập bộ nhớ GPU" (Memory Bandwidth & Tensor Concatenation).
- **Nội dung trả lời:**
  > *"Dạ em xin cảm ơn Thầy/Cô. Đây là một hiện tượng kỹ thuật rất kinh điển giữa chi phí tính toán lý thuyết và hiệu năng phần cứng thực tế:*
  > 
  > *1. DenseNet121 tiết kiệm tham số nhờ cơ chế tái sử dụng đặc trưng, nhưng nó kết nối các lớp bằng phép **ghép kênh (Channel Concatenation)**: $[x_0, x_1, x_2, ...]$.*
  > *2. Trong GPU, mỗi lần thực hiện phép Concatenation, hệ thống bắt buộc phải cấp phát các khối bộ nhớ VRAM liên tục mới và sao chép toàn bộ tensor cũ sang địa chỉ mới. Thao tác này tiêu tốn rất nhiều băng thông bộ nhớ (Memory Access Cost - MAC).*
  > *3. Ngược lại, ResNet50 sử dụng phép **cộng tắt (Identity Addition)**: $x + F(x)$. Phép cộng này được tối ưu hóa trực tiếp ngay tại chỗ (in-place operation), không phải tái cấp phát bộ nhớ.*
  > 
  > *Vì vậy, trên phần cứng GPU thực tế, ResNet50 thực hiện các phép toán song song mượt mà hơn và hoàn thành huấn luyện nhanh hơn DenseNet121 tới hơn 3 phút."*

---

### CÂU HỎI 8: "Nếu triển khai mô hình này vào thực tế bệnh viện, nhóm dự kiến quy trình tích hợp như thế nào?"
- **Người trả lời ưu tiên:** Cả hai thành viên.
- **Chiến lược trả lời:** Nêu giải pháp triển khai API chuẩn DICOM, PACS và vai trò Second-reader (Người đọc thứ hai).
- **Nội dung trả lời:**
  > *"Dạ em xin cảm ơn Thầy/Cô. Nhóm định hướng hệ thống PneumoGuard-AI hoạt động theo mô hình **Second-Reader (Người trợ lý đồng hành)** chứ không thay thế bác sĩ:*
  > 
  > *1. **Tích hợp:** Mô hình được đóng gói dưới dạng Docker container microservice, kết nối trực tiếp với máy chủ lưu trữ hình ảnh **PACS** của bệnh viện qua giao thức chuẩn **DICOM**.*
  > *2. **Quy trình hoạt động:** Ngay khi máy X-quang chụp xong, ảnh DICOM được gửi tự động đến mô hình. Mô hình tiền xử lý, chạy suy luận trong vòng dưới 0,5 giây, và trả về 2 kết quả:*
  >   * Xác suất dự đoán viêm phổi kèm nhãn phân loại.
  >   * Bản đồ nhiệt Grad-CAM phủ màu trực tiếp lên ảnh phổi để bác sĩ quan sát ngay vị trí nghi ngờ thâm nhiễm.*
  > *3. **Phân luồng khẩn cấp (Triage):** Với những ca có xác suất viêm phổi trên 90%, hệ thống sẽ đẩy ca đó lên đầu danh sách chờ của bác sĩ trực cấp cứu để can thiệp kịp thời."*

---

# LỜI KHUYÊN DÀNH CHO HAI BẠN KHI LÊN BÁO CÁO
1. **Tác phong & Phân chia:** Đứng thẳng, hướng mắt về Hội đồng và màn chiếu, luân chuyển micro nhịp nhàng giữa Thịnh và Duy. Khi một bạn nói, bạn còn lại chú ý theo dõi slide hoặc hỗ trợ bấm slide.
2. **Ngôn ngữ cơ thể:** Dùng que chỉ hoặc con trỏ chuột laser chỉ thẳng vào các con số mấu chốt trên slide (ví dụ: con số 90,61%, p-value 0,000070, +30 FP, Epoch 8).
3. **Tâm thế khi phản biện:** Luôn bắt đầu bằng câu *"Dạ, chúng em xin cảm ơn câu hỏi rất hay của Thầy/Cô ạ!"*, trả lời dứt khoát, tự tin, không vòng vo, dẫn chứng số liệu thực tế từ bài báo và báo cáo.

