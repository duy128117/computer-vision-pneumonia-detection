import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import streamlit as st
from PIL import Image
from src.models import load_checkpoint
from src.inference import validate_image, predict
from src.utils import device

ROOT = Path(__file__).resolve().parents[1]
st.set_page_config(page_title="Pneumonia • X-ray", page_icon="🫁", layout="wide")
st.title("Phân loại viêm phổi từ X-quang ngực")
st.caption("ResNet50 · DenseNet121 · Grad-CAM")
st.info(
    "Demo nghiên cứu, không thay thế chẩn đoán lâm sàng. Grad-CAM không chứng minh mô hình suy luận giống bác sĩ."
)
name = st.sidebar.selectbox("Mô hình", ["resnet50", "densenet121"])
checkpoint = ROOT / "checkpoints" / f"best_{name}.pth"
st.sidebar.caption("NORMAL = 0 · PNEUMONIA = 1 | Ngưỡng 0.5")
if not checkpoint.exists():
    st.warning(f"Chưa có checkpoint đã huấn luyện cho {name}. Hãy chạy pipeline trước khi dự đoán.")
    st.stop()
upload = st.file_uploader("Tải ảnh X-quang ngực", type=["jpg", "jpeg", "png"])
st.caption("Kiểm tra định dạng ảnh; hệ thống chưa có bộ phát hiện ảnh ngoài miền X-quang.")
if upload is not None:
    try:
        if upload.size > 20 * 1024 * 1024:
            raise ValueError("Dung lượng tối đa 20 MB.")
        image = validate_image(upload)
        with st.spinner("Đang dự đoán và tính Grad-CAM…"):
            model, metadata = load_checkpoint(checkpoint, device())
            result, images = predict(model, metadata, image)
        st.subheader("Viêm phổi" if result["prediction"] == "PNEUMONIA" else "Bình thường")
        columns = st.columns(3)
        for column, key, label in zip(
            columns,
            ["confidence", "normal", "pneumonia"],
            ["Độ tin cậy dự đoán", "NORMAL", "PNEUMONIA"],
        ):
            column.metric(label, f"{result[key]:.2%}")
        st.caption("Xác suất đầu ra mô hình, chưa được hiệu chỉnh xác suất lâm sàng.")
        for column, picture, label in zip(
            st.columns(3),
            (image, images[1], images[2]),
            ["Ảnh X-quang gốc", "Grad-CAM Heatmap", "Grad-CAM Overlay"],
        ):
            column.image(picture, caption=label, use_container_width=True)
    except (ValueError, OSError, RuntimeError, Image.DecompressionBombError) as error:
        st.error(str(error))
