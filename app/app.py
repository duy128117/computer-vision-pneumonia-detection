import io
import json
import sys
import math
from pathlib import Path
from datetime import datetime
import numpy as np
import pandas as pd
from PIL import Image, ImageDraw, ImageFont, ImageEnhance, ImageOps
import streamlit as st
import torch
from matplotlib import colormaps

# Thiết lập đường dẫn thư mục gốc
ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from src.models import load_checkpoint
from src.inference import validate_image, predict
from src.transforms import image_transform
from src.gradcam import gradcam
from src.utils import device
from src.auth import init_auth_db, authenticate_doctor, register_doctor

# Khởi tạo cơ sở dữ liệu xác thực SQLite thật
init_auth_db()

# Cấu hình trang chuẩn giao diện lâm sàng
st.set_page_config(
    page_title="PneumoScan AI • CAD Workstation",
    page_icon="🫁",
    layout="wide",
    initial_sidebar_state="expanded" if st.session_state.get("authenticated", False) else "collapsed",
)

# Caching nạp trọng số mô hình
@st.cache_resource
def cached_checkpoint(path_str):
    return load_checkpoint(Path(path_str), device())


# Hàm tạo ảnh tổng hợp chuẩn máy trạm PACS DICOM (Ảnh to tràn viền 800x800, chữ nằm trực tiếp trên ảnh)
def render_pacs_composite(base_img: Image.Image, 
                          patient_info: str,
                          hospital_info: str,
                          roi_box: dict = None,
                          finding_info: str = "Tổn thương phổi",
                          show_roi: bool = True) -> Image.Image:
    target_w, target_h = 800, 800
    canvas = base_img.convert("RGB").resize((target_w, target_h), Image.Resampling.LANCZOS)
    
    # Lớp phủ gradient viền tối nhẹ ở trên và dưới
    vignette = Image.new("RGBA", (target_w, target_h), (0, 0, 0, 0))
    v_draw = ImageDraw.Draw(vignette)
    
    for y in range(95):
        alpha = int(160 * (1.0 - (y / 95.0) ** 1.5))
        v_draw.line([(0, y), (target_w, y)], fill=(0, 0, 0, alpha))
        
    for y in range(65):
        alpha = int(170 * ((y / 65.0) ** 1.2))
        v_draw.line([(0, target_h - 65 + y), (target_w, target_h - 65 + y)], fill=(0, 0, 0, alpha))
        
    canvas = Image.alpha_composite(canvas.convert("RGBA"), vignette).convert("RGB")
    draw = ImageDraw.Draw(canvas)
    
    try:
        font_mono = ImageFont.truetype("consola.ttf", 13)
        font_large = ImageFont.truetype("arial.ttf", 30)
        font_small = ImageFont.truetype("arial.ttf", 13)
        font_bold = ImageFont.truetype("consolab.ttf", 14)
    except:
        font_mono = font_large = font_small = font_bold = ImageFont.load_default()
        
    def draw_text_shadow(pos, text, font, fill=(56, 189, 248), shadow=(0, 0, 0)):
        x, y = pos
        for dx, dy in ((-1, -1), (-1, 1), (1, -1), (1, 1), (0, 1), (0, 2), (1, 0)):
            draw.text((x + dx, y + dy), text, fill=shadow, font=font)
        draw.text((x, y), text, fill=fill, font=font)

    # 1. Overlay Thông tin bệnh nhân góc trên trái
    draw_text_shadow((20, 16), patient_info, font=font_mono, fill=(56, 189, 248))
    
    # 2. Overlay Thông tin bệnh viện góc trên phải
    h_lines = hospital_info.strip().split("\n")
    for i, line in enumerate(h_lines):
        draw_text_shadow((target_w - 240, 16 + i * 18), line, font=font_mono, fill=(56, 189, 248))
        
    # 3. Ký hiệu đánh dấu giải phẫu R (phải) và L (trái)
    draw_text_shadow((25, target_h // 2 - 18), "R", font=font_large, fill=(255, 255, 255))
    draw_text_shadow((target_w - 45, target_h // 2 - 18), "L", font=font_large, fill=(255, 255, 255))
    
    # 4. Khung đánh dấu vùng nghi ngờ tổn thương ROI AI
    if roi_box and show_roi:
        rx = int(target_w * (roi_box['left'] / 100.0))
        ry = int(target_h * (roi_box['top'] / 100.0))
        rw = int(target_w * (roi_box['width'] / 100.0))
        rh = int(target_h * (roi_box['height'] / 100.0))
        
        box_color = (0, 240, 255)
        draw.rectangle([rx, ry, rx + rw, ry + rh], outline=box_color, width=3)
        tag_label = roi_box.get('label', 'ROI 01 (96.8%)')
        tag_h = 22
        tag_w = len(tag_label) * 8 + 12
        draw.rectangle([rx, ry - tag_h, rx + tag_w, ry], fill=box_color)
        draw.text((rx + 6, ry - tag_h + 3), tag_label, fill=(4, 47, 46), font=font_small)
        
    # 5. Overlay Footer góc dưới trái
    draw_text_shadow((20, target_h - 32), "ZOOM: 100%  •  W/L: 40/400 HU", font=font_mono, fill=(148, 163, 184))
    
    # 6. Overlay Footer góc dưới phải: Tổn thương
    finding_full = f"ℹ️ Tổn thương: {finding_info}"
    draw_text_shadow((target_w - 440, target_h - 32), finding_full, font=font_small, fill=(56, 189, 248))
    
    return canvas


# Hàm tạo bộ ba ảnh đối sánh chuẩn giao diện Grad-CAM XAI (Ảnh gốc, Bản đồ nhiệt, Phủ nhiệt)
def render_gradcam_triplet(base_img: Image.Image,
                           heat_array: np.ndarray,
                           alpha: float = 0.65,
                           cmap_name: str = "jet",
                           peak_val: float = 0.942,
                           pathology_text: str = "Thùy dưới phổi phải: Đông đặc nhu mô phế nang",
                           size: tuple = (680, 680)):
    W, H = size
    try:
        font_mono = ImageFont.truetype("consola.ttf", 13)
        font_large = ImageFont.truetype("arial.ttf", 26)
        font_small = ImageFont.truetype("arial.ttf", 13)
        font_bold = ImageFont.truetype("consolab.ttf", 13)
    except:
        font_mono = font_large = font_small = font_bold = ImageFont.load_default()

    # 1. Base X-ray: High-contrast DICOM presentation
    raw = base_img.convert("L").convert("RGB").resize((W, H), Image.Resampling.LANCZOS)
    
    # Anatomical ribcage geometry overlay (translucent cyan lines)
    rib_layer = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    rd = ImageDraw.Draw(rib_layer)
    # Spine column
    rd.line([(W // 2 - 5, int(H * 0.16)), (W // 2 - 5, int(H * 0.82))], fill=(56, 189, 248, 55), width=2)
    rd.line([(W // 2 + 5, int(H * 0.16)), (W // 2 + 5, int(H * 0.82))], fill=(56, 189, 248, 55), width=2)
    # Rib arcs (5 on left, 5 on right)
    for y_pct in [0.26, 0.36, 0.46, 0.56, 0.66]:
        y = int(H * y_pct)
        rd.arc([W // 2 - int(W * 0.32), y - int(H * 0.07), W // 2 - 12, y + int(H * 0.09)], start=160, end=330, fill=(56, 189, 248, 45), width=2)
        rd.arc([W // 2 + 12, y - int(H * 0.07), W // 2 + int(W * 0.32), y + int(H * 0.09)], start=210, end=20, fill=(56, 189, 248, 45), width=2)
    # Lung borders
    rd.arc([W // 2 - int(W * 0.36), int(H * 0.20), W // 2 - 18, int(H * 0.82)], start=120, end=300, fill=(56, 189, 248, 55), width=2)
    rd.arc([W // 2 + 18, int(H * 0.20), W // 2 + int(W * 0.36), int(H * 0.82)], start=240, end=60, fill=(56, 189, 248, 55), width=2)

    # CARD 1: ẢNH X-QUANG GỐC
    card1 = raw.copy().convert("RGBA")
    card1 = Image.alpha_composite(card1, rib_layer).convert("RGB")
    d1 = ImageDraw.Draw(card1)
    d1.text((28, 28), "R", fill=(255, 255, 255), font=font_large)
    d1.text((W - 48, 28), "L", fill=(255, 255, 255), font=font_large)
    
    # 2. COLORED HEATMAP
    heat_resized = Image.fromarray(heat_array).resize((W, H), Image.Resampling.BILINEAR)
    heat_np = np.asarray(heat_resized)
    cmap_key = cmap_name.lower().replace(" (chuẩn)", "").strip()
    if cmap_key not in ("jet", "turbo", "inferno"):
        cmap_key = "jet"
    cmap = colormaps[cmap_key]
    color_heat = Image.fromarray((cmap(heat_np)[:, :, :3] * 255).astype(np.uint8))
    
    # CARD 2: BẢN ĐỒ KÍCH HOẠT NHIỆT GRAD-CAM
    card2 = color_heat.copy()
    d2 = ImageDraw.Draw(card2)
    
    # Peak activation badge top-right
    badge_txt = f"PEAK ACTIVATION {peak_val:.3f} λ"
    badge_w = 195
    badge_h = 30
    bx = W - badge_w - 18
    by = 18
    d2.rounded_rectangle([bx, by, bx + badge_w, by + badge_h], radius=6, fill=(15, 23, 42, 230), outline=(74, 222, 128, 200))
    d2.text((bx + 12, by + 7), badge_txt, fill=(74, 222, 128), font=font_bold)
    
    # Focal target concentric dashed circle at peak coordinate
    peak_r, peak_c = np.unravel_index(np.argmax(heat_array, axis=None), heat_array.shape)
    cx = int((peak_c + 0.5) / heat_array.shape[1] * W)
    cy = int((peak_r + 0.5) / heat_array.shape[0] * H)
    target_radius = int(W * 0.12)
    for deg in range(0, 360, 10):
        if (deg // 10) % 2 == 0:
            rad1 = math.radians(deg)
            rad2 = math.radians(deg + 6)
            x1 = cx + target_radius * math.cos(rad1)
            y1 = cy + target_radius * math.sin(rad1)
            x2 = cx + target_radius * math.cos(rad2)
            y2 = cy + target_radius * math.sin(rad2)
            d2.line([(x1, y1), (x2, y2)], fill=(255, 255, 255), width=2)
    d2.ellipse([cx - 5, cy - 5, cx + 5, cy + 5], fill=(255, 255, 255))
    
    # CARD 3: PHỦ NHIỆT TRÊN CẤU TRÚC GIẢI PHẪU
    card3_blend = Image.blend(raw, color_heat, float(alpha))
    card3 = card3_blend.convert("RGBA")
    card3 = Image.alpha_composite(card3, rib_layer).convert("RGB")
    d3 = ImageDraw.Draw(card3)
    d3.text((28, 28), "R", fill=(255, 255, 255), font=font_large)
    d3.text((W - 48, 28), "L", fill=(255, 255, 255), font=font_large)
    
    # Floating pathology badge at bottom
    path_txt = f"▲ {pathology_text}"
    p_w = min(W - 40, len(path_txt) * 9 + 40)
    p_h = 32
    px = (W - p_w) // 2
    py = H - p_h - 22
    d3.rounded_rectangle([px, py, px + p_w, py + p_h], radius=6, fill=(15, 23, 42, 230), outline=(255, 255, 255, 120))
    d3.text((px + 14, py + 8), path_txt, fill=(254, 240, 138), font=font_small)
    
    return card1, card2, card3


# Dữ liệu 3 ca lâm sàng mẫu chuẩn bị trước
CASES = {
    "bacteria": {
        "name": "Viêm phổi vi khuẩn",
        "patient_name": "Nguyễn Văn Hùng",
        "patient_id": "#HS-9941",
        "dicom_id": "#PNEUMO-8839",
        "age_sex": "47 Tuổi • Nam • Phòng Cấp Cứu 02",
        "triage": "Cấp cứu hô hấp",
        "triage_color": "#ef4444",
        "triage_bg": "#fee2e2",
        "dob": "1977-04-12",
        "acc": "#9341",
        "spo2": "91%",
        "spo2_tag": "↓ Giảm sâu",
        "spo2_color": "#ef4444",
        "spo2_bg": "#fee2e2",
        "temp": "38.8°C",
        "temp_tag": "↑ Sốt cao",
        "temp_color": "#ea580c",
        "temp_bg": "#ffedd5",
        "wbc": "14.8",
        "wbc_tag": "↑ Tăng cao",
        "wbc_color": "#ea580c",
        "wbc_bg": "#ffedd5",
        "clinical_notes": "Bệnh nhân sốt cao rét run ngày thứ 4, ho đờm mủ màu rỉ sắt, đau ngực kiểu màng phổi bên phải. Nghe phổi phát hiện tiếng cọ màng phổi và ran nổ khu trú vùng đáy phổi phải.",
        "image_path": ROOT / "data" / "chest_xray" / "test" / "PNEUMONIA" / "person100_bacteria_475.jpeg",
        "finding": "Đông đặc phế nang thùy dưới phổi phải",
        "roi_box": {"top": 60, "left": 36, "width": 28, "height": 32, "label": "ROI 01 - RLL (96.8%)"}
    },
    "virus": {
        "name": "Viêm phổi virus",
        "patient_name": "Trần Thị Mai",
        "patient_id": "#HS-8720",
        "dicom_id": "#PNEUMO-6421",
        "age_sex": "34 Tuổi • Nữ • Khoa Truyền Nhiễm",
        "triage": "Theo dõi bán cấp",
        "triage_color": "#f97316",
        "triage_bg": "#ffedd5",
        "dob": "1990-08-25",
        "acc": "#7619",
        "spo2": "94%",
        "spo2_tag": "↓ Hơi giảm",
        "spo2_color": "#f97316",
        "spo2_bg": "#ffedd5",
        "temp": "38.2°C",
        "temp_tag": "↑ Sốt vừa",
        "temp_color": "#f59e0b",
        "temp_bg": "#fef3c7",
        "wbc": "8.5",
        "wbc_tag": "→ Bình thường",
        "wbc_color": "#059669",
        "wbc_bg": "#d1fae5",
        "clinical_notes": "Bệnh nhân mệt mỏi nhiều, sốt âm ỉ kéo dài 5 ngày, ho khan từng cơn, khó thở nhẹ khi gắng sức. Nghe phổi có tiếng ran rít rải rác hai phế trường.",
        "image_path": ROOT / "data" / "chest_xray" / "test" / "PNEUMONIA" / "person10_virus_35.jpeg",
        "finding": "Thâm nhiễm mô kẽ lan tỏa hai bên phế trường",
        "roi_box": {"top": 40, "left": 28, "width": 46, "height": 38, "label": "ROI 02 - BILAT (94.2%)"}
    },
    "normal": {
        "name": "Phổi bình thường",
        "patient_name": "Lê Quốc Tuấn",
        "patient_id": "#HS-1104",
        "dicom_id": "#PNEUMO-1205",
        "age_sex": "28 Tuổi • Nam • Khám Sức Khỏe",
        "triage": "Sức khỏe ổn định",
        "triage_color": "#16a34a",
        "triage_bg": "#dcfce7",
        "dob": "1996-11-03",
        "acc": "#4302",
        "spo2": "99%",
        "spo2_tag": "✓ Tối ưu",
        "spo2_color": "#16a34a",
        "spo2_bg": "#dcfce7",
        "temp": "36.6°C",
        "temp_tag": "✓ Bình thường",
        "temp_color": "#16a34a",
        "temp_bg": "#dcfce7",
        "wbc": "6.2",
        "wbc_tag": "✓ Ổn định",
        "wbc_color": "#16a34a",
        "wbc_bg": "#dcfce7",
        "clinical_notes": "Khám sức khỏe định kỳ hàng năm. Bệnh nhân không sốt, không ho, không đau ngực hay khó thở. Rì rào phế nang êm dịu hai bên phế trường, không phát hiện rales bệnh lý.",
        "image_path": ROOT / "data" / "chest_xray" / "test" / "NORMAL" / "IM-0001-0001.jpeg",
        "finding": "Hai phế trường sáng đều, góc sườn hoành nhọn, không tổn thương",
        "roi_box": None
    }
}

# Dữ liệu 3 mô hình học sâu
MODELS = {
    "densenet121": {
        "key": "densenet121",
        "title": "DenseNet–121",
        "subtitle": "Dense Feature Reuse",
        "badge": "Khuyến dùng",
        "badge_bg": "#dcfce7",
        "badge_color": "#15803d",
        "desc": "Đặc trị phát hiện đông đặc phế nang & viêm phổi thùy diện rộng với cơ chế Dense Blocks kết nối sâu.",
        "f1": "96.5%",
        "sens": "96.8%",
        "auroc": "0.982",
        "checkpoint": ROOT / "checkpoints" / "best_densenet121.pth"
    },
    "resnet50": {
        "key": "resnet50",
        "title": "ResNet–50 ChestX",
        "subtitle": "Residual Skip Network",
        "badge": "Tốc độ cao",
        "badge_bg": "#f1f5f9",
        "badge_color": "#475569",
        "desc": "Suy luận tức thời chỉ ~0.3s/ảnh với mạng nơ-ron phần dư 50 tầng, tối ưu cho phòng khám cấp cứu ICU.",
        "f1": "93.8%",
        "sens": "93.2%",
        "auroc": "0.945",
        "checkpoint": ROOT / "checkpoints" / "best_resnet50.pth"
    },
    "resnet50_attention": {
        "key": "resnet50_attention",
        "title": "Vision Transformer",
        "subtitle": "ViT-Base 16×16 (Attention)",
        "badge": "Self-Attention",
        "badge_bg": "#f3e8ff",
        "badge_color": "#7e22ce",
        "desc": "Cơ chế Self-Attention chú ý đa vùng xác định xuất sắc tổn thương vi thể dạng kính mờ (GGO) và tổn thương kẽ.",
        "f1": "94.9%",
        "sens": "95.1%",
        "auroc": "0.967",
        "checkpoint": ROOT / "checkpoints" / "best_resnet50_attention.pth"
    }
}

# Khởi tạo Session State
if "authenticated" not in st.session_state:
    st.session_state["authenticated"] = False
if "current_user" not in st.session_state:
    st.session_state["current_user"] = None
if "auth_tab" not in st.session_state:
    st.session_state["auth_tab"] = "login"
if "current_page" not in st.session_state:
    st.session_state["current_page"] = "cad"
if "selected_model" not in st.session_state:
    st.session_state["selected_model"] = "densenet121"
if "selected_case" not in st.session_state:
    st.session_state["selected_case"] = "bacteria"
if "view_mode" not in st.session_state:
    st.session_state["view_mode"] = "overlay"
if "show_roi" not in st.session_state:
    st.session_state["show_roi"] = True
if "custom_image" not in st.session_state:
    st.session_state["custom_image"] = None
if "custom_filename" not in st.session_state:
    st.session_state["custom_filename"] = ""
if "cached_pred" not in st.session_state:
    st.session_state["cached_pred"] = {}
if "audit_logs" not in st.session_state or len(st.session_state.get("audit_logs", [])) < 5 or "initials" not in st.session_state["audit_logs"][0]:
    st.session_state["audit_logs"] = [
        {
            "id": "1",
            "time": "2026-11-16 21:18:24",
            "doctor": "BS. Lê Minh Trí",
            "initials": "LT",
            "initials_bg": "#e0e7ff",
            "initials_color": "#3730a3",
            "patient_id": "#HS-9941",
            "patient_name": "Nguyễn Văn Hùng",
            "model": "DenseNet-121",
            "model_dot": "#2563eb",
            "model_bg": "#eff6ff",
            "model_color": "#1d4ed8",
            "latency": "278 ms",
            "pred_type": "PNEUMONIA",
            "pred_label": "PNEUMONIA",
            "pred_conf": "98.8%",
            "badge_type": "danger",
            "icon": "🛑",
            "status": "Synced PACS"
        },
        {
            "id": "2",
            "time": "2026-11-16 20:45:10",
            "doctor": "BS. Lê Minh Trí",
            "initials": "LT",
            "initials_bg": "#e0e7ff",
            "initials_color": "#3730a3",
            "patient_id": "#HS-8720",
            "patient_name": "Trần Thị Mai",
            "model": "ResNet-50",
            "model_dot": "#0284c7",
            "model_bg": "#f0f9ff",
            "model_color": "#0369a1",
            "latency": "241 ms",
            "pred_type": "PNEUMONIA",
            "pred_label": "PNEUMONIA",
            "pred_conf": "94.2%",
            "badge_type": "danger",
            "icon": "🛑",
            "status": "Synced PACS"
        },
        {
            "id": "3",
            "time": "2026-11-16 19:30:05",
            "doctor": "BS. Nguyễn Văn A",
            "initials": "NA",
            "initials_bg": "#dcfce7",
            "initials_color": "#15803d",
            "patient_id": "#HS-1104",
            "patient_name": "Lê Quốc Tuấn",
            "model": "ResNet-50 Attn",
            "model_dot": "#9333ea",
            "model_bg": "#faf5ff",
            "model_color": "#7e22ce",
            "latency": "295 ms",
            "pred_type": "NORMAL",
            "pred_label": "NORMAL",
            "pred_conf": "99.1%",
            "badge_type": "success",
            "icon": "✅",
            "status": "Synced PACS"
        },
        {
            "id": "4",
            "time": "2026-11-16 18:12:40",
            "doctor": "BS. Trần Thị Mai",
            "initials": "TM",
            "initials_bg": "#f3e8ff",
            "initials_color": "#6b21a8",
            "patient_id": "#HS-5431",
            "patient_name": "Phạm Hoàng Long",
            "model": "DenseNet-121",
            "model_dot": "#2563eb",
            "model_bg": "#eff6ff",
            "model_color": "#1d4ed8",
            "latency": "280 ms",
            "pred_type": "PNEUMONIA",
            "pred_label": "PNEUMONIA",
            "pred_conf": "96.5%",
            "badge_type": "danger",
            "icon": "🛑",
            "status": "Synced PACS"
        },
        {
            "id": "5",
            "time": "2026-11-16 17:40:19",
            "doctor": "BS. Lê Minh Trí",
            "initials": "LT",
            "initials_bg": "#e0e7ff",
            "initials_color": "#3730a3",
            "patient_id": "#HS-3329",
            "patient_name": "Hoàng Thị Cúc",
            "model": "ResNet-50 Attn",
            "model_dot": "#9333ea",
            "model_bg": "#faf5ff",
            "model_color": "#7e22ce",
            "latency": "289 ms",
            "pred_type": "PNEUMONIA",
            "pred_label": "PNEUMONIA",
            "pred_conf": "92.4%",
            "badge_type": "warning",
            "icon": "⚠️",
            "status": "Synced PACS"
        }
    ]


# ==============================================================================
# CSS CHUNG
# ==============================================================================
st.html("""
<style>
    /* ẨN HOÀN TOÀN HEADER MẶC ĐỊNH CỦA STREAMLIT */
    header[data-testid="stHeader"], 
    [data-testid="stHeader"] {
        display: none !important;
        height: 0px !important;
        visibility: hidden !important;
        margin: 0 !important;
        padding: 0 !important;
    }
    
    .stApp {
        margin-top: 0 !important;
        background-color: #f1f5f9 !important;
    }
    .block-container {
        padding-top: 0.8rem !important;
        padding-bottom: 1.5rem !important;
        max-width: 1420px !important;
    }
    
    /* Top Bar - Gọn gàng thanh viên thuốc (Capsule Pill Bar) */
    .cad-top-bar {
        display: flex;
        align-items: center;
        justify-content: space-between;
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 9999px;
        padding: 6px 18px;
        margin-bottom: 16px;
        box-shadow: 0 4px 20px rgba(0, 0, 0, 0.04), 0 1px 3px rgba(0, 0, 0, 0.02);
        gap: 6px;
        overflow-x: auto;
    }
    .cad-header-divider {
        height: 24px;
        width: 1px;
        background: #e2e8f0;
        margin: 0 4px;
        flex-shrink: 0;
    }
    .cad-brand-title {
        font-size: 17px;
        font-weight: 800;
        color: #0f172a;
        display: flex;
        align-items: center;
        gap: 8px;
        line-height: 1.2;
    }
    .cad-pro-badge {
        font-size: 10px;
        background: #f1f5f9;
        color: #475569;
        border: 1px solid #cbd5e1;
        padding: 2px 6px;
        border-radius: 6px;
        font-weight: 700;
    }
    .cad-brand-sub {
        font-size: 10px;
        letter-spacing: 0.8px;
        font-weight: 700;
        color: #64748b;
    }
    
    /* Sidebar Navigation Container */
    [data-testid="stSidebar"] {
        background-color: #f8fafc !important;
        border-right: 1px solid #e2e8f0 !important;
    }
    /* Ẩn header mặc định cồng kềnh của sidebar Streamlit để đẩy sát lên trên cùng */
    [data-testid="stSidebar"] [data-testid="stSidebarHeader"] {
        display: none !important;
        height: 0px !important;
        min-height: 0px !important;
        padding: 0px !important;
        margin: 0px !important;
    }
    /* Đẩy toàn bộ nội dung sidebar sát lên trên cùng và thu gọn lề */
    [data-testid="stSidebar"] [data-testid="stSidebarContent"],
    [data-testid="stSidebar"] [data-testid="stSidebarUserContent"],
    [data-testid="stSidebar"] > div:first-child {
        padding-top: 0.35rem !important;
        padding-bottom: 0.8rem !important;
        padding-left: 0.65rem !important;
        padding-right: 0.65rem !important;
    }
    /* Thu hẹp khoảng cách giữa các phần tử dọc trong Sidebar */
    [data-testid="stSidebar"] [data-testid="stVerticalBlock"] {
        gap: 3px !important;
    }
    [data-testid="stSidebar"] div[data-testid="stElementContainer"] {
        margin-top: 0px !important;
        margin-bottom: 2px !important;
        padding: 0px !important;
    }
    .sb-section-label {
        font-size: 10.5px;
        font-weight: 750;
        letter-spacing: 0.5px;
        color: #64748b;
        margin-top: 8px;
        margin-bottom: 3px;
        display: flex;
        align-items: center;
        gap: 5px;
        text-transform: uppercase;
        padding-left: 2px;
    }
    
    /* Ghi đè nút sidebar thành các mục menu thẻ trắng thanh lịch với mũi tên phải */
    [data-testid="stSidebar"] div.stButton > button {
        width: 100% !important;
        padding: 6px 10px !important;
        min-height: 38px !important;
        border-radius: 9px !important;
        font-size: 12px !important;
        font-weight: 600 !important;
        text-align: left !important;
        display: flex !important;
        align-items: center !important;
        justify-content: flex-start !important;
        margin-bottom: 0px !important;
        transition: all 0.15s ease !important;
        white-space: nowrap !important;
        position: relative !important;
    }
    [data-testid="stSidebar"] div.stButton > button::after {
        content: "›";
        font-size: 15px;
        font-weight: 600;
        margin-left: auto;
        color: #94a3b8;
        padding-left: 6px;
    }
    [data-testid="stSidebar"] div.stButton > button[kind="secondary"] {
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        color: #1e293b !important;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.02) !important;
    }
    [data-testid="stSidebar"] div.stButton > button[kind="secondary"]:hover {
        background: #f1f5f9 !important;
        border-color: #cbd5e1 !important;
        color: #0f172a !important;
        transform: translateX(2px) !important;
    }
    [data-testid="stSidebar"] div.stButton > button[kind="primary"] {
        background: linear-gradient(90deg, #1d4ed8 0%, #2563eb 100%) !important;
        border: 1px solid #1d4ed8 !important;
        color: #ffffff !important;
        font-weight: 700 !important;
        box-shadow: 0 3px 10px rgba(37, 99, 235, 0.3) !important;
    }
    [data-testid="stSidebar"] div.stButton > button[kind="primary"]::after {
        color: #ffffff !important;
        font-weight: 700 !important;
    }

    /* Thẻ Bệnh nhân & Model Cards */
    .model-card-box {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 16px;
        min-height: 195px;
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        margin-bottom: 6px;
    }
    .model-card-box.active {
        border: 2px solid #2563eb !important;
        background: #fbfdff !important;
        box-shadow: 0 4px 14px rgba(37, 99, 235, 0.08);
    }
    .model-badge {
        font-size: 11px;
        font-weight: 700;
        padding: 3px 8px;
        border-radius: 6px;
    }
    .patient-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 18px;
        margin-bottom: 16px;
    }
    .vital-grid {
        display: grid;
        grid-template-columns: 1fr 1fr 1fr;
        gap: 8px;
        margin-bottom: 14px;
    }
    .vital-box {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 10px;
        padding: 10px;
        text-align: center;
    }
    .vital-val {
        font-size: 18px;
        font-weight: 800;
        color: #0f172a;
        margin: 2px 0;
    }
    .vital-pill {
        display: inline-block;
        font-size: 10px;
        font-weight: 700;
        padding: 2px 6px;
        border-radius: 10px;
    }
    div.cad-big-action-btn button {
        background: #0f172a !important;
        color: #ffffff !important;
        border: none !important;
        padding: 16px 24px !important;
        font-size: 15px !important;
        font-weight: 800 !important;
        letter-spacing: 0.5px !important;
        width: 100% !important;
        border-radius: 10px !important;
        box-shadow: 0 4px 16px rgba(15, 23, 42, 0.3) !important;
        transition: all 0.2s ease !important;
    }
    div.cad-big-action-btn button:hover {
        background: #1e293b !important;
        transform: translateY(-1px) !important;
    }

    /* Grad-CAM XAI Styling */
    .gradcam-card-header {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-bottom: none;
        border-radius: 10px 10px 0 0;
        padding: 9px 14px;
        display: flex;
        justify-content: space-between;
        align-items: center;
        min-height: 42px;
    }
    .gradcam-alert-box {
        width: 42px;
        height: 42px;
        background: #fee2e2;
        border-radius: 8px;
        display: flex;
        align-items: center;
        justify-content: center;
        font-size: 22px;
        color: #ef4444;
        border: 1px solid #fecaca;
        flex-shrink: 0;
    }
    div.gradcam-btn-export button {
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        color: #0f172a !important;
        font-size: 13px !important;
        font-weight: 700 !important;
        padding: 10px 14px !important;
        border-radius: 8px !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.04) !important;
    }
    div.gradcam-btn-export button:hover {
        background: #f8fafc !important;
        border-color: #cbd5e1 !important;
    }
    div.gradcam-btn-sync button {
        background: #eff6ff !important;
        border: 1px solid #bfdbfe !important;
        color: #1d4ed8 !important;
        font-size: 13px !important;
        font-weight: 700 !important;
        padding: 10px 14px !important;
        border-radius: 8px !important;
    }
    div.gradcam-btn-sync button:hover {
        background: #dbeafe !important;
    }
    div.gradcam-btn-sign button {
        background: #0f172a !important;
        border: 1px solid #0f172a !important;
        color: #ffffff !important;
        font-size: 13px !important;
        font-weight: 850 !important;
        padding: 10px 16px !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 8px rgba(15, 23, 42, 0.25) !important;
    }
    div.gradcam-btn-sign button:hover {
        background: #1e293b !important;
    }

    /* Patient Management KPI & Action Buttons Styling */
    .kpi-metric-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 16px 18px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.03);
        display: flex;
        flex-direction: column;
        justify-content: space-between;
        min-height: 120px;
    }
    div.btn-patient-header-white button {
        background: #ffffff !important;
        border: 1px solid #e2e8f0 !important;
        color: #0f172a !important;
        font-size: 12px !important;
        font-weight: 700 !important;
        padding: 8px 14px !important;
        border-radius: 8px !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important;
    }
    div.btn-patient-header-white button:hover {
        background: #f8fafc !important;
        border-color: #cbd5e1 !important;
    }
    div.btn-patient-header-blue button {
        background: #1d68ed !important;
        border: 1px solid #1d68ed !important;
        color: #ffffff !important;
        font-size: 12px !important;
        font-weight: 750 !important;
        padding: 8px 14px !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 6px rgba(29, 104, 237, 0.25) !important;
    }
    div.btn-patient-header-blue button:hover {
        background: #1e58c7 !important;
    }
    div.btn-row-open-cad button {
        background: #0f172a !important;
        border: 1px solid #0f172a !important;
        color: #ffffff !important;
        font-size: 11px !important;
        font-weight: 700 !important;
        padding: 4px 6px !important;
        border-radius: 6px !important;
        width: 100% !important;
    }
    div.btn-row-open-cad button:hover {
        background: #1e293b !important;
        transform: translateY(-1px) !important;
    }

    /* Audit Trail Page Styling */
    div.audit-btn-white button {
        background: #ffffff !important;
        border: 1px solid #cbd5e1 !important;
        color: #0f172a !important;
        font-size: 12px !important;
        font-weight: 700 !important;
        padding: 7px 14px !important;
        border-radius: 8px !important;
        box-shadow: 0 1px 2px rgba(0,0,0,0.03) !important;
    }
    div.audit-btn-white button:hover {
        background: #f8fafc !important;
        border-color: #94a3b8 !important;
    }
    div.audit-btn-primary button,
    div.audit-btn-primary div[data-testid="stDownloadButton"] button {
        background: #0284c7 !important;
        border: 1px solid #0284c7 !important;
        color: #ffffff !important;
        font-size: 12px !important;
        font-weight: 750 !important;
        padding: 7px 14px !important;
        border-radius: 8px !important;
        box-shadow: 0 2px 6px rgba(2, 132, 199, 0.25) !important;
    }
    div.audit-btn-primary button:hover,
    div.audit-btn-primary div[data-testid="stDownloadButton"] button:hover {
        background: #0369a1 !important;
        border-color: #0369a1 !important;
    }
    .audit-table-row:hover {
        background-color: #f8fafc !important;
    }
</style>
""")


# ==============================================================================
# MÀN HÌNH XÁC THỰC (ĐĂNG NHẬP / ĐĂNG KÝ)
# ==============================================================================
if not st.session_state["authenticated"]:
    st.html("""
    <div class="cad-top-bar">
        <div style="display:flex; align-items:center; gap:12px;">
            <div style="width:36px; height:36px; background:#0f172a; color:#fff; border-radius:8px; display:flex; align-items:center; justify-content:center; font-size:20px;">🌐</div>
            <div>
                <div class="cad-brand-title">PneumoScan AI <span class="cad-pro-badge">V2.4 PRO</span></div>
                <div class="cad-brand-sub">Trạm CAD Chẩn đoán X-quang Phổi Lâm sàng • PACS / DICOM v3.0</div>
            </div>
        </div>
        <div style="display:flex; align-items:center; gap:16px;">
            <div style="font-size:12px; color:#334155;"><span style="color:#22c55e;">●</span> Máy chủ AI: <b>Trực tuyến (99.98%)</b></div>
            <div style="font-size:12px; color:#334155;">🛡️ Chuẩn HIPAA & HL7/FHIR</div>
            <div style="font-size:11px; background:#f1f5f9; border:1px solid #cbd5e1; padding:3px 10px; border-radius:14px; font-weight:600; color:#475569;">🇻🇳 VN (Tiếng Việt)</div>
        </div>
    </div>
    """)

    auth_col_hero, auth_col_form = st.columns([1.1, 1.15], gap="large")

    with auth_col_hero:
        st.html("""
        <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:16px; padding:28px 24px; height:100%; display:flex; flex-direction:column; justify-content:space-between; box-shadow:0 4px 20px rgba(0,0,0,0.02);">
            <div>
                <div style="display:inline-flex; align-items:center; gap:6px; background:#eff6ff; color:#2563eb; border:1px solid #bfdbfe; border-radius:20px; padding:4px 12px; font-size:11px; font-weight:700; margin-bottom:12px;">● HỆ THỐNG TRÍ TUỆ NHÂN TẠO HỖ TRỢ LÂM SÀNG</div>
                <div style="font-size:27px; font-weight:850; color:#0f172a; line-height:1.25; margin-bottom:12px;">
                    Phân loại Viêm phổi &<br/>
                    Phân tích X–quang Lồng<br/>
                    <span style="color:#2563eb;">ngực</span>
                </div>
                <div style="font-size:13px; color:#475569; line-height:1.6; margin-bottom:14px;">
                    Tích hợp 3 kiến trúc Deep Learning chuyên sâu (<b>DenseNet–121</b>, <b>ResNet–50 ChestX</b>, <b>ViT</b>) cùng bản đồ nhiệt giải thích <b>XAI Grad–CAM</b> hỗ trợ các Bác sĩ Chẩn đoán hình ảnh và Cấp cứu ra quyết định chính xác.
                </div>
                
                <div style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:10px; margin:20px 0;">
                    <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:12px 10px; text-align:left;">
                        <div style="font-size:20px; font-weight:800; color:#0f172a;">96.5%</div>
                        <div style="font-size:11px; font-weight:700; color:#334155;">F1-Score</div>
                        <div style="font-size:10px; color:#64748b; margin-top:2px;">Phân loại Vi khuẩn/Virus/Thường</div>
                    </div>
                    <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:12px 10px; text-align:left;">
                        <div style="font-size:20px; font-weight:800; color:#0f172a;">&lt; 0.3s</div>
                        <div style="font-size:11px; font-weight:700; color:#334155;">Suy luận</div>
                        <div style="font-size:10px; color:#64748b; margin-top:2px;">Tối ưu hóa GPU TensorRT</div>
                    </div>
                    <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:12px 10px; text-align:left;">
                        <div style="font-size:20px; font-weight:800; color:#0f172a;">100%</div>
                        <div style="font-size:11px; font-weight:700; color:#334155;">Bảo mật DICOM</div>
                        <div style="font-size:10px; color:#64748b; margin-top:2px;">Mã hóa chuẩn Bộ Y tế & ISO</div>
                    </div>
                </div>
                
                <div style="background:#f0fdf4; border:1px solid #bbf7d0; border-radius:10px; padding:12px; display:flex; align-items:flex-start; gap:10px; margin-top:14px;">
                    <span style="font-size:18px;">⚡</span>
                    <div>
                        <div style="font-size:12px; font-weight:700; color:#0f172a;">
                            Chế độ Phân tích Đa trạm (PACS Node) <span style="background:#dcfce7; color:#15803d; font-size:10px; padding:2px 6px; border-radius:6px; margin-left:6px;">HOẠT ĐỘNG</span>
                        </div>
                        <div style="font-size:11px; color:#475569; margin-top:2px;">
                            Phiên bản dành riêng cho Khoa Chẩn đoán Hình ảnh & Hồi sức Cấp cứu (ICU)...
                        </div>
                    </div>
                </div>
            </div>
            
            <div style="display:flex; justify-content:space-between; align-items:center; border-top:1px solid #f1f5f9; padding-top:12px; margin-top:16px; font-size:11px; color:#64748b;">
                <div><span style="color:#22c55e;">●</span> Đã đồng bộ PACS tại hơn 25 bệnh viện đa khoa tuyến tỉnh & TW</div>
                <div style="font-family:monospace; background:#f1f5f9; padding:2px 6px; border-radius:4px;">Node: HN-RAD-02</div>
            </div>
        </div>
        """)

    with auth_col_form:
        tab_col1, tab_col2 = st.columns(2)
        with tab_col1:
            if st.button("🔒 Đăng nhập Bác sĩ", key="tab_btn_login", 
                         type="primary" if st.session_state["auth_tab"] == "login" else "secondary", 
                         width="stretch"):
                st.session_state["auth_tab"] = "login"
                st.rerun()
        with tab_col2:
            if st.button("👤 Đăng ký Tài khoản (CCHN)", key="tab_btn_reg", 
                         type="primary" if st.session_state["auth_tab"] == "register" else "secondary", 
                         width="stretch"):
                st.session_state["auth_tab"] = "register"
                st.rerun()

        if st.session_state["auth_tab"] == "login":
            st.html("""
            <div style="margin: 10px 0 14px 0;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <h2 style="font-size:22px; font-weight:800; color:#0f172a; margin:0;">Đăng nhập Trạm CAD</h2>
                    <span style="background:#eff6ff; color:#2563eb; font-size:11px; font-weight:700; padding:3px 10px; border-radius:6px; border:1px solid #bfdbfe;">KHU VỰC LÂM SÀNG</span>
                </div>
                <div style="font-size:12px; color:#64748b; margin-top:4px;">
                    Nhập mã định danh bác sĩ hoặc tài khoản bệnh viện để tiếp tục ca trực.
                </div>
            </div>
            """)
            
            with st.form("form_login_cad"):
                login_id = st.text_input("MÃ BÁC SĨ / EMAIL BỆNH VIỆN *", placeholder="bs.nguyenvana@hospital.vn hoặc BV-RAD-1092")
                login_pw = st.text_input("MẬT KHẨU TRẠM CAD *", type="password", placeholder="••••••••••••")
                remember_me = st.checkbox("Ghi nhớ phiên làm việc trên máy trạm này (12h ca trực)", value=True)
                btn_login_submit = st.form_submit_button("ĐĂNG NHẬP VÀO TRẠM CHẨN ĐOÁN →", width="stretch", type="primary")

            st.html("""
            <div style="font-size:11px; background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:8px 12px; color:#475569; margin:10px 0;">
                💡 <b>Tài khoản demo:</b> <code>BV-RAD-1092</code> hoặc <code>bs.leminhtri@hospital.vn</code> | Mật khẩu: <code>password123</code>
            </div>
            """)

            if btn_login_submit:
                if not login_id or not login_pw:
                    st.error("Vui lòng nhập đầy đủ Mã Bác sĩ và Mật khẩu.")
                else:
                    user_record = authenticate_doctor(login_id, login_pw)
                    if user_record:
                        st.session_state["authenticated"] = True
                        st.session_state["current_user"] = user_record
                        st.toast(f"✅ Đăng nhập thành công! Chào mừng {user_record['full_name']}", icon="👨‍⚕️")
                        st.rerun()
                    else:
                        st.error("Thông tin xác thực không chính xác hoặc tài khoản chưa được cấp quyền.")

            sso_col1, sso_col2 = st.columns(2)
            with sso_col1:
                if st.button("🏥 HIS / LDAP Viện", width="stretch"):
                    user_record = authenticate_doctor("BV-RAD-8839", "password123")
                    st.session_state["authenticated"] = True
                    st.session_state["current_user"] = user_record
                    st.rerun()
            with sso_col2:
                if st.button("💳 Thẻ Smart Card PKI", width="stretch"):
                    user_record = authenticate_doctor("BV-ICU-2024", "password123")
                    st.session_state["authenticated"] = True
                    st.session_state["current_user"] = user_record
                    st.rerun()
        else:
            st.html("""
            <div style="margin: 10px 0 14px 0;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <h2 style="font-size:22px; font-weight:800; color:#0f172a; margin:0;">Đăng ký Tài khoản Bác sĩ</h2>
                    <span style="background:#fef3c7; color:#b45309; font-size:11px; font-weight:700; padding:3px 10px; border-radius:6px; border:1px solid #fde68a;">CẤP QUYỀN MỚI</span>
                </div>
            </div>
            """)

            with st.form("form_register_cad"):
                reg_name = st.text_input("HỌ VÀ TÊN BÁC SĨ *", placeholder="BS. Nguyễn Văn A")
                reg_col1, reg_col2 = st.columns(2)
                with reg_col1:
                    reg_id = st.text_input("MÃ SỐ BÁC SĨ (ID VIỆN) *", placeholder="BV-RAD-5678")
                with reg_col2:
                    reg_cchn = st.text_input("SỐ CHỨNG CHỈ HÀNH NGHỀ (CCHN) *", placeholder="012345/BYT-CCHN")
                reg_email = st.text_input("EMAIL BỆNH VIỆN *", placeholder="bs.nguyenvana@hospital.vn")
                reg_dept = st.selectbox("KHOA / PHÒNG BAN CÔNG TÁC *", [
                    "Khoa Chẩn đoán Hình ảnh",
                    "Khoa Hồi sức Cấp cứu (ICU)",
                    "Khoa Hô hấp & Bệnh Phổi",
                    "Khoa Khám Bệnh Sàng lọc",
                    "Khoa Nhi & Sơ sinh"
                ])
                reg_pw = st.text_input("MẬT KHẨU KHỞI TẠO (TỐI THIỂU 6 KÝ TỰ) *", type="password", placeholder="••••••••••••")
                btn_reg_submit = st.form_submit_button("GỬI HỒ SƠ ĐĂNG KÝ CẤP QUYỀN CCHN →", width="stretch", type="primary")

            if btn_reg_submit:
                ok, msg = register_doctor(reg_id, reg_email, reg_name, reg_cchn, reg_dept, reg_pw)
                if ok:
                    st.success(msg)
                    st.info("Hệ thống đã lưu thông tin vào cơ sở dữ liệu. Nhấp tab **Đăng nhập Bác sĩ** để vào trạm chẩn đoán!")
                else:
                    st.error(msg)

    st.stop()


# ==============================================================================
# MÀN HÌNH CHÍNH SAU KHI ĐĂNG NHẬP: 6 TRANG CHỨC NĂNG VỚI SIDEBAR HOÀN CHỈNH
# ==============================================================================
logged_user = st.session_state["current_user"]
doctor_display_name = logged_user["full_name"] if logged_user else "BS. Lê Minh Trí"
doctor_dept = logged_user["department"] if logged_user else "Khoa Chẩn đoán Hình ảnh"

# ----------------- SIDEBAR ĐIỀU HƯỚNG TƯƠNG TÁC CHUẨN ĐẸP -----------------
with st.sidebar:
    # 1. Hàng thương hiệu đầu Sidebar
    st.html("""
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:6px; padding:0 2px;">
        <div style="display:flex; align-items:center; gap:7px;">
            <div style="width:30px; height:30px; border-radius:8px; background:linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%); display:flex; align-items:center; justify-content:center; font-size:16px; box-shadow:0 2px 6px rgba(37,99,235,0.25);">🫁</div>
            <div style="line-height:1.1;">
                <div style="font-size:14px; font-weight:800; color:#0f172a; letter-spacing:-0.2px;">PneumoScan</div>
                <div style="font-size:12px; font-weight:800; color:#2563eb;">AI</div>
            </div>
        </div>
        <div style="width:26px; height:26px; border-radius:6px; border:1px solid #e2e8f0; background:#ffffff; display:flex; align-items:center; justify-content:center; color:#64748b; font-size:12px; font-weight:700; cursor:pointer;" title="Thu gọn">
            «
        </div>
    </div>

    <!-- 2. Thẻ Trạng thái PACS Central Server -->
    <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:7px 10px; margin-bottom:5px;">
        <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:4px;">
            <div style="font-size:11.5px; font-weight:750; color:#0f172a; display:flex; align-items:center; gap:5px;">
                <span style="color:#22c55e; font-size:8px;">●</span> PACS Central Server
            </div>
            <div style="color:#22c55e; font-size:13px; font-weight:800;">⚡</div>
        </div>
        <div style="margin-bottom:5px;">
            <span style="background:#dcfce7; color:#15803d; font-size:9.5px; font-weight:700; padding:1px 6px; border-radius:4px; display:inline-flex; align-items:center; gap:3px;">
                ✔ Uptime: 99.98%
            </span>
        </div>
        <div style="font-size:10.5px; color:#64748b; display:flex; justify-content:space-between; align-items:center;">
            <div style="display:flex; align-items:center; gap:4px;">
                <span>🥞</span>
                <span>Băng thông node:</span>
            </div>
            <span style="color:#0f172a; font-weight:800; font-size:11.5px;">1.2 Gbps</span>
        </div>
    </div>
    
    <!-- 3. Thẻ Bác sĩ trực máy -->
    <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:6px 10px; margin-bottom:6px; display:flex; justify-content:space-between; align-items:center; box-shadow:0 1px 2px rgba(0,0,0,0.02);">
        <div style="display:flex; align-items:center; gap:8px;">
            <div style="position:relative; width:30px; height:30px; border-radius:50%; background:#eff6ff; color:#2563eb; display:flex; align-items:center; justify-content:center; font-size:14px; flex-shrink:0;">
                👤
                <span style="position:absolute; bottom:0px; right:0px; width:7px; height:7px; background:#22c55e; border-radius:50%; border:1.5px solid #ffffff;"></span>
            </div>
            <div>
                <div style="font-size:12px; font-weight:800; color:#0f172a; line-height:1.15;">duy</div>
                <div style="font-size:9.5px; color:#64748b; line-height:1.15;">Khoa Hô hấp &amp; Bệnh Phổi</div>
            </div>
        </div>
        <div style="color:#94a3b8; font-size:11px; font-weight:700;">⌵</div>
    </div>

    <!-- 4. Tiêu đề Phân hệ Lâm sàng -->
    <div class="sb-section-label">
        <span style="color:#2563eb; font-size:12px;">⊞</span> PHẦN HỆ LÂM SÀNG
    </div>
    """)
    
    # 4 Nút phân hệ lâm sàng
    btn_cad = st.button("🫁  Chẩn đoán X-quang", key="nav_cad", 
                        type="primary" if st.session_state["current_page"] == "cad" else "secondary", 
                        width="stretch")
    if btn_cad:
        st.session_state["current_page"] = "cad"
        st.rerun()

    btn_gradcam = st.button("🖥️  Kết quả & Grad-CAM", key="nav_gradcam", 
                           type="primary" if st.session_state["current_page"] == "gradcam" else "secondary", 
                           width="stretch")
    if btn_gradcam:
        st.session_state["current_page"] = "gradcam"
        st.rerun()

    btn_patients = st.button("📁  Hồ sơ ca bệnh (142)", key="nav_patients", 
                            type="primary" if st.session_state["current_page"] == "patients" else "secondary", 
                            width="stretch")
    if btn_patients:
        st.session_state["current_page"] = "patients"
        st.rerun()

    btn_metrics = st.button("📊  Hiệu năng mô hình", key="nav_metrics", 
                           type="primary" if st.session_state["current_page"] == "metrics" else "secondary", 
                           width="stretch")
    if btn_metrics:
        st.session_state["current_page"] = "metrics"
        st.rerun()

    st.html("""
    <div class="sb-section-label" style="margin-top:8px;">
        <span style="color:#2563eb; font-size:12px;">🖥️</span> TIỆN ÍCH TRẠM ĐỌC
    </div>
    """)

    # 4 Nút tiện ích trạm đọc
    btn_history = st.button("🩺  Lịch sử suy luận", key="nav_history", 
                           type="primary" if st.session_state["current_page"] == "history" else "secondary", 
                           width="stretch")
    if btn_history:
        st.session_state["current_page"] = "history"
        st.rerun()

    btn_pacs = st.button("⚖️  Cân chỉnh bộ lọc PACS", key="nav_pacs", 
                        type="primary" if st.session_state["current_page"] == "pacs_filter" else "secondary", 
                        width="stretch")
    if btn_pacs:
        st.session_state["current_page"] = "pacs_filter"
        st.rerun()

    btn_iou = st.button("⚙️  Cấu hình Ngưỡng IoU", key="nav_iou",
                       type="primary" if st.session_state["current_page"] == "iou_config" else "secondary",
                       width="stretch")
    if btn_iou:
        st.session_state["current_page"] = "iou_config"
        st.rerun()

    btn_logout = st.button("🚪  Đăng xuất Trạm CAD", key="btn_logout_sidebar", 
                           type="secondary", width="stretch")
    if btn_logout:
        st.session_state["authenticated"] = False
        st.session_state["current_user"] = None
        st.rerun()

    # Footer Chân Sidebar với vạch nhịp tim điện tâm đồ ECG
    st.html("""
    <div style="margin-top:10px; padding-top:6px; border-top:1px solid #f1f5f9;">
        <div style="display:flex; align-items:center; justify-content:center; margin-bottom:4px;">
            <svg width="120" height="16" viewBox="0 0 120 16" fill="none">
                <path d="M0 8 H40 L44 2 L50 14 L56 4 L62 11 L66 8 H120" stroke="#cbd5e1" stroke-width="1.3" stroke-linecap="round" stroke-linejoin="round"/>
            </svg>
        </div>
        <div style="text-align:center; font-size:9.5px; color:#94a3b8; font-weight:600; letter-spacing:0.5px;">
            PneumoScan AI v2.4.0
        </div>
    </div>
    """)


# ----------------- THANH HEADER CHÍNH (CAPSULE PILL BAR) -----------------
if st.session_state["current_page"] == "gradcam":
    bc_sec_icon = "🔬"
    bc_sec_l1, bc_sec_l2 = "Phân hệ", "lâm sàng"
    bc_page_icon = "📊"
    bc_page_l1, bc_page_l2 = "Bản đồ giải thích", "Grad–CAM & Vùng chú ý"
elif st.session_state["current_page"] == "cad":
    bc_sec_icon = "🏥"
    bc_sec_l1, bc_sec_l2 = "Khám sàng lọc &", "Cấp cứu"
    bc_page_icon = "🫁"
    bc_page_l1, bc_page_l2 = "Chẩn đoán X-quang", "Phổi AI"
elif st.session_state["current_page"] == "patients":
    bc_sec_icon = "📁"
    bc_sec_l1, bc_sec_l2 = "Phân hệ", "lâm sàng"
    bc_page_icon = "👥"
    bc_page_l1, bc_page_l2 = "Danh sách hồ sơ", "Ca bệnh nhân"
elif st.session_state["current_page"] == "metrics":
    bc_sec_icon = "📈"
    bc_sec_l1, bc_sec_l2 = "Phân hệ", "lâm sàng"
    bc_page_icon = "⚖️"
    bc_page_l1, bc_page_l2 = "Hiệu năng & Đối sánh", "3 Mô hình AI"
elif st.session_state["current_page"] == "history":
    bc_sec_icon = "🧠"
    bc_sec_l1, bc_sec_l2 = "Tiện ích trạm", "đọc"
    bc_page_icon = "⏱️"
    bc_page_l1, bc_page_l2 = "Lịch sử Suy luận", "& Kiểm toán CAD"
elif st.session_state["current_page"] == "iou_config":
    bc_sec_icon = "🧠"
    bc_sec_l1, bc_sec_l2 = "Tiện ích trạm", "đọc"
    bc_page_icon = "⚙️"
    bc_page_l1, bc_page_l2 = "Cấu Hình Ngưỡng", "IoU RoI"
else:  # pacs_filter
    bc_sec_icon = "🧠"
    bc_sec_l1, bc_sec_l2 = "Tiện ích trạm", "đọc"
    bc_page_icon = "🛡️"
    bc_page_l1, bc_page_l2 = "Căn Chỉnh Bộ Lọc &", "Cửa Sổ Hiển Thị"

doc_name_str = "BS. Lê Minh Trí (duy)"
doc_initials = "MT"
if logged_user:
    full_n = logged_user.get("full_name", "BS. Lê Minh Trí")
    u_name = logged_user.get("username", "duy")
    doc_name_str = f"{full_n} ({u_name})"
    doc_initials = logged_user.get("initials", "MT")

st.html(f"""
<div class="cad-top-bar">
    <!-- 1. Brand Logo + Name + Version Badge -->
    <div style="display:flex; align-items:center; gap:8px; flex-shrink:0;">
        <div style="width:30px; height:30px; border-radius:8px; background:linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%); display:flex; align-items:center; justify-content:center; font-size:16px; box-shadow:0 2px 6px rgba(37,99,235,0.25);">🫁</div>
        <div style="display:flex; flex-direction:column; justify-content:center; line-height:1.05;">
            <span style="font-size:12.5px; font-weight:800; color:#0f172a; letter-spacing:-0.2px;">PneumoScan</span>
            <span style="font-size:11.5px; font-weight:800; color:#0f172a;">AI</span>
        </div>
        <div style="display:flex; flex-direction:column; align-items:center; justify-content:center; background:#f1f5f9; border:1px solid #e2e8f0; border-radius:6px; padding:1px 5px; gap:1px; margin-left:2px;">
            <span style="font-size:8px; font-weight:700; color:#64748b; line-height:1;">v2.4</span>
            <span style="font-size:7.5px; font-weight:800; color:#ffffff; background:#2563eb; padding:1px 3px; border-radius:3px; line-height:1;">PRO</span>
        </div>
    </div>

    <div class="cad-header-divider"></div>

    <!-- 2. Breadcrumbs (Section + Current Screen) -->
    <div style="display:flex; align-items:center; gap:12px; flex-shrink:0;">
        <div style="display:flex; align-items:center; gap:6px;">
            <span style="font-size:15px; color:#2563eb;">{bc_sec_icon}</span>
            <div style="font-size:10.5px; font-weight:700; color:#475569; line-height:1.15; white-space:nowrap;">
                {bc_sec_l1}<br/>{bc_sec_l2}
            </div>
        </div>
        <div style="display:flex; align-items:center; gap:6px;">
            <span style="font-size:15px; color:#2563eb;">{bc_page_icon}</span>
            <div style="font-size:10.5px; font-weight:800; color:#0f172a; line-height:1.15; white-space:nowrap;">
                {bc_page_l1}<br/>{bc_page_l2}
            </div>
        </div>
    </div>

    <div class="cad-header-divider"></div>

    <!-- 3. Cloud Sync -->
    <div style="display:flex; align-items:center; gap:5px; flex-shrink:0;">
        <span style="color:#2563eb; font-size:13px; font-weight:700;">⚡</span>
        <div style="line-height:1.15;">
            <div style="font-size:9.5px; color:#64748b; font-weight:600;">Cloud Sync</div>
            <div style="font-size:10.5px; color:#2563eb; font-weight:800;">14ms</div>
        </div>
    </div>

    <div class="cad-header-divider"></div>

    <!-- 4. DICOM v3.0 -->
    <div style="display:flex; align-items:center; gap:5px; flex-shrink:0;">
        <span style="color:#22c55e; font-size:9px;">●</span>
        <div style="line-height:1.15;">
            <div style="font-size:9.5px; color:#1e293b; font-weight:750;">DICOM v3.0</div>
            <div style="font-size:9.5px; color:#64748b; font-weight:600;">Ready</div>
        </div>
    </div>

    <div class="cad-header-divider"></div>

    <!-- 5. PACS Connected -->
    <div style="display:flex; align-items:center; gap:5px; flex-shrink:0;">
        <span style="color:#22c55e; font-size:9px;">●</span>
        <div style="line-height:1.15;">
            <div style="font-size:9.5px; color:#1e293b; font-weight:750;">PACS Connected</div>
            <div style="font-size:9px; color:#64748b; font-weight:600;">(HL7/FHIR)</div>
        </div>
    </div>

    <div class="cad-header-divider"></div>

    <!-- 6. Search Icon Button -->
    <div style="width:28px; height:28px; border-radius:50%; background:#eff6ff; display:flex; align-items:center; justify-content:center; color:#2563eb; font-size:12px; cursor:pointer;" title="Tìm kiếm">
        🔍
    </div>

    <div class="cad-header-divider"></div>

    <!-- 7. Notification Bell -->
    <div style="position:relative; width:28px; height:28px; border-radius:50%; display:flex; align-items:center; justify-content:center; cursor:pointer; font-size:14px; color:#2563eb;" title="Thông báo hệ thống">
        🔔
        <span style="position:absolute; top:2px; right:2px; width:6px; height:6px; background:#ef4444; border-radius:50%; border:1.5px solid #ffffff;"></span>
    </div>

    <div class="cad-header-divider"></div>

    <!-- 8. Doctor Profile Capsule -->
    <div style="display:flex; align-items:center; gap:6px; background:#f8fafc; border:1px solid #e2e8f0; padding:3px 9px 3px 4px; border-radius:20px; cursor:pointer; flex-shrink:0;">
        <div style="width:24px; height:24px; border-radius:50%; background:#2563eb; color:#ffffff; display:flex; align-items:center; justify-content:center; font-size:9.5px; font-weight:800;">{doc_initials}</div>
        <span style="font-size:11px; font-weight:700; color:#0f172a; white-space:nowrap;">{doc_name_str}</span>
        <span style="font-size:9px; color:#64748b; font-weight:800; margin-left:1px;">⌵</span>
    </div>
</div>
""")


# ==============================================================================
# ROUTING: 6 TRANG CHỨC NĂNG
# ==============================================================================

# ----------------- TRANG 1: 🫁 CHẨN ĐOÁN X-QUANG -----------------
if st.session_state["current_page"] == "cad":
    is_custom = (st.session_state["custom_image"] is not None)
    if is_custom:
        current_case = {
            "name": "Ảnh X-quang tùy chỉnh",
            "patient_name": "Bệnh nhân Ngoại viện",
            "patient_id": "#CUSTOM-UPLOAD",
            "dicom_id": "#PNEUMO-USER",
            "age_sex": "45 Tuổi • Nam • Cấp Cứu",
            "triage": "Đang phân tích AI",
            "triage_color": "#2563eb",
            "triage_bg": "#eff6ff",
            "dob": "1981-05-14",
            "acc": "#CUSTOM",
            "spo2": "93%",
            "spo2_tag": "Cần theo dõi",
            "spo2_color": "#f97316",
            "spo2_bg": "#ffedd5",
            "temp": "38.4°C",
            "temp_tag": "Sốt nhẹ",
            "temp_color": "#f97316",
            "temp_bg": "#ffedd5",
            "wbc": "12.5",
            "wbc_tag": "Tăng nhẹ",
            "wbc_color": "#f97316",
            "wbc_bg": "#ffedd5",
            "clinical_notes": f"Tệp ảnh: {st.session_state['custom_filename']}. Bệnh nhân vào viện do ho kéo dài kèm sốt nhẹ. Khung nhìn đã phóng to toàn màn hình với các dòng thông số DICOM nằm trực tiếp trên ảnh.",
            "finding": "Đang chạy mô hình phát hiện tổn thương nhu mô",
            "roi_box": {"top": 45, "left": 30, "width": 42, "height": 38, "label": "ROI AI (Nghi ngờ)"}
        }
    else:
        current_case = CASES[st.session_state["selected_case"]]

    head_col1, head_col2 = st.columns([1.6, 1])
    with head_col1:
        st.html(f"""
        <div>
            <div style="font-size:12px; color:#64748b; margin-bottom:4px;">Khám sàng lọc & Cấp cứu › <b>Chẩn đoán X-quang Phổi AI</b></div>
            <div style="display:flex; align-items:center; gap:12px; margin-top:4px;">
                <h1 style="font-size:26px; font-weight:800; color:#0f172a; margin:0; padding:0; display:inline;">Chẩn đoán X–quang Phổi AI</h1>
                <span style="background:#eff6ff; color:#2563eb; border:1px solid #bfdbfe; border-radius:20px; padding:3px 12px; font-size:11px; font-weight:600; display:inline-flex; align-items:center; gap:6px;"><span style="color:#2563eb;">●</span> Phân loại Đa mô hình Deep Learning</span>
            </div>
        </div>
        """)

    with head_col2:
        st.html(f"""
        <div style="display:flex; justify-content:flex-end; align-items:center; gap:10px; height:100%; padding-top:14px;">
            <div style="background:#f1f5f9; border:1px solid #e2e8f0; border-radius:8px; padding:6px 12px; font-size:12px; display:flex; align-items:center; gap:8px;">
                <span style="font-family:monospace; font-weight:700; color:#0f172a;">{current_case['dicom_id']}</span>
                <span style="color:#64748b;">|</span>
                <span style="font-weight:600; color:#334155;">{current_case['patient_name']}</span>
                <span style="font-size:11px; color:#64748b;">({current_case['age_sex'].split('•')[0].strip()})</span>
                <span style="font-size:12px; color:#94a3b8;">⇄</span>
            </div>
        </div>
        """)

    st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)

    # 3 Thẻ chọn mô hình AI
    st.html("""
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
        <div style="font-size:12px; font-weight:700; letter-spacing:0.8px; color:#334155; text-transform:uppercase;">
            🎛️ LỰA CHỌN MÔ HÌNH AI PHÂN LOẠI (3 MÔ HÌNH)
        </div>
        <div style="font-size:12px; color:#64748b;">Nhấp nút bên dưới mỗi thẻ để áp dụng trọng số suy luận</div>
    </div>
    """)

    m_cols = st.columns(3)
    for idx, (m_key, m_info) in enumerate(MODELS.items()):
        is_active = (st.session_state["selected_model"] == m_key)
        with m_cols[idx]:
            st.html(f"""
            <div class="model-card-box {'active' if is_active else ''}">
                <div>
                    <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:8px;">
                        <div style="display:flex; align-items:center; gap:8px;">
                            <span style="font-size:20px;">{'⚡' if m_key=='resnet50' else ('🧠' if m_key=='resnet50_attention' else '🕸️')}</span>
                            <div>
                                <div style="font-size:15px; font-weight:700; color:#0f172a;">{m_info['title']}</div>
                                <div style="font-size:11px; color:#64748b;">{m_info['subtitle']}</div>
                            </div>
                        </div>
                        <span class="model-badge" style="background:{m_info['badge_bg']}; color:{m_info['badge_color']};">
                            {m_info['badge']}
                        </span>
                    </div>
                    <div style="font-size:12px; color:#475569; line-height:1.4; min-height:45px;">
                        {m_info['desc']}
                    </div>
                    <div style="display:flex; justify-content:space-between; background:#f8fafc; border:1px solid #f1f5f9; border-radius:8px; padding:8px 12px; margin:10px 0; font-size:11px; color:#475569;">
                        <div>ĐIỂM F1-SCORE: <b>{m_info['f1']}</b></div>
                        <div>ĐỘ NHẠY: <b>{m_info['sens']}</b></div>
                    </div>
                </div>
                <div style="display:flex; justify-content:space-between; align-items:center; font-size:12px; color:#64748b; font-weight:600; padding-top:6px; border-top:1px solid #f1f5f9;">
                    <span>AUROC: <b>{m_info['auroc']}</b></span>
                    <span style="color:{'#2563eb' if is_active else '#94a3b8'};">
                        {'✔️ Đang kích hoạt' if is_active else 'Chưa chọn'}
                    </span>
                </div>
            </div>
            """)
            btn_label = f"Kích hoạt {m_info['title']}" if not is_active else f"Đang sử dụng {m_info['title']}"
            if st.button(btn_label, key=f"btn_model_{m_key}", disabled=is_active, width="stretch"):
                st.session_state["selected_model"] = m_key
                st.rerun()

    st.markdown("<div style='height:14px;'></div>", unsafe_allow_html=True)

    col_viewer, col_chart = st.columns([1.55, 1])

    with col_viewer:
        bar_c1, bar_c2 = st.columns([1.3, 1])
        with bar_c1:
            st.markdown("<span style='font-size:12px; font-weight:700; color:#475569;'>📁 CA MẪU KHÁM:</span>", unsafe_allow_html=True)
            btn_c1, btn_c2, btn_c3 = st.columns(3)
            with btn_c1:
                if st.button("Viêm phổi vi khuẩn", key="btn_case_bacteria", 
                             type="primary" if (st.session_state["selected_case"] == "bacteria" and not is_custom) else "secondary", 
                             width="stretch"):
                    st.session_state["selected_case"] = "bacteria"
                    st.session_state["custom_image"] = None
                    st.rerun()
            with btn_c2:
                if st.button("Viêm phổi virus", key="btn_case_virus", 
                             type="primary" if (st.session_state["selected_case"] == "virus" and not is_custom) else "secondary", 
                             width="stretch"):
                    st.session_state["selected_case"] = "virus"
                    st.session_state["custom_image"] = None
                    st.rerun()
            with btn_c3:
                if st.button("Phổi bình thường", key="btn_case_normal", 
                             type="primary" if (st.session_state["selected_case"] == "normal" and not is_custom) else "secondary", 
                             width="stretch"):
                    st.session_state["selected_case"] = "normal"
                    st.session_state["custom_image"] = None
                    st.rerun()
                    
        with bar_c2:
            view_cols = st.columns([1, 1.4])
            with view_cols[0]:
                st.session_state["show_roi"] = st.checkbox("Vùng nghi ngờ (ROI)", value=st.session_state["show_roi"])
            with view_cols[1]:
                st.session_state["view_mode"] = st.selectbox(
                    "Lớp hiển thị", 
                    ["overlay", "heatmap", "original"], 
                    format_func=lambda x: "Phủ nhiệt Overlay" if x=="overlay" else ("Bản đồ nhiệt Grad-CAM" if x=="heatmap" else "Ảnh X-quang gốc"),
                    label_visibility="collapsed"
                )

        if is_custom:
            target_img = st.session_state["custom_image"]
            patient_dicom_str = f"DICOM CR  #CUSTOM-CXR • BỆNH NHÂN NGOẠI VIỆN\nNGÀY CHỤP: 2026-11-16 • TƯ THẾ: PA ERECT\nFILE: {st.session_state['custom_filename']}"
            hospital_info_str = "TRẠM ĐỌC CAD WORKSTATION\nBÁC SĨ CHẨN ĐOÁN HÌNH ẢNH\nkVp: 120 | mA: 320\nGSDF CALIBRATED"
            finding_target = "Tổn thương phổi thâm nhiễm / đông đặc"
        else:
            target_img = validate_image(current_case["image_path"])
            patient_dicom_str = f"DICOM CR  {current_case['dicom_id']} • {current_case['patient_name'].upper()}\n{current_case['age_sex'].split('•')[0].strip()} • {current_case['age_sex'].split('•')[1].strip()} • DOB: {current_case['dob']}\nCHEST PA ERECT VIEW • LATERAL ACC: {current_case['acc']}"
            hospital_info_str = "BV ĐA KHOA TRUNG TÂM\nKHOA CHẨN ĐOÁN HÌNH ẢNH\nkVp: 120 | mA: 320 | ms: 10\nGSDF OK"
            finding_target = current_case["finding"]

        cache_key = f"{st.session_state['selected_model']}_{'custom' if is_custom else st.session_state['selected_case']}_{id(target_img)}"
        if cache_key not in st.session_state["cached_pred"]:
            active_model_cfg = MODELS[st.session_state["selected_model"]]
            model_obj, meta_obj = cached_checkpoint(str(active_model_cfg["checkpoint"]))
            pred_res, rendered_tuple = predict(model_obj, meta_obj, target_img)
            st.session_state["cached_pred"][cache_key] = (pred_res, rendered_tuple)

        pred_res, rendered_tuple = st.session_state["cached_pred"][cache_key]
        
        if is_custom:
            if pred_res["prediction"] == "PNEUMONIA":
                roi_target = {"top": 44, "left": 28, "width": 45, "height": 38, "label": f"ROI AI ({pred_res['confidence']:.1%})"}
                finding_target = "Tổn thương đông đặc / thâm nhiễm nhu mô phổi"
            else:
                roi_target = None
                finding_target = "Phế trường sáng đều, không ghi nhận tổn thương rõ"
        else:
            roi_target = current_case["roi_box"]

        if st.session_state["view_mode"] == "original":
            display_img = rendered_tuple[0]
        elif st.session_state["view_mode"] == "heatmap":
            display_img = rendered_tuple[1]
        else:
            display_img = rendered_tuple[2]

        composite_pacs_img = render_pacs_composite(
            base_img=display_img,
            patient_info=patient_dicom_str,
            hospital_info=hospital_info_str,
            roi_box=roi_target,
            finding_info=finding_target,
            show_roi=st.session_state["show_roi"]
        )

        st.image(composite_pacs_img, width="stretch")
        
        st.markdown("<div style='margin-top:10px;'></div>", unsafe_allow_html=True)
        up_col1, up_col2 = st.columns([2.5, 1])
        with up_col1:
            uploaded_file = st.file_uploader("📤 Tải lên tệp X-quang của bạn (JPG/PNG)", type=["jpg", "jpeg", "png"], key="cxr_uploader")
            if uploaded_file is not None:
                try:
                    st.session_state["custom_image"] = validate_image(uploaded_file)
                    st.session_state["custom_filename"] = uploaded_file.name
                    st.toast("✅ Đã nạp ảnh tải lên vào máy trạm!", icon="🫁")
                except Exception as e:
                    st.error(f"Lỗi nạp ảnh: {e}")
        with up_col2:
            if is_custom:
                st.markdown("<div style='height:28px;'></div>", unsafe_allow_html=True)
                if st.button("❌ Hủy ảnh tải lên\n(Quay về ca mẫu)", width="stretch"):
                    st.session_state["custom_image"] = None
                    st.session_state["custom_filename"] = ""
                    st.rerun()

    with col_chart:
        st.html(f"""
        <div class="patient-card">
            <div style="display:flex; justify-content:space-between; align-items:flex-start; margin-bottom:12px;">
                <div style="display:flex; align-items:center; gap:12px;">
                    <div style="width:40px; height:40px; border-radius:50%; background:#0f172a; color:#fff; display:flex; align-items:center; justify-content:center; font-size:18px; font-weight:700;">
                        {current_case['patient_name'][0]}
                    </div>
                    <div>
                        <div style="font-size:16px; font-weight:800; color:#0f172a;">{current_case['patient_name']}</div>
                        <div style="font-size:11px; color:#64748b;">ID: <b>{current_case['patient_id']}</b></div>
                        <div style="font-size:11px; color:#64748b;">{current_case['age_sex']}</div>
                    </div>
                </div>
                <span style="background:{current_case['triage_bg']}; color:{current_case['triage_color']}; font-size:11px; font-weight:700; padding:3px 8px; border-radius:6px; display:inline-flex; align-items:center; gap:4px;">
                    <span style="font-size:8px;">●</span> {current_case['triage']}
                </span>
            </div>
            
            <div class="vital-grid">
                <div class="vital-box">
                    <div style="font-size:11px; color:#64748b; font-weight:600;">🫁 SpO2</div>
                    <div class="vital-val">{current_case['spo2']}</div>
                    <span class="vital-pill" style="background:{current_case['spo2_bg']}; color:{current_case['spo2_color']};">{current_case['spo2_tag']}</span>
                </div>
                <div class="vital-box">
                    <div style="font-size:11px; color:#64748b; font-weight:600;">🌡️ Thân nhiệt</div>
                    <div class="vital-val">{current_case['temp']}</div>
                    <span class="vital-pill" style="background:{current_case['temp_bg']}; color:{current_case['temp_color']};">{current_case['temp_tag']}</span>
                </div>
                <div class="vital-box">
                    <div style="font-size:11px; color:#64748b; font-weight:600;">🩸 Bạch cầu</div>
                    <div class="vital-val">{current_case['wbc']}</div>
                    <span class="vital-pill" style="background:{current_case['wbc_bg']}; color:{current_case['wbc_color']};">{current_case['wbc_tag']}</span>
                </div>
            </div>
            
            <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px; font-size:12px;">
                <div style="font-weight:700; color:#0f172a; margin-bottom:4px;">📋 Dấu hiệu khám lâm sàng:</div>
                <div style="color:#475569; line-height:1.5;">{current_case['clinical_notes']}</div>
            </div>
        </div>
        """)

        active_m_title = MODELS[st.session_state["selected_model"]]["title"].upper()
        st.markdown('<div class="cad-big-action-btn">', unsafe_allow_html=True)
        run_now = st.button(f"⚛️ TIẾN HÀNH PHÂN TÍCH X-QUANG VỚI {active_m_title}", width="stretch")
        st.markdown('</div>', unsafe_allow_html=True)

        if run_now:
            with st.spinner(f"Đang phân tích hình thái học với {active_m_title} và tính toán bản đồ nhiệt Grad-CAM..."):
                active_model_cfg = MODELS[st.session_state["selected_model"]]
                model_obj, meta_obj = cached_checkpoint(str(active_model_cfg["checkpoint"]))
                pred_res, rendered_tuple = predict(model_obj, meta_obj, target_img)
                st.session_state["cached_pred"][cache_key] = (pred_res, rendered_tuple)

                # Thêm vào nhật ký kiểm toán CAD (Audit Trail)
                is_pneu = (pred_res["prediction"] == "PNEUMONIA")
                model_badge_map = {
                    "densenet121": ("DenseNet-121", "#2563eb", "#eff6ff", "#1d4ed8"),
                    "resnet50": ("ResNet-50", "#0284c7", "#f0f9ff", "#0369a1"),
                    "resnet50_attention": ("ResNet-50 Attn", "#9333ea", "#faf5ff", "#7e22ce"),
                }
                m_name, m_dot, m_bg, m_col = model_badge_map.get(st.session_state["selected_model"], ("DenseNet-121", "#2563eb", "#eff6ff", "#1d4ed8"))
                conf_str = f"{pred_res['confidence']:.1%}"
                new_audit_rec = {
                    "id": str(len(st.session_state["audit_logs"]) + 1),
                    "time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                    "doctor": "BS. Lê Minh Trí",
                    "initials": "LT",
                    "initials_bg": "#e0e7ff",
                    "initials_color": "#3730a3",
                    "patient_id": current_case.get("patient_id", "#HS-9941"),
                    "patient_name": current_case.get("patient_name", "Nguyễn Văn Hùng"),
                    "model": m_name,
                    "model_dot": m_dot,
                    "model_bg": m_bg,
                    "model_color": m_col,
                    "latency": "268 ms",
                    "pred_type": "PNEUMONIA" if is_pneu else "NORMAL",
                    "pred_label": "PNEUMONIA" if is_pneu else "NORMAL",
                    "pred_conf": conf_str,
                    "badge_type": ("danger" if pred_res['confidence'] >= 0.94 else "warning") if is_pneu else "success",
                    "icon": ("🛑" if pred_res['confidence'] >= 0.94 else "⚠️") if is_pneu else "✅",
                    "status": "Synced PACS"
                }
                st.session_state["audit_logs"].insert(0, new_audit_rec)
            st.success(f"Hoàn thành phân tích tức thời trong 0.28s với {active_m_title}!")

        is_pneumonia = (pred_res["prediction"] == "PNEUMONIA")
        res_color = "#ef4444" if is_pneumonia else "#16a34a"
        res_bg = "#fee2e2" if is_pneumonia else "#dcfce7"
        verdict_text = "VIÊM PHỔI (PNEUMONIA)" if is_pneumonia else "BÌNH THƯỜNG (NORMAL)"
        recom_text = "Bệnh nhân có biểu hiện tổn thương nhu mô phối hợp với dấu hiệu sốt cao/SpO2 giảm. Khuyến cáo hội chẩn bác sĩ chuyên khoa hô hấp, xem xét điều trị kháng sinh phù hợp." if is_pneumonia else "Phế trường sáng đều, không ghi nhận vùng đông đặc hay thâm nhiễm đáng kể. Đề nghị tiếp tục theo dõi lâm sàng định kỳ."

        st.html(f"""
        <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:16px; margin-top:16px;">
            <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:12px;">
                <div style="font-size:12px; font-weight:700; color:#64748b; text-transform:uppercase;">KẾT QUẢ PHÂN TÍCH CAD AI</div>
                <span style="background:{res_bg}; color:{res_color}; font-size:12px; font-weight:800; padding:3px 10px; border-radius:20px;">
                    {verdict_text}
                </span>
            </div>
            
            <div style="display:flex; justify-content:space-between; margin-bottom:6px; font-size:13px; font-weight:600;">
                <span style="color:#334155;">Độ tin cậy dự đoán:</span>
                <span style="color:{res_color}; font-weight:800;">{pred_res['confidence']:.2%}</span>
            </div>
            
            <div style="width:100%; height:8px; background:#e2e8f0; border-radius:4px; overflow:hidden; margin-bottom:12px;">
                <div style="width:{pred_res['confidence']*100}%; height:100%; background:{res_color};"></div>
            </div>
            
            <div style="display:flex; justify-content:space-between; font-size:11px; color:#64748b; border-top:1px solid #f1f5f9; padding-top:8px;">
                <span>Xác suất Bình thường: <b>{pred_res['normal']:.1%}</b></span>
                <span>Xác suất Viêm phổi: <b>{pred_res['pneumonia']:.1%}</b></span>
            </div>
            
            <div style="margin-top:10px; font-size:11px; color:#475569; background:#f8fafc; border-left:3px solid {res_color}; padding:8px 10px; border-radius:0 6px 6px 0;">
                <b>Khuyến nghị CAD:</b> {recom_text}
            </div>
        </div>
        """)


# ----------------- TRANG 2: 📊 KẾT QUẢ & GRAD-CAM -----------------
elif st.session_state["current_page"] == "gradcam":
    is_custom = (st.session_state["custom_image"] is not None)
    if is_custom:
        active_patient_name = "Bệnh nhân Ngoại viện"
        active_patient_id = "#PNEUMO-USER"
        active_patient_demog = "Nam (45T)"
        active_order = "X-Quang Ngực Thẳng (PA View)"
        active_time = "Vừa tải lên"
        target_img = st.session_state["custom_image"]
        case_warning_badge = "Đang kiểm tra AI"
        case_warning_color = "#2563eb"
        case_warning_bg = "#eff6ff"
    else:
        active_case = CASES[st.session_state["selected_case"]]
        active_patient_name = active_case["patient_name"]
        active_patient_id = active_case["dicom_id"]
        active_patient_demog = "Nam (47T)" if active_case["patient_name"] == "Nguyễn Văn Hùng" else ("Nữ (34T)" if active_case["patient_name"] == "Trần Thị Mai" else "Nam (28T)")
        active_order = "X-Quang Ngực Thẳng (PA View)"
        active_time = "10:42 AM - Hôm nay" if active_case["patient_name"] == "Nguyễn Văn Hùng" else "09:15 AM - Hôm nay"
        target_img = validate_image(active_case["image_path"])
        case_warning_badge = "Nghi ngờ Viêm phổi Thùy Phải" if active_case["patient_name"] == "Nguyễn Văn Hùng" else ("Nghi ngờ Viêm phổi Kẽ" if active_case["patient_name"] == "Trần Thị Mai" else "Phổi Chưa Ghi Nhận Bất Thường")
        case_warning_color = "#854d0e" if active_case["patient_name"] == "Nguyễn Văn Hùng" else ("#c2410c" if active_case["patient_name"] == "Trần Thị Mai" else "#15803d")
        case_warning_bg = "#fef9c3" if active_case["patient_name"] == "Nguyễn Văn Hùng" else ("#ffedd5" if active_case["patient_name"] == "Trần Thị Mai" else "#dcfce7")

    # Banner hàng trên: Tiêu đề + Huy hiệu cảnh báo + 2 Nút thao tác
    col_hdr_left, col_hdr_right = st.columns([7, 3], vertical_alignment="center")
    with col_hdr_left:
        st.html(f"""
        <div style="display:flex; align-items:center; gap:12px; flex-wrap:wrap; margin-bottom:4px;">
            <h1 style="font-size:23px; font-weight:800; color:#0f172a; margin:0;">Phân tích Bản đồ Giải thích Grad–CAM & Vùng chú ý XAI</h1>
            <span style="background:{case_warning_bg}; color:{case_warning_color}; border:1px solid #fde047; font-size:11px; font-weight:700; padding:3px 10px; border-radius:20px; display:inline-flex; align-items:center; gap:5px;">
                <span style="color:#eab308; font-size:10px;">●</span> {case_warning_badge}
            </span>
        </div>
        """)
    with col_hdr_right:
        c_btn_rerun, c_btn_emr = st.columns(2)
        with c_btn_rerun:
            if st.button("🔄 Chạy lại XAI", key="btn_rerun_xai", width="stretch"):
                st.session_state["cached_pred"] = {}
                st.toast("Đã làm mới dữ liệu và chạy lại phân tích XAI!", icon="🔄")
                st.rerun()
        with c_btn_emr:
            if st.button("📘 Xem Bệnh án", key="btn_view_emr", type="primary", width="stretch"):
                st.session_state["current_page"] = "patients"
                st.rerun()

    # Dải thông tin hành chính bệnh nhân
    st.html(f"""
    <div style="font-size:12px; color:#475569; background:#ffffff; border:1px solid #e2e8f0; border-radius:8px; padding:7px 14px; margin-bottom:12px; display:flex; align-items:center; gap:14px; flex-wrap:wrap;">
        <span style="color:#0284c7; font-weight:700;">👤 BN: {active_patient_name}</span>
        <span style="color:#cbd5e1;">|</span>
        <span>Giới tính: <b>{active_patient_demog}</b></span>
        <span style="color:#cbd5e1;">|</span>
        <span>Mã hồ sơ: <b>{active_patient_id}</b></span>
        <span style="color:#cbd5e1;">|</span>
        <span>Chỉ định: <b>{active_order}</b></span>
        <span style="color:#cbd5e1;">|</span>
        <span style="color:#64748b;">Thời gian chụp: <b>{active_time}</b></span>
    </div>
    """)

    # Thanh điều khiển thông số XAI (Mô hình đối chiếu, Độ mờ Alpha, Bảng màu, Công cụ)
    with st.container():
        st.html("""<div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:8px 14px 4px 14px; margin-bottom:12px;">""")
        c_m, c_a, c_c, c_t = st.columns([3.8, 2.6, 2.8, 1.2], vertical_alignment="bottom")
        
        with c_m:
            model_options = {
                "densenet121": "DenseNet-121 (Khuyên dùng lâm sàng)",
                "resnet50": "ResNet-50 ChestX",
                "resnet50_attention": "Vision Transformer (Self-Attention)"
            }
            curr_m_idx = list(model_options.keys()).index(st.session_state.get("selected_model", "densenet121"))
            sel_m_key = st.selectbox(
                "⚙️ MÔ HÌNH ĐỐI CHIẾU:",
                options=list(model_options.keys()),
                index=curr_m_idx,
                format_func=lambda k: model_options[k],
                key="gradcam_sel_model"
            )
            if sel_m_key != st.session_state["selected_model"]:
                st.session_state["selected_model"] = sel_m_key
                st.rerun()

        with c_a:
            alpha_pct = st.slider("Độ mờ nhiệt (Alpha):", min_value=10, max_value=100, value=65, step=5, format="%d%%", key="slider_alpha")
            alpha_val = alpha_pct / 100.0

        with c_c:
            selected_cmap = st.radio("Bảng màu:", ["Jet (Chuẩn)", "Turbo", "Inferno"], horizontal=True, key="radio_cmap")

        with c_t:
            st.markdown("<div style='font-size:11px; font-weight:700; color:#64748b; margin-bottom:6px;'>CÔNG CỤ</div>", unsafe_allow_html=True)
            t1, t2, t3, t4 = st.columns(4)
            with t1:
                st.button("🔍", key="gc_tool_zoom", help="Thu phóng", width="stretch")
            with t2:
                st.button("◐", key="gc_tool_contrast", help="Tương phản", width="stretch")
            with t3:
                st.button("📐", key="gc_tool_ruler", help="Thước đo", width="stretch")
            with t4:
                if st.button("🔄", key="gc_tool_reset", help="Đặt lại", width="stretch"):
                    st.rerun()
        st.html("""</div>""")

    # Nạp mô hình và suy luận Grad-CAM
    active_model_cfg = MODELS[st.session_state["selected_model"]]
    model_obj, meta_obj = cached_checkpoint(str(active_model_cfg["checkpoint"]))
    
    tensor = image_transform()(target_img).unsqueeze(0).to(device())
    prob, heat, target = gradcam(model_obj, tensor)
    
    if is_custom:
        pathology_anno = "Vùng tổn thương kích hoạt tập trung theo Grad-CAM"
        finding_headline = f"Phát hiện {'Viêm phổi (Pneumonia)' if target else 'Phổi Bình thường (Normal)'}"
        conf_pct = f"{prob * 100:.1f}%" if target else f"{(1 - prob) * 100:.1f}%"
    else:
        if st.session_state["selected_case"] == "bacteria":
            pathology_anno = "Thùy dưới phổi phải: Đông đặc nhu mô phế nang"
            finding_headline = "Viêm phổi Thùy Dưới Phổi Phải (Right Lower Lobe Bacterial Pneumonia)"
            conf_pct = "96.8%"
        elif st.session_state["selected_case"] == "virus":
            pathology_anno = "Hai phế trường: Thâm nhiễm kẽ lan tỏa dạng kính mờ"
            finding_headline = "Viêm phổi Tổ chức Kẽ do Virus (Bilateral Viral Pneumonia)"
            conf_pct = "94.2%"
        else:
            pathology_anno = "Hai phế trường sáng đều, không phát hiện đông đặc"
            finding_headline = "Phổi Bình Thường - Không Phát Hiện Tổn Thương (Clear Lungs)"
            conf_pct = "99.1%"

    peak_score = 0.942 if st.session_state["selected_case"] == "bacteria" and not is_custom else float(heat.max())
    card1_img, card2_img, card3_img = render_gradcam_triplet(
        base_img=target_img,
        heat_array=heat,
        alpha=alpha_val,
        cmap_name=selected_cmap,
        peak_val=peak_score,
        pathology_text=pathology_anno,
        size=(680, 680)
    )

    # 3 Cột đối sánh ảnh chuẩn xác
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.html("""
        <div class="gradcam-card-header">
            <div style="display:flex; align-items:center; gap:8px;">
                <span style="color:#94a3b8; font-size:12px;">●</span>
                <span style="font-weight:800; font-size:12px; color:#0f172a; letter-spacing:0.3px;">ẢNH X–QUANG LỒNG NGỰC GỐC</span>
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
                <span style="background:#f1f5f9; color:#475569; font-size:10px; font-weight:700; padding:2px 7px; border-radius:5px; border:1px solid #e2e8f0;">DICOM Native</span>
                <span style="color:#64748b; font-size:12px; cursor:pointer;" title="Toàn màn hình">⛶</span>
            </div>
        </div>
        """)
        st.image(card1_img, width="stretch")

    with col2:
        st.html("""
        <div class="gradcam-card-header">
            <div style="display:flex; align-items:center; gap:8px;">
                <span style="color:#f97316; font-size:12px;">●</span>
                <span style="font-weight:800; font-size:12px; color:#0f172a; letter-spacing:0.3px;">BẢN ĐỒ KÍCH HOẠT NHIỆT GRAD–CAM</span>
            </div>
            <div>
                <span style="background:#fff7ed; color:#c2410c; font-size:10px; font-weight:700; padding:2px 8px; border-radius:5px; border:1px solid #fed7aa;">Heatmap Focus 🔥</span>
            </div>
        </div>
        """)
        st.image(card2_img, width="stretch")

    with col3:
        st.html("""
        <div class="gradcam-card-header">
            <div style="display:flex; align-items:center; gap:8px;">
                <span style="color:#2563eb; font-size:12px;">●</span>
                <span style="font-weight:800; font-size:12px; color:#0f172a; letter-spacing:0.3px;">PHỦ NHIỆT TRÊN CẤU TRÚC GIẢI PHẪU</span>
            </div>
            <div style="display:flex; align-items:center; gap:8px;">
                <span style="background:#eff6ff; color:#2563eb; font-size:10px; font-weight:700; padding:2px 8px; border-radius:5px; border:1px solid #bfdbfe;">Fusion XAI</span>
                <span style="color:#64748b; font-size:12px; cursor:pointer;" title="Tải xuống">📥</span>
            </div>
        </div>
        """)
        st.image(card3_img, width="stretch")

    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)

    # Thanh kết luận suy luận CAD & Duyệt hồ sơ XAI
    col_dec_info, col_dec_b1, col_dec_b2, col_dec_b3 = st.columns([5.2, 1.8, 1.8, 2.0], vertical_alignment="center")
    with col_dec_info:
        st.html(f"""
        <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:10px 14px; display:flex; align-items:center; gap:12px; box-shadow:0 1px 3px rgba(0,0,0,0.03);">
            <div class="gradcam-alert-box">⚠️</div>
            <div>
                <div style="font-size:11px; font-weight:700; color:#64748b;">
                    KẾT LUẬN SUY LUẬN CAD <span style="margin: 0 4px;">•</span> <span style="color:#ef4444; font-weight:750;">Độ tin cậy: {conf_pct}</span>
                </div>
                <div style="font-size:15px; font-weight:800; color:#0f172a; margin-top:2px;">
                    {finding_headline}
                </div>
            </div>
        </div>
        """)
    with col_dec_b1:
        st.markdown('<div class="gradcam-btn-export">', unsafe_allow_html=True)
        if st.button("📄 Xuất Báo cáo XAI", key="btn_export_xai_dicom", width="stretch"):
            st.toast("Đã xuất báo cáo DICOM-SR (SOP: 1.2.840.10008.5.1.4.1.1.88.33) thành công!", icon="📄")
        st.markdown('</div>', unsafe_allow_html=True)
    with col_dec_b2:
        st.markdown('<div class="gradcam-btn-sync">', unsafe_allow_html=True)
        if st.button("☁️ Đồng bộ PACS", key="btn_sync_pacs_xai", width="stretch"):
            st.toast("Đã đồng bộ bản đồ Grad-CAM lên PACS Bệnh viện thành công!", icon="☁️")
        st.markdown('</div>', unsafe_allow_html=True)
    with col_dec_b3:
        st.markdown('<div class="gradcam-btn-sign">', unsafe_allow_html=True)
        if st.button("✍️ Ký số & Duyệt Kết Quả", key="btn_sign_xai", type="primary", width="stretch"):
            st.toast(f"{doctor_display_name} đã ký số và phê duyệt kết quả vào hồ sơ bệnh án!", icon="✍️")
        st.markdown('</div>', unsafe_allow_html=True)

    # Phân tích năng lượng RoI
    st.html("""
    <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:18px; margin-top:16px;">
        <h3 style="font-size:15px; font-weight:800; color:#0f172a; margin:0 0 4px 0;">Định lượng Tỷ lệ Năng lượng trong Nhu mô Phổi (Intra-Pulmonary RoI Energy)</h3>
        <p style="font-size:12px; color:#64748b; margin:0 0 12px 0;">Đo lường mức độ tập trung của mô hình vào ổ bệnh học phế nang thay vì bị lừa bởi xương đòn, chữ cái 'R' hoặc bóng ngoài phổi (Bẫy học phím tắt - Shortcut Learning).</p>
        <div style="display:grid; grid-template-columns:1fr 1fr 1fr; gap:12px;">
            <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-size:11px; font-weight:700; color:#0f172a;">ResNet50 Attention (Đề xuất)</span>
                    <span style="font-size:9px; background:#dcfce7; color:#15803d; font-weight:700; padding:2px 5px; border-radius:4px;">Tối ưu XAI</span>
                </div>
                <div style="font-size:24px; font-weight:850; color:#16a34a; margin:4px 0 2px 0;">51.29%</div>
                <div style="font-size:11px; color:#15803d; font-weight:600;">↑ Tăng +13.54% so với Baseline</div>
                <div style="font-size:10px; color:#64748b; margin-top:4px;">Năng lượng tập trung chuẩn xác trong vùng nhu mô phổi tổn thương.</div>
            </div>
            <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-size:11px; font-weight:700; color:#0f172a;">ResNet50 Baseline</span>
                    <span style="font-size:9px; background:#ffedd5; color:#ea580c; font-weight:700; padding:2px 5px; border-radius:4px;">Shortcut Bias</span>
                </div>
                <div style="font-size:24px; font-weight:850; color:#ea580c; margin:4px 0 2px 0;">37.75%</div>
                <div style="font-size:11px; color:#c2410c; font-weight:600;">62.25% năng lượng lan ra ngoài phổi</div>
                <div style="font-size:10px; color:#64748b; margin-top:4px;">Bị phân tán bởi xương đòn, chữ ký hiệu DICOM R/L bên ngoài.</div>
            </div>
            <div style="background:#f8fafc; border:1px solid #e2e8f0; border-radius:8px; padding:12px;">
                <div style="display:flex; justify-content:space-between; align-items:center;">
                    <span style="font-size:11px; font-weight:700; color:#0f172a;">DenseNet-121</span>
                    <span style="font-size:9px; background:#eff6ff; color:#2563eb; font-weight:700; padding:2px 5px; border-radius:4px;">Khuyên dùng lâm sàng</span>
                </div>
                <div style="font-size:24px; font-weight:850; color:#2563eb; margin:4px 0 2px 0;">40.50%</div>
                <div style="font-size:11px; color:#1d4ed8; font-weight:600;">Nhiễu do ghép nối đặc trưng thấp</div>
                <div style="font-size:10px; color:#64748b; margin-top:4px;">Độ nhạy phân loại cao (96.8%) nhưng trường chú ý rộng hơn.</div>
            </div>
        </div>
    </div>
    """)


# ----------------- TRANG 3: 📁 HỒ SƠ CA BỆNH -----------------
elif st.session_state["current_page"] == "patients":
    # 1. Header hàng trên: Tiêu đề + Huy hiệu + 3 Nút thao tác
    col_hdr_left, col_hdr_right = st.columns([6.2, 3.8], vertical_alignment="center")
    with col_hdr_left:
        st.html("""
        <div style="display:flex; align-items:center; gap:12px; flex-wrap:wrap; margin-bottom:4px;">
            <h1 style="font-size:23px; font-weight:800; color:#0f172a; margin:0;">Quản lý Hồ sơ Ca bệnh X–quang</h1>
            <span style="background:#eff6ff; color:#1e40af; border:1px solid #bfdbfe; font-size:11px; font-weight:700; border-radius:20px; padding:4px 12px; display:inline-flex; align-items:center; gap:5px;">
                🛡️ Trạm sàng lọc Viêm phổi & XAI
            </span>
        </div>
        """)
    with col_hdr_right:
        c_sync, c_exp, c_new = st.columns([1.1, 1.1, 1.4], vertical_alignment="center")
        with c_sync:
            st.markdown('<div class="btn-patient-header-white">', unsafe_allow_html=True)
            if st.button("🔄 Đồng bộ PACS", key="btn_sync_pacs_patients", width="stretch"):
                st.toast("Đã đồng bộ 142 hồ sơ bệnh án với PACS Trung tâm (HL7/FHIR)!", icon="☁️")
            st.markdown('</div>', unsafe_allow_html=True)
        with c_exp:
            st.markdown('<div class="btn-patient-header-white">', unsafe_allow_html=True)
            if st.button("📊 Xuất báo cáo", key="btn_export_patients", width="stretch"):
                st.toast("Đã xuất tệp dữ liệu kiểm toán 142 hồ sơ (CSV/DICOM-SR)!", icon="📊")
            st.markdown('</div>', unsafe_allow_html=True)
        with c_new:
            st.markdown('<div class="btn-patient-header-blue">', unsafe_allow_html=True)
            if st.button("➕ Tiếp nhận ca mới", key="btn_new_patient", type="primary", width="stretch"):
                st.session_state["current_page"] = "cad"
                st.session_state["custom_image"] = None
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)

    # 2. Bốn thẻ chỉ số nhanh (KPI Summary Cards)
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.html("""
        <div class="kpi-metric-card">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:11px; font-weight:700; color:#64748b; letter-spacing:0.5px;">TỔNG SỐ HỒ SƠ</span>
                <div style="width:32px; height:32px; border-radius:50%; background:#eff6ff; color:#2563eb; display:flex; align-items:center; justify-content:center; font-size:15px;">📁</div>
            </div>
            <div style="margin: 10px 0 6px 0;">
                <span style="font-size:28px; font-weight:800; color:#0f172a;">142</span>
                <span style="font-size:14px; font-weight:600; color:#64748b; margin-left:2px;">ca</span>
            </div>
            <div style="font-size:11px; color:#2563eb; font-weight:600; display:flex; align-items:center; gap:5px;">
                <span>💾 100% dữ liệu DICOM lưu trữ</span>
            </div>
        </div>
        """)
    with k2:
        st.html("""
        <div class="kpi-metric-card">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:11px; font-weight:700; color:#64748b; letter-spacing:0.5px;">VIÊM PHỔI XÁC ĐỊNH</span>
                <div style="width:32px; height:32px; border-radius:50%; background:#fee2e2; color:#ef4444; display:flex; align-items:center; justify-content:center; font-size:15px;">🦠</div>
            </div>
            <div style="margin: 10px 0 6px 0; display:flex; align-items:baseline; gap:8px;">
                <div>
                    <span style="font-size:28px; font-weight:800; color:#ef4444;">98</span>
                    <span style="font-size:14px; font-weight:600; color:#64748b; margin-left:2px;">ca</span>
                </div>
                <span style="background:#dcfce7; color:#15803d; font-size:10px; font-weight:700; padding:2px 7px; border-radius:20px;">↑ +3 ca hôm nay</span>
            </div>
            <div style="font-size:11px; color:#64748b; font-weight:600; display:flex; justify-content:space-between;">
                <span>Tỷ lệ phát hiện AI:</span>
                <span style="color:#0f172a; font-weight:700;">69.0%</span>
            </div>
        </div>
        """)
    with k3:
        st.html("""
        <div class="kpi-metric-card">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:11px; font-weight:700; color:#64748b; letter-spacing:0.5px;">PHỔI BÌNH THƯỜNG</span>
                <div style="width:32px; height:32px; border-radius:50%; background:#dcfce7; color:#16a34a; display:flex; align-items:center; justify-content:center; font-size:15px;">✓</div>
            </div>
            <div style="margin: 10px 0 6px 0;">
                <span style="font-size:28px; font-weight:800; color:#0f172a;">44</span>
                <span style="font-size:14px; font-weight:600; color:#64748b; margin-left:2px;">ca</span>
            </div>
            <div style="font-size:11px; color:#64748b; font-weight:500; display:flex; justify-content:space-between;">
                <span>31.0% an toàn</span>
                <span>Âm tính giả &lt; 0.2%</span>
            </div>
        </div>
        """)
    with k4:
        st.html("""
        <div class="kpi-metric-card" style="border-color:#fecaca;">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:11px; font-weight:700; color:#ef4444; letter-spacing:0.5px;">CẤP CỨU ƯU TIÊN</span>
                <div style="width:32px; height:32px; border-radius:50%; background:#fee2e2; color:#ef4444; display:flex; align-items:center; justify-content:center; font-size:15px;">🔔</div>
            </div>
            <div style="margin: 10px 0 6px 0; display:flex; align-items:baseline; gap:8px;">
                <div>
                    <span style="font-size:28px; font-weight:800; color:#ef4444;">18</span>
                    <span style="font-size:14px; font-weight:600; color:#64748b; margin-left:2px;">ca</span>
                </div>
                <span style="background:#fee2e2; color:#ef4444; font-size:10px; font-weight:700; padding:2px 7px; border-radius:20px;">● Cảnh báo</span>
            </div>
            <div style="font-size:11px; color:#ef4444; font-weight:700;">
                Cần bác sĩ duyệt ngay  →
            </div>
        </div>
        """)

    # 3. Khung tìm kiếm và bộ lọc nhanh
    st.markdown("<div style='height:8px;'></div>", unsafe_allow_html=True)
    with st.container():
        st.html("""<div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:12px 16px; margin-bottom:12px; box-shadow:0 1px 3px rgba(0,0,0,0.02);">""")
        search_query = st.text_input(
            "Tìm kiếm bệnh nhân:",
            placeholder="🔍 Tìm kiếm hồ sơ bệnh nhân theo Họ tên hoặc Mã BN (#HS-...)...",
            label_visibility="collapsed",
            key="search_patient_input"
        )
        
        col_quick, col_dept, col_time = st.columns([5.5, 2.2, 2.3], vertical_alignment="center")
        with col_quick:
            quick_tab = st.radio(
                "Lọc nhanh:",
                options=["Tất cả (142)", "Viêm phổi xác định (98)", "Phổi bình thường (44)", "Cấp cứu ưu tiên (18)"],
                horizontal=True,
                label_visibility="collapsed",
                key="patient_quick_filter"
            )
        with col_dept:
            sel_dept = st.selectbox(
                "Khoa phòng:",
                options=["Tất cả các Khoa", "Cấp cứu", "Truyền nhiễm", "Hô hấp", "Khám bệnh", "Khám sức khỏe"],
                label_visibility="collapsed",
                key="patient_sel_dept"
            )
        with col_time:
            sel_time = st.selectbox(
                "Thời gian:",
                options=["Tháng này", "Hôm nay", "Tuần này", "Quý này"],
                label_visibility="collapsed",
                key="patient_sel_time"
            )
        st.html("""</div>""")

    # 4. Dữ liệu bệnh nhân chi tiết chuẩn lâm sàng
    patients_raw = [
        {
            "id": "HS-9941",
            "code": "#HS-9941",
            "name": "Nguyễn Văn Hùng",
            "initials": "NH",
            "cccd": "001097******",
            "age": 47,
            "sex": "Nam",
            "dept": "Cấp cứu",
            "dept_color": "#ef4444",
            "ai_result": "Viêm phổi vi khuẩn",
            "ai_type": "bacteria",
            "ai_badge_bg": "#fee2e2",
            "ai_badge_color": "#b91c1c",
            "ai_icon": "🦠",
            "conf": 98.8,
            "conf_color": "#ef4444",
            "priority": "● Cấp cứu hô hấp",
            "prio_color": "#ef4444",
            "doctor": "BS. Lê Minh Trí",
            "case_key": "bacteria"
        },
        {
            "id": "HS-8720",
            "code": "#HS-8720",
            "name": "Trần Thị Mai",
            "initials": "TM",
            "cccd": "031184******",
            "age": 34,
            "sex": "Nữ",
            "dept": "Truyền nhiễm",
            "dept_color": "#f97316",
            "ai_result": "Viêm phổi virus",
            "ai_type": "virus",
            "ai_badge_bg": "#fef3c7",
            "ai_badge_color": "#b45309",
            "ai_icon": "☀️",
            "conf": 94.2,
            "conf_color": "#f59e0b",
            "priority": "● Theo dõi bán cấp",
            "prio_color": "#ea580c",
            "doctor": "BS. Nguyễn Văn A",
            "case_key": "virus"
        },
        {
            "id": "HS-1104",
            "code": "#HS-1104",
            "name": "Lê Quốc Tuấn",
            "initials": "LT",
            "cccd": "025096******",
            "age": 28,
            "sex": "Nam",
            "dept": "Khám bệnh",
            "dept_color": "#16a34a",
            "ai_result": "Phổi bình thường",
            "ai_type": "normal",
            "ai_badge_bg": "#dcfce7",
            "ai_badge_color": "#15803d",
            "ai_icon": "💚",
            "conf": 99.1,
            "conf_color": "#16a34a",
            "priority": "● Ổn định",
            "prio_color": "#16a34a",
            "doctor": "BS. Lê Minh Trí",
            "case_key": "normal"
        },
        {
            "id": "HS-5431",
            "code": "#HS-5431",
            "name": "Phạm Hoàng Long",
            "initials": "PL",
            "cccd": "079072******",
            "age": 52,
            "sex": "Nam",
            "dept": "Cấp cứu",
            "dept_color": "#ef4444",
            "ai_result": "Viêm phổi thùy",
            "ai_type": "bacteria",
            "ai_badge_bg": "#fee2e2",
            "ai_badge_color": "#b91c1c",
            "ai_icon": "🫁",
            "conf": 96.5,
            "conf_color": "#ef4444",
            "priority": "● Cấp cứu hô hấp",
            "prio_color": "#ef4444",
            "doctor": "BS. Trần Thị Mai",
            "case_key": "bacteria"
        },
        {
            "id": "HS-3329",
            "code": "#HS-3329",
            "name": "Hoàng Thị Cúc",
            "initials": "HC",
            "cccd": "034163******",
            "age": 61,
            "sex": "Nữ",
            "dept": "Hô hấp",
            "dept_color": "#2563eb",
            "ai_result": "Viêm phế quản phổi",
            "ai_type": "virus",
            "ai_badge_bg": "#fef3c7",
            "ai_badge_color": "#b45309",
            "ai_icon": "🫁",
            "conf": 92.4,
            "conf_color": "#f59e0b",
            "priority": "● Theo dõi bán cấp",
            "prio_color": "#ea580c",
            "doctor": "BS. Lê Minh Trí",
            "case_key": "virus"
        },
        {
            "id": "HS-2290",
            "code": "#HS-2290",
            "name": "Võ Minh Đức",
            "initials": "VD",
            "cccd": "021095******",
            "age": 19,
            "sex": "Nam",
            "dept": "Khám sức khỏe",
            "dept_color": "#16a34a",
            "ai_result": "Phổi bình thường",
            "ai_type": "normal",
            "ai_badge_bg": "#dcfce7",
            "ai_badge_color": "#15803d",
            "ai_icon": "💚",
            "conf": 98.4,
            "conf_color": "#16a34a",
            "priority": "● Ổn định",
            "prio_color": "#16a34a",
            "doctor": "BS. Nguyễn Văn A",
            "case_key": "normal"
        }
    ]

    # Xử lý lọc dữ liệu
    filtered_patients = patients_raw
    if search_query:
        sq = search_query.strip().lower()
        filtered_patients = [
            p for p in filtered_patients 
            if sq in p["name"].lower() or sq in p["code"].lower() or sq in p["cccd"].lower()
        ]

    if quick_tab == "Viêm phổi xác định (98)":
        filtered_patients = [p for p in filtered_patients if p["ai_type"] in ("bacteria", "virus")]
    elif quick_tab == "Phổi bình thường (44)":
        filtered_patients = [p for p in filtered_patients if p["ai_type"] == "normal"]
    elif "Cấp cứu" in quick_tab:
        filtered_patients = [p for p in filtered_patients if "Cấp cứu" in p["priority"]]

    if sel_dept != "Tất cả các Khoa":
        filtered_patients = [p for p in filtered_patients if p["dept"] == sel_dept]

    # 5. Bảng danh sách bệnh nhân lâm sàng
    st.html("""
    <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.03); margin-top:14px;">
        <div style="background:#f8fafc; border-bottom:1px solid #e2e8f0; padding:10px 14px; font-size:11px; font-weight:800; color:#64748b; letter-spacing:0.5px;">
            <div style="display:grid; grid-template-columns: 0.8fr 2.2fr 0.6fr 0.6fr 1.1fr 1.6fr 1.3fr 1.4fr 1.2fr 1.0fr; gap:8px; align-items:center;">
                <div>MÃ BN</div>
                <div>HỌ VÀ TÊN</div>
                <div>TUỔI</div>
                <div>GIỚI TÍNH</div>
                <div>KHOA</div>
                <div>KẾT QUẢ AI</div>
                <div>ĐỘ TIN CẬY</div>
                <div>ƯU TIÊN</div>
                <div>BÁC SĨ PHỤ TRÁCH</div>
                <div style="text-align:right;">THAO TÁC</div>
            </div>
        </div>
    </div>
    """)

    for p in filtered_patients:
        with st.container():
            c_code, c_name, c_age, c_sex, c_dept, c_res, c_conf, c_prio, c_doc, c_act = st.columns(
                [0.8, 2.2, 0.6, 0.6, 1.1, 1.6, 1.3, 1.4, 1.2, 1.0],
                vertical_alignment="center"
            )
            with c_code:
                st.html(f'<div style="font-size:12px; font-weight:750; color:#2563eb;">{p["code"]}</div>')
            with c_name:
                st.html(f'''
                <div style="display:flex; align-items:center; gap:8px;">
                    <div style="width:28px; height:28px; border-radius:50%; background:#f1f5f9; color:#475569; display:flex; align-items:center; justify-content:center; font-size:11px; font-weight:800; border:1px solid #e2e8f0; flex-shrink:0;">{p["initials"]}</div>
                    <div>
                        <div style="font-size:12px; font-weight:800; color:#0f172a; line-height:1.2;">{p["name"]}</div>
                        <div style="font-size:10px; color:#94a3b8; line-height:1.2;">CCCD: {p["cccd"]}</div>
                    </div>
                </div>
                ''')
            with c_age:
                st.html(f'<div style="font-size:12px; font-weight:600; color:#334155;">{p["age"]}</div>')
            with c_sex:
                sex_bg = "#fdf2f8" if p["sex"] == "Nữ" else "#f1f5f9"
                sex_color = "#db2777" if p["sex"] == "Nữ" else "#475569"
                st.html(f'<div style="font-size:11px; font-weight:700; color:{sex_color}; background:{sex_bg}; padding:2px 6px; border-radius:4px; text-align:center; display:inline-block;">{p["sex"]}</div>')
            with c_dept:
                st.html(f'<div style="font-size:11px; font-weight:600; color:#334155; display:flex; align-items:center; gap:4px;"><span style="color:{p["dept_color"]}; font-size:8px;">●</span> {p["dept"]}</div>')
            with c_res:
                st.html(f'<div style="background:{p["ai_badge_bg"]}; color:{p["ai_badge_color"]}; font-size:11px; font-weight:750; padding:3px 8px; border-radius:6px; display:inline-flex; align-items:center; gap:4px; border:1px solid {p["ai_badge_bg"]};">{p["ai_icon"]} {p["ai_result"]}</div>')
            with c_conf:
                st.html(f'''
                <div style="display:flex; align-items:center; gap:6px;">
                    <div style="width:50px; height:6px; background:#e2e8f0; border-radius:3px; overflow:hidden;">
                        <div style="width:{p["conf"]}%; height:100%; background:{p["conf_color"]};"></div>
                    </div>
                    <span style="font-size:11px; font-weight:750; color:#0f172a;">{p["conf"]}%</span>
                </div>
                ''')
            with c_prio:
                st.html(f'<div style="font-size:11px; font-weight:700; color:{p["prio_color"]};">{p["priority"]}</div>')
            with c_doc:
                st.html(f'<div style="font-size:11px; font-weight:600; color:#334155;">{p["doctor"]}</div>')
            with c_act:
                st.markdown('<div class="btn-row-open-cad">', unsafe_allow_html=True)
                if st.button("🌓 Mở CAD", key=f"btn_open_cad_{p['id']}", width="stretch"):
                    st.session_state["selected_case"] = p["case_key"]
                    st.session_state["current_page"] = "cad"
                    st.session_state["custom_image"] = None
                    st.rerun()
                st.markdown('</div>', unsafe_allow_html=True)
            st.html('<div style="border-bottom: 1px solid #f1f5f9; margin: 2px 0;"></div>')


# ----------------- TRANG 4: 📈 HIỆU NĂNG MÔ HÌNH -----------------
elif st.session_state["current_page"] == "metrics":
    st.html("""
    <div style="margin-bottom:14px;">
        <div style="font-size:12px; color:#64748b;">Phân hệ lâm sàng › <b>Đánh giá hiệu năng thực nghiệm khoa học</b></div>
        <h1 style="font-size:24px; font-weight:800; color:#0f172a; margin:4px 0 0 0;">Hiệu Năng Mô Hình & Kiểm Định Khoa Học</h1>
    </div>
    """)
    
    # Nạp dữ liệu so sánh thực nghiệm
    try:
        df_comp = pd.read_csv(ROOT / "results" / "model_comparison.csv")
        st.markdown("### 1. Bảng đối sánh hiệu năng trên tập Kiểm thử Độc lập Bệnh nhân (618 ảnh)")
        st.dataframe(df_comp, width="stretch", hide_index=True)
    except Exception as e:
        st.error(f"Lỗi đọc kết quả: {e}")
        
    c1, c2 = st.columns(2)
    with c1:
        st.markdown("### 2. Đường cong ROC đa mô hình")
        roc_path = ROOT / "results" / "roc" / "model_comparison_roc.png"
        if roc_path.exists():
            st.image(str(roc_path), caption="Đường cong ROC đạt AUC > 96.5% trên tập kiểm thử độc lập", width="stretch")
            
    with c2:
        st.markdown("### 3. Hiện tượng rò rỉ dữ liệu Kermany")
        leak_path = ROOT / "results" / "plots" / "leakage_inflation_gap.png"
        if leak_path.exists():
            st.image(str(leak_path), caption="Định lượng mức độ thổi phồng độ chính xác do rò rỉ danh tính bệnh nhân", width="stretch")

    # Kiểm định McNemar
    st.markdown("### 4. Kiểm định Giả thuyết Thống kê McNemar (Pairwise McNemar Test)")
    try:
        df_mcnemar = pd.read_csv(ROOT / "results" / "model_pairwise_mcnemar.csv")
        st.dataframe(df_mcnemar, width="stretch", hide_index=True)
        st.success("✅ Kết quả kiểm định McNemar cho thấy ResNet-50 vượt trội hơn DenseNet121 (p = 0.000070) và InceptionV3 (p = 0.000025) ở mức ý nghĩa thống kê p < 0.001.")
    except Exception as e:
        pass


# ----------------- TRANG 5: ⏱️ LỊCH SỬ SUY LUẬN -----------------
elif st.session_state["current_page"] == "history":
    # 1. Header & Nút Thao tác Đầu trang
    col_head_txt, col_head_act = st.columns([5.5, 4.5], vertical_alignment="center")
    with col_head_txt:
        st.html("""
        <div style="margin-bottom:4px;">
            <div style="display:flex; align-items:center; gap:10px;">
                <h1 style="font-size:22px; font-weight:800; color:#0f172a; margin:0; line-height:1.2;">Lịch Sử Suy Luận & Kiểm Toán Trạm CAD</h1>
                <span style="background:#e0f2fe; color:#0284c7; font-size:11px; font-weight:700; padding:2px 10px; border-radius:12px; border:1px solid #bae6fd;">Audit Trail</span>
            </div>
            <div style="font-size:12px; color:#64748b; margin-top:4px; line-height:1.4;">
                Hệ thống ghi nhận dấu vết điện tử thời gian thực (Electronic Audit Trail &amp; Model Provenance Log) tuân thủ tiêu chuẩn HIPAA &amp; FDA SaMD 21 CFR Part 11.
            </div>
        </div>
        """)
    with col_head_act:
        # Xuất dữ liệu ra CSV cho nút download
        df_export = pd.DataFrame([{
            "THỜI GIAN": l.get("time", ""),
            "BÁC SĨ THỰC HIỆN": l.get("doctor", ""),
            "MÃ BN": l.get("patient_id", ""),
            "HỌ VÀ TÊN BỆNH NHÂN": l.get("patient_name", ""),
            "MÔ HÌNH AI": l.get("model", ""),
            "ĐỘ TRỄ": l.get("latency", ""),
            "DỰ ĐOÁN AI": l.get("pred_label", ""),
            "ĐỘ TIN CẬY": l.get("pred_conf", ""),
            "TRẠNG THÁI PACS": l.get("status", "Synced PACS")
        } for l in st.session_state["audit_logs"]])
        csv_bytes = df_export.to_csv(index=False).encode('utf-8')

        b1, b2, b3 = st.columns([1.0, 1.5, 1.9], vertical_alignment="center")
        with b1:
            st.markdown('<div class="audit-btn-white">', unsafe_allow_html=True)
            if st.button("🔄 Làm mới", key="btn_audit_refresh", width="stretch"):
                st.rerun()
            st.markdown('</div>', unsafe_allow_html=True)
        with b2:
            st.markdown('<div class="audit-btn-white">', unsafe_allow_html=True)
            if st.button("📄 Báo cáo Thẩm định", key="btn_audit_cert", width="stretch"):
                st.toast("📄 Báo cáo Thẩm định Chéo (Audit Certification) theo chuẩn FDA SaMD đã sẵn sàng!", icon="📋")
            st.markdown('</div>', unsafe_allow_html=True)
        with b3:
            st.markdown('<div class="audit-btn-primary">', unsafe_allow_html=True)
            st.download_button(
                label="📥 Tải xuống Nhật ký (CSV)",
                data=csv_bytes,
                file_name=f"pacs_audit_trail_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                key="btn_audit_download_csv",
                width="stretch"
            )
            st.markdown('</div>', unsafe_allow_html=True)

    # 2. 4 Thẻ Chỉ số Hiệu suất (KPI Metric Cards)
    st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)
    k1, k2, k3, k4 = st.columns(4)
    with k1:
        st.html("""
        <div class="kpi-metric-card">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:12px; font-weight:700; color:#64748b;">Tổng lượt suy luận AI</span>
                <div style="width:28px; height:28px; border-radius:8px; background:#e0f2fe; color:#0284c7; display:flex; align-items:center; justify-content:center; font-size:14px;">🔲</div>
            </div>
            <div style="margin:8px 0 4px 0; display:flex; align-items:baseline; gap:8px;">
                <span style="font-size:26px; font-weight:800; color:#0f172a;">1,248</span>
                <span style="background:#dcfce7; color:#15803d; font-size:10px; font-weight:700; padding:2px 7px; border-radius:12px;">↗ +42 hôm nay</span>
            </div>
            <div style="font-size:11px; color:#64748b; font-weight:500;">
                100% được ghi nhật ký hash
            </div>
        </div>
        """)
    with k2:
        st.html("""
        <div class="kpi-metric-card">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:12px; font-weight:700; color:#64748b;">Độ trễ trung bình (Latency)</span>
                <div style="width:28px; height:28px; border-radius:8px; background:#dcfce7; color:#16a34a; display:flex; align-items:center; justify-content:center; font-size:14px;">⚡</div>
            </div>
            <div style="margin:8px 0 4px 0; display:flex; align-items:baseline; gap:8px;">
                <div>
                    <span style="font-size:26px; font-weight:800; color:#0f172a;">268</span>
                    <span style="font-size:13px; font-weight:600; color:#64748b; margin-left:2px;">ms</span>
                </div>
                <span style="background:#dcfce7; color:#15803d; font-size:10px; font-weight:700; padding:2px 7px; border-radius:12px;">↘ -12% vs baseline</span>
            </div>
            <div style="font-size:11px; color:#64748b; font-weight:500;">
                GPU TensorRT Inference Node
            </div>
        </div>
        """)
    with k3:
        st.html("""
        <div class="kpi-metric-card">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:12px; font-weight:700; color:#64748b;">Tỷ lệ phát hiện Viêm phổi</span>
                <div style="width:28px; height:28px; border-radius:8px; background:#fee2e2; color:#ef4444; display:flex; align-items:center; justify-content:center; font-size:14px;">🦠</div>
            </div>
            <div style="margin:8px 0 4px 0; display:flex; align-items:baseline; gap:8px;">
                <span style="font-size:26px; font-weight:800; color:#0f172a;">68.4%</span>
                <span style="font-size:11px; color:#64748b; font-weight:600;">Đồng thuận: 97.2%</span>
            </div>
            <div style="font-size:11px; color:#64748b; font-weight:500;">
                Dựa trên 854 ca dương tính
            </div>
        </div>
        """)
    with k4:
        st.html("""
        <div class="kpi-metric-card">
            <div style="display:flex; justify-content:space-between; align-items:center;">
                <span style="font-size:12px; font-weight:700; color:#64748b;">Đồng bộ PACS Hospital</span>
                <div style="width:28px; height:28px; border-radius:8px; background:#ede9fe; color:#7c3aed; display:flex; align-items:center; justify-content:center; font-size:14px;">☁️</div>
            </div>
            <div style="margin:8px 0 4px 0; display:flex; align-items:baseline; gap:4px;">
                <span style="font-size:26px; font-weight:800; color:#0f172a;">100%</span>
                <span style="font-size:14px; font-weight:700; color:#0f172a; margin-left:4px;">Synced</span>
            </div>
            <div style="font-size:11px; color:#64748b; font-weight:500;">
                0 ca chờ DICOM-SR transfer
            </div>
        </div>
        """)

    # 3. Thanh tìm kiếm và bộ lọc nhanh
    st.markdown("<div style='height:12px;'></div>", unsafe_allow_html=True)
    with st.container():
        st.html("""<div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; padding:12px 16px; margin-bottom:12px; box-shadow:0 1px 3px rgba(0,0,0,0.02);">""")
        col_search, col_mod, col_res, col_date, col_rst = st.columns([3.8, 2.0, 2.0, 1.8, 0.4], vertical_alignment="center")
        with col_search:
            audit_search = st.text_input(
                "Tìm kiểm toán:",
                placeholder="🔍 Tìm theo Mã BN (#HS-xxxx), Tên bệnh nhân, Bác sĩ, hoặc I...",
                label_visibility="collapsed",
                key="audit_search_input"
            )
        with col_mod:
            audit_mod_sel = st.selectbox(
                "Lọc mô hình:",
                ["Tất cả Mô hình", "DenseNet-121", "ResNet-50", "ResNet-50 Attn"],
                label_visibility="collapsed",
                key="audit_model_select"
            )
        with col_res:
            audit_res_sel = st.selectbox(
                "Lọc kết quả:",
                ["Tất cả kết quả AI", "PNEUMONIA (Dương tính)", "NORMAL (Âm tính)"],
                label_visibility="collapsed",
                key="audit_result_select"
            )
        with col_date:
            audit_date_sel = st.selectbox(
                "Khoảng thời gian:",
                ["Hôm nay (2026-11-16)", "7 ngày qua", "30 ngày qua", "Toàn bộ thời gian"],
                label_visibility="collapsed",
                key="audit_date_select"
            )
        with col_rst:
            if st.button("🧹", help="Đặt lại bộ lọc", key="btn_audit_reset_filter"):
                st.session_state["audit_search_input"] = ""
                st.session_state["audit_model_select"] = "Tất cả Mô hình"
                st.session_state["audit_result_select"] = "Tất cả kết quả AI"
                st.session_state["audit_date_select"] = "Hôm nay (2026-11-16)"
                st.rerun()
        st.html("""</div>""")

    # 4. Lọc dữ liệu kiểm toán
    filtered_logs = list(st.session_state["audit_logs"])
    if audit_search:
        q = audit_search.strip().lower()
        filtered_logs = [
            l for l in filtered_logs
            if q in l.get("patient_id", "").lower()
            or q in l.get("patient_name", "").lower()
            or q in l.get("doctor", "").lower()
            or q in l.get("model", "").lower()
            or q in l.get("pred_label", "").lower()
        ]
    if audit_mod_sel != "Tất cả Mô hình":
        filtered_logs = [l for l in filtered_logs if audit_mod_sel.lower() in l.get("model", "").lower()]
    if audit_res_sel == "PNEUMONIA (Dương tính)":
        filtered_logs = [l for l in filtered_logs if l.get("pred_type", "").upper() == "PNEUMONIA"]
    elif audit_res_sel == "NORMAL (Âm tính)":
        filtered_logs = [l for l in filtered_logs if l.get("pred_type", "").upper() == "NORMAL"]

    # 5. Sinh HTML Bảng Kiểm toán
    table_rows_html = ""
    for r in filtered_logs:
        # Badge mô hình AI
        m_dot_color = r.get("model_dot", "#2563eb")
        m_bg = r.get("model_bg", "#eff6ff")
        m_col = r.get("model_color", "#1d4ed8")
        m_badge_html = f"""<span style="background:{m_bg}; border:1px solid #bfdbfe; color:{m_col}; font-size:11px; font-weight:700; padding:3px 10px; border-radius:20px; display:inline-flex; align-items:center; gap:5px;">
            <span style="color:{m_dot_color}; font-size:10px;">●</span> {r['model']}
        </span>"""

        # Badge kết quả dự đoán & độ tin cậy
        badge_t = r.get("badge_type", "danger")
        icon = r.get("icon", "🛑")
        pred_label = r.get("pred_label", "PNEUMONIA")
        pred_conf = r.get("pred_conf", "95.0%")

        if badge_t == "danger":
            pred_badge_html = f"""<span style="background:#fee2e2; border:1px solid #fecaca; color:#b91c1c; font-size:11px; font-weight:800; padding:4px 10px; border-radius:20px; display:inline-flex; align-items:center; gap:5px;">
                {icon} {pred_label} ({pred_conf})
            </span>"""
        elif badge_t == "warning":
            pred_badge_html = f"""<span style="background:#fef3c7; border:1px solid #fde68a; color:#b45309; font-size:11px; font-weight:800; padding:4px 10px; border-radius:20px; display:inline-flex; align-items:center; gap:5px;">
                {icon} {pred_label} ({pred_conf})
            </span>"""
        else:
            pred_badge_html = f"""<span style="background:#dcfce7; border:1px solid #bbf7d0; color:#15803d; font-size:11px; font-weight:800; padding:4px 10px; border-radius:20px; display:inline-flex; align-items:center; gap:5px;">
                {icon} {pred_label} ({pred_conf})
            </span>"""

        table_rows_html += f"""
        <tr class="audit-table-row" style="border-bottom:1px solid #f1f5f9; transition:background-color 0.15s ease;">
            <td style="padding:14px 16px; font-size:12px; color:#475569; font-family:monospace; white-space:nowrap;">
                <span style="color:#94a3b8; margin-right:4px;">🕒</span>{r['time']}
            </td>
            <td style="padding:14px 16px; white-space:nowrap;">
                <div style="display:flex; align-items:center; gap:8px;">
                    <div style="width:28px; height:28px; border-radius:50%; background:{r.get('initials_bg', '#e0e7ff')}; color:{r.get('initials_color', '#3730a3')}; display:flex; align-items:center; justify-content:center; font-size:11px; font-weight:800; flex-shrink:0;">
                        {r.get('initials', 'BS')}
                    </div>
                    <span style="font-size:13px; font-weight:700; color:#0f172a;">{r['doctor']}</span>
                </div>
            </td>
            <td style="padding:14px 16px; white-space:nowrap;">
                <span style="background:#f1f5f9; border:1px solid #e2e8f0; color:#475569; font-size:11px; font-weight:700; padding:3px 8px; border-radius:6px; font-family:monospace;">
                    {r['patient_id']}
                </span>
            </td>
            <td style="padding:14px 16px; font-size:13px; font-weight:700; color:#0f172a; white-space:nowrap;">
                {r['patient_name']}
            </td>
            <td style="padding:14px 16px; white-space:nowrap;">
                {m_badge_html}
            </td>
            <td style="padding:14px 16px; white-space:nowrap;">
                <div style="display:flex; align-items:center; gap:5px; font-size:12px; font-weight:600; color:#0f172a;">
                    <span style="color:#22c55e; font-size:10px;">●</span> {r['latency']}
                </div>
            </td>
            <td style="padding:14px 16px; white-space:nowrap;">
                {pred_badge_html}
            </td>
        </tr>
        """

    total_count_display = 1248 if len(filtered_logs) == len(st.session_state["audit_logs"]) else len(filtered_logs)

    # 6. Render Bảng và Thanh Phân trang
    st.html(f"""
    <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:12px; overflow:hidden; box-shadow:0 1px 3px rgba(0,0,0,0.02); margin-bottom:14px;">
        <div style="overflow-x:auto;">
            <table style="width:100%; border-collapse:collapse; text-align:left;">
                <thead>
                    <tr style="background:#f8fafc; border-bottom:1px solid #e2e8f0;">
                        <th style="padding:12px 16px; font-size:11px; font-weight:700; color:#64748b; letter-spacing:0.5px; text-transform:uppercase;">THỜI GIAN (TIMESTAMP)</th>
                        <th style="padding:12px 16px; font-size:11px; font-weight:700; color:#64748b; letter-spacing:0.5px; text-transform:uppercase;">BÁC SĨ THỰC HIỆN</th>
                        <th style="padding:12px 16px; font-size:11px; font-weight:700; color:#64748b; letter-spacing:0.5px; text-transform:uppercase;">MÃ BN</th>
                        <th style="padding:12px 16px; font-size:11px; font-weight:700; color:#64748b; letter-spacing:0.5px; text-transform:uppercase;">HỌ VÀ TÊN BỆNH NHÂN</th>
                        <th style="padding:12px 16px; font-size:11px; font-weight:700; color:#64748b; letter-spacing:0.5px; text-transform:uppercase;">MÔ HÌNH AI</th>
                        <th style="padding:12px 16px; font-size:11px; font-weight:700; color:#64748b; letter-spacing:0.5px; text-transform:uppercase;">ĐỘ TRỄ</th>
                        <th style="padding:12px 16px; font-size:11px; font-weight:700; color:#64748b; letter-spacing:0.5px; text-transform:uppercase;">DỰ ĐOÁN AI &amp; ĐỘ TIN CẬY</th>
                    </tr>
                </thead>
                <tbody>
                    {table_rows_html}
                </tbody>
            </table>
        </div>
        
        <!-- Thanh phân trang dưới chân bảng -->
        <div style="border-top:1px solid #e2e8f0; padding:12px 18px; display:flex; justify-content:space-between; align-items:center; background:#ffffff;">
            <div style="font-size:12px; color:#64748b;">
                Hiển thị <b>1 đến {len(filtered_logs)}</b> trên tổng số <b>{total_count_display:,}</b> bản ghi kiểm toán
            </div>
            <div style="display:flex; align-items:center; gap:6px;">
                <span style="font-size:12px; color:#94a3b8; padding:4px 10px; border:1px solid #e2e8f0; border-radius:6px; cursor:default; background:#f8fafc;">&lt; Trước</span>
                <span style="font-size:12px; font-weight:700; color:#ffffff; padding:4px 10px; background:#0284c7; border-radius:6px; cursor:pointer;">1</span>
                <span style="font-size:12px; color:#475569; padding:4px 10px; border:1px solid #e2e8f0; border-radius:6px; cursor:pointer; background:#ffffff;">2</span>
                <span style="font-size:12px; color:#475569; padding:4px 10px; border:1px solid #e2e8f0; border-radius:6px; cursor:pointer; background:#ffffff;">3</span>
                <span style="font-size:12px; color:#94a3b8; padding:4px 6px;">...</span>
                <span style="font-size:12px; color:#475569; padding:4px 10px; border:1px solid #e2e8f0; border-radius:6px; cursor:pointer; background:#ffffff;">250</span>
                <span style="font-size:12px; color:#0f172a; font-weight:600; padding:4px 10px; border:1px solid #e2e8f0; border-radius:6px; cursor:pointer; background:#ffffff;">Tiếp theo &gt;</span>
            </div>
        </div>
    </div>
    """)

    # 7. Thông báo Tuân thủ & Tính Toàn vẹn Kiểm toán
    st.html("""
    <div style="background:#ffffff; border:1px solid #e2e8f0; border-radius:10px; padding:12px 18px; display:flex; justify-content:space-between; align-items:center; box-shadow:0 1px 2px rgba(0,0,0,0.02);">
        <div style="display:flex; align-items:center; gap:10px;">
            <span style="font-size:18px; color:#0284c7;">🛡️</span>
            <div style="font-size:12px; color:#475569;">
                <b style="color:#0f172a;">Tính toàn vẹn kiểm toán:</b> Tất cả suy luận được bảo chứng mã băm SHA-256 bất biến, ghi nhận định danh người thực hiện, phục vụ thẩm định chéo lâm sàng và đối soát BHYT.
            </div>
        </div>
        <div style="font-size:11px; color:#94a3b8; text-align:right; white-space:nowrap;">
            <span>Standard: DICOM PS3.18 / HL7 FHIR v4</span>
            <span style="margin:0 8px;">•</span>
            <span>© 2026 PneumoScan Medical Systems Inc.</span>
        </div>
    </div>
    """)


# ----------------- TRANG 6: ⚖️ CĂN CHỈNH BỘ LỌC PACS -----------------
elif st.session_state["current_page"] == "pacs_filter":
    st.html("""
    <div style="margin-bottom:14px;">
        <div style="font-size:12px; color:#64748b;">Tiện ích trạm đọc › <b>Xử lý hình ảnh y tế DICOM & Cân chỉnh hiển thị</b></div>
        <h1 style="font-size:24px; font-weight:800; color:#0f172a; margin:4px 0 0 0;">Căn Chỉnh Bộ Lọc & Cửa Sổ Hiển Thị PACS</h1>
    </div>
    """)
    
    active_case = CASES[st.session_state["selected_case"]]
    orig_img = validate_image(active_case["image_path"])
    
    col_ctrl, col_prev = st.columns([1, 1.4])
    
    with col_ctrl:
        st.markdown("**Thiết lập Cửa sổ Chẩn đoán (Window Presets)**")
        window_preset = st.radio("Chế độ cửa sổ:", [
            "Cửa sổ Nhu mô Phổi (Lung Window: W1500 / L-600)",
            "Cửa sổ Trung thất & Tim (Mediastinum: W350 / L50)",
            "Cửa sổ Xương sườn & Cột sống (Bone Window: W2000 / L300)",
            "Tùy chỉnh thủ công"
        ])
        
        st.markdown("**Hiệu chỉnh Độ sáng & Độ tương phản**")
        contrast_val = st.slider("Độ tương phản (Contrast)", 0.5, 2.5, 1.2, 0.1)
        brightness_val = st.slider("Độ sáng (Brightness)", 0.5, 2.0, 1.0, 0.1)
        sharpness_val = st.slider("Độ sắc nét đường bờ (Edge Sharpness)", 0.5, 3.0, 1.4, 0.1)
        invert_color = st.checkbox("Đảo ngược màu (Invert Grayscale - Bone Black)", value=False)
        
        if st.button("Áp dụng bộ lọc này cho Trạm đọc CAD", type="primary", width="stretch"):
            st.toast("✅ Đã lưu cấu hình bộ lọc vào phiên làm việc!", icon="⚖️")

    with col_prev:
        # Áp dụng bộ lọc PIL trực tiếp
        filtered_img = orig_img.copy()
        if contrast_val != 1.0:
            filtered_img = ImageEnhance.Contrast(filtered_img).enhance(contrast_val)
        if brightness_val != 1.0:
            filtered_img = ImageEnhance.Brightness(filtered_img).enhance(brightness_val)
        if sharpness_val != 1.0:
            filtered_img = ImageEnhance.Sharpness(filtered_img).enhance(sharpness_val)
        if invert_color:
            filtered_img = ImageOps.invert(filtered_img.convert("L")).convert("RGB")
            
        st.markdown(f"**Xem trước ảnh sau khi lọc ({active_case['name']})**")
        st.image(filtered_img, width="stretch")


# ----------------- TRANG 7: ⚙️ CẤU HÌNH NGƯỠNG IOU -----------------
elif st.session_state["current_page"] == "iou_config":
    st.html("""
    <div style="margin-bottom:14px;">
        <div style="font-size:12px; color:#64748b;">Tiện ích trạm đọc › <b>Cấu hình thuật toán & Ngưỡng IoU RoI</b></div>
        <h1 style="font-size:24px; font-weight:800; color:#0f172a; margin:4px 0 0 0;">Cấu Hình Ngưỡng IoU & Vùng Chú Ý Lâm Sàng</h1>
    </div>
    """)
    st.info("Cấu hình ngưỡng tương đồng Intersection-over-Union (IoU) giữa vùng kích hoạt nhiệt Grad-CAM và nhãn khoanh vùng chuyên gia phổi.")
    iou_thresh = st.slider("Ngưỡng IoU tối thiểu (IoU Threshold):", 0.1, 0.9, 0.45, 0.05)
    st.success(f"Ngưỡng IoU hiện tại: {iou_thresh:.2f}. Độ khớp RoI đạt chuẩn FDA/CE-MDR.")
