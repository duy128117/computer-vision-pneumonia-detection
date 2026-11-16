# 🫁 PneumoScan AI - Chest X-Ray Pneumonia Detection CAD

A Deep Learning Computer-Aided Diagnosis (CAD) system and clinical workstation for automated pediatric pneumonia screening and Explainable AI (Grad-CAM) from Chest X-rays.

[![Python 3.11](https://img.shields.io/badge/Python-3.11-3776AB.svg?style=flat&logo=python&logoColor=white)](https://www.python.org/)
[![PyTorch 2.6](https://img.shields.io/badge/PyTorch-2.6-EE4C2C.svg?style=flat&logo=pytorch&logoColor=white)](https://pytorch.org/)
[![Streamlit](https://img.shields.io/badge/Streamlit-App-FF4B4B.svg?style=flat&logo=streamlit&logoColor=white)](https://streamlit.io/)
[![License: MIT](https://img.shields.io/badge/License-MIT-blue.svg)](LICENSE)

---

## ⚡ Highlights

- **Deep Learning Architectures:** Evaluates ResNet-50, DenseNet-121, InceptionV3, and ResNet-50 with Spatial Attention Gate.
- **High Sensitivity:** Reaches **99.48% Sensitivity** and **90.61% Accuracy** on a patient-disjoint test set ($N = 618$ images).
- **Explainable AI (Grad-CAM):** Real-time activation heatmaps (Jet, Turbo, Inferno) pinpointing pulmonary lesions and consolidations.
- **Clinical CAD Workstation:** Full-featured Streamlit interface featuring PACS windowing presets, patient record tracking, and electronic audit trail.

---

## 📊 Benchmark Results

Evaluated on the patient-disjoint test distribution (387 Pneumonia, 231 Normal):

| Model | Accuracy | Sensitivity (Recall) | Specificity | F1-Score | AUROC |
|:---|:---:|:---:|:---:|:---:|:---:|
| **ResNet-50 (Best)** | **90.61%** | **99.48%** | **75.76%** | **0.9300** | **0.9679** |
| **ResNet-50 Attention** | 88.03% | 99.22% | 69.26% | 0.9121 | 0.9624 |
| **DenseNet-121** | 86.57% | 99.74% | 64.50% | 0.9029 | 0.9578 |
| **InceptionV3** | 86.08% | 100.0% | 62.77% | 0.9000 | 0.9629 |

---

## 🖥️ Workstation Features

- **🫁 CAD Diagnosis:** Upload custom CXR images or select clinical presets; instant inference with SpO2, Body Temp, and WBC vitals cross-checking.
- **📊 Grad-CAM XAI:** Triplet visualization (Native CXR, Activation Heatmap, Anatomical Fusion) with alpha opacity slider.
- **📁 Patient Records:** Management dashboard for 142 patient cases with clinical triage priority levels.
- **📈 Benchmark Metrics:** Interactive model performance comparison, confusion matrices, and ROC curves.
- **⏱️ Audit Trail:** Tamper-evident electronic logging with SHA-256 verification and live CSV export.
- **⚖️ PACS Windowing:** Medical display presets (Lung Window, Mediastinum, Bone) with contrast and sharpness tuning.

---

## 🚀 Quickstart

### 1. Setup Environment
```bash
git clone https://github.com/duy128117/computer-vision-pneumonia-detection.git
cd computer-vision-pneumonia-detection

# Create and activate virtual environment
python -m venv .venv
.\.venv\Scripts\Activate.ps1  # Windows PowerShell
# source .venv/bin/activate    # Linux/macOS
```

### 2. Install Dependencies
```bash
pip install -r requirements.txt
```

### 3. Run Application
```bash
streamlit run app/app.py
```
Open **`http://localhost:8501`** in your browser.

---

## 🛠️ Tech Stack

- **Frameworks:** PyTorch 2.6, Torchvision, Streamlit
- **Image Processing & XAI:** OpenCV, Pillow, Grad-CAM
- **Data & Evaluation:** NumPy, Pandas, Scikit-learn, Matplotlib

---

## 👥 Authors

- **Mai Phước Thịnh** & **Nguyễn Đình Duy** (Group 03)
- **Advisor:** Dr. Huỳnh Hữu Hưng
- **Course:** Advanced Computer Vision (*Thị giác máy tính nâng cao*)
