# -*- coding: utf-8 -*-
"""
STEP 4: Quantitative Grad-CAM Energy-in-Lung-RoI Ratio Analysis
Measures the percentage of Grad-CAM saliency energy located inside the pulmonary RoI
versus non-pulmonary shortcut regions (clavicle, neck, peripheral borders, machine markers)
across ResNet-50, InceptionV3, and ResNet50-LungRoI-Attention on the clean test set.
"""

import os
import sys
import io
import json
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import numpy as np
import pandas as pd
import torch
from PIL import Image
from torchvision import transforms

from src.models import build_model, load_checkpoint
from src.gradcam import gradcam
from src.utils import device

def create_lung_roi_mask(size=(224, 224)):
    """
    Standard pediatric thoracic pulmonary field mask.
    Lungs occupy roughly y: [15% to 85%], x: [12% to 47%] (left) & [53% to 88%] (right).
    Excludes clavicle/neck (y < 15%), sub-diaphragm (y > 85%), mediastinum/spine (47% to 53%),
    and lateral borders/artifacts (x < 12% or x > 88%).
    """
    h, w = size
    mask = np.zeros((h, w), dtype=np.float32)
    
    y_min, y_max = int(h * 0.15), int(h * 0.85)
    
    # Left Lung Field
    x_l_min, x_l_max = int(w * 0.12), int(w * 0.47)
    mask[y_min:y_max, x_l_min:x_l_max] = 1.0
    
    # Right Lung Field
    x_r_min, x_r_max = int(w * 0.53), int(w * 0.88)
    mask[y_min:y_max, x_r_min:x_r_max] = 1.0
    
    return mask

def compute_energy_ratios(heat, lung_mask):
    """
    heat: 2D numpy array [0, 1]
    lung_mask: 2D binary numpy array (1 in lung, 0 outside)
    """
    # Resize heat to mask shape if different
    if heat.shape != lung_mask.shape:
        heat_img = Image.fromarray(heat)
        heat = np.asarray(heat_img.resize((lung_mask.shape[1], lung_mask.shape[0]), Image.Resampling.BILINEAR))
        
    total_energy = float(np.sum(heat))
    if total_energy <= 1e-12:
        return 0.0, 0.0
        
    lung_energy = float(np.sum(heat * lung_mask))
    peripheral_energy = total_energy - lung_energy
    
    ratio_lung = lung_energy / total_energy
    ratio_peripheral = peripheral_energy / total_energy
    return ratio_lung, ratio_peripheral

def run_step4(sample_limit=None):
    print("=" * 75)
    print("STEP 4: QUANTITATIVE GRAD-CAM ENERGY-IN-LUNG-ROI RATIO ANALYSIS")
    print("=" * 75)
    
    target_device = device()
    lung_mask = create_lung_roi_mask((224, 224))
    
    manifest_path = ROOT / "results" / "manifest.csv"
    manifest = pd.read_csv(manifest_path)
    test_df = manifest[manifest.split == "test"].copy().reset_index(drop=True)
    
    if sample_limit is not None and sample_limit < len(test_df):
        test_df = test_df.sample(sample_limit, random_state=42).reset_index(drop=True)
        print(f"Evaluating on {len(test_df)} sampled test images...")
    else:
        print(f"Evaluating on all {len(test_df)} clean test images...")
        
    models_to_test = [
        ("resnet50", ROOT / "checkpoints" / "best_resnet50.pth", "ResNet-50 Baseline"),
        ("inceptionv3", ROOT / "checkpoints" / "best_inceptionv3.pth", "Inception-V3"),
        ("resnet50_attention", ROOT / "checkpoints" / "best_resnet50_attention.pth", "ResNet50_LungRoI_Attention")
    ]
    
    preprocess = transforms.Compose([
        transforms.Resize((224, 224)),
        transforms.ToTensor(),
        transforms.Normalize(mean=[0.485, 0.456, 0.406], std=[0.229, 0.224, 0.225]),
    ])
    
    results = {}
    
    for model_name, ckpt_path, display_name in models_to_test:
        if not ckpt_path.exists():
            print(f"Warning: Checkpoint not found at {ckpt_path}, skipping {display_name}")
            continue
            
        print(f"\nProcessing {display_name}...")
        model, _ = load_checkpoint(ckpt_path, target_device)
        model.eval()
        
        lung_ratios = []
        peripheral_ratios = []
        lung_ratios_pneumonia = []
        lung_ratios_normal = []
        
        for i, row in test_df.iterrows():
            img_path = Path(row["path"])
            label = int(row["label"])
            if not img_path.exists():
                continue
                
            try:
                img = Image.open(img_path).convert("RGB")
                tensor = preprocess(img).unsqueeze(0).to(target_device)
                
                # Compute Grad-CAM
                prob, heat, target = gradcam(model, tensor)
                r_lung, r_periph = compute_energy_ratios(heat, lung_mask)
                
                lung_ratios.append(r_lung)
                peripheral_ratios.append(r_periph)
                if label == 1:
                    lung_ratios_pneumonia.append(r_lung)
                else:
                    lung_ratios_normal.append(r_lung)
            except Exception as e:
                continue
                
        lung_ratios = np.array(lung_ratios)
        peripheral_ratios = np.array(peripheral_ratios)
        
        mean_lung = float(np.mean(lung_ratios))
        std_lung = float(np.std(lung_ratios))
        mean_periph = float(np.mean(peripheral_ratios))
        std_periph = float(np.std(peripheral_ratios))
        mean_lung_pneu = float(np.mean(lung_ratios_pneumonia)) if len(lung_ratios_pneumonia) > 0 else 0.0
        mean_lung_norm = float(np.mean(lung_ratios_normal)) if len(lung_ratios_normal) > 0 else 0.0
        
        print(f"  {display_name}:")
        print(f"    Energy Inside Lung RoI:       {mean_lung*100:.2f}% ± {std_lung*100:.2f}%")
        print(f"    Peripheral / Shortcut Energy: {mean_periph*100:.2f}% ± {std_periph*100:.2f}%")
        print(f"    Lung RoI (Pneumonia cases):   {mean_lung_pneu*100:.2f}%")
        print(f"    Lung RoI (Normal cases):      {mean_lung_norm*100:.2f}%")
        
        results[model_name] = {
            "display_name": display_name,
            "samples_evaluated": len(lung_ratios),
            "energy_in_lung_roi_mean": mean_lung,
            "energy_in_lung_roi_std": std_lung,
            "peripheral_shortcut_energy_mean": mean_periph,
            "peripheral_shortcut_energy_std": std_periph,
            "lung_energy_pneumonia_mean": mean_lung_pneu,
            "lung_energy_normal_mean": mean_lung_norm
        }
        
    out_path = ROOT / "results" / "gradcam_energy_metrics.json"
    with open(out_path, "w", encoding="utf-8") as f:
        json.dump(results, f, indent=2, ensure_ascii=False)
    print(f"\nSaved Grad-CAM Energy Metrics to: {out_path}")
    print("=" * 75)
    return results

if __name__ == "__main__":
    run_step4()
