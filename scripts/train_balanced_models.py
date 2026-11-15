# -*- coding: utf-8 -*-
"""
train_balanced_models.py

1. Balances the training set by oversampling the NORMAL class (1,078 -> 2,737 images)
   using medical-safe dynamic data augmentations (Affine rotation +-7 deg, scale 0.95-1.05,
   translation +-3%, color jitter +-8%, strictly no horizontal flips).
   Resulting in an exact 1:1 ratio (2,737 Normal : 2,737 Pneumonia).
2. Keeps Validation (938) and Test (618) splits completely identical to the
   certified patient-disjoint benchmark for direct, objective comparison.
3. Trains ResNet-50 and ResNet-50-LungRoI-Attention with unweighted BCEWithLogitsLoss (pos_weight=1.0).
4. Evaluates on the 618-image Test set and outputs a comparative analysis.
"""

import os
import sys
import io
import time
import json
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import pandas as pd
import numpy as np
import torch
import torch.nn as nn
from torch.utils.data import DataLoader

from src.models import build_model, classifier
from src.dataset import XrayDataset, validate_manifest
from src.metrics import binary_metrics
from src.utils import device, seed_everything, save_json

def create_balanced_manifest():
    manifest_path = ROOT / "results" / "manifest.csv"
    manifest = pd.read_csv(manifest_path)
    
    train_orig = manifest[manifest.split == "train"].copy()
    val_orig = manifest[manifest.split == "val"].copy()
    test_orig = manifest[manifest.split == "test"].copy()
    
    normal_train = train_orig[train_orig.label == 0]
    pneumonia_train = train_orig[train_orig.label == 1]
    
    n_normal = len(normal_train)
    n_pneumonia = len(pneumonia_train)
    n_oversample = n_pneumonia - n_normal
    
    print(f"Original Train: Normal={n_normal}, Pneumonia={n_pneumonia} (Ratio 1:{n_pneumonia/n_normal:.2f})")
    print(f"Oversampling Normal class by adding {n_oversample} samples with dynamic medical-safe augmentation...")
    
    # Random sampling with replacement from the existing Normal training set
    oversampled_normal = normal_train.sample(n=n_oversample, replace=True, random_state=42)
    
    train_balanced = pd.concat([normal_train, oversampled_normal, pneumonia_train], ignore_index=True)
    # Shuffle the balanced training set
    train_balanced = train_balanced.sample(frac=1.0, random_state=42).reset_index(drop=True)
    
    print(f"Balanced Train: Normal={(train_balanced.label == 0).sum()}, Pneumonia={(train_balanced.label == 1).sum()} (Total={len(train_balanced)})")
    
    # Verify strict patient isolation is preserved
    train_patients = set(train_balanced.patient)
    val_patients = set(val_orig.patient)
    test_patients = set(test_orig.patient)
    assert train_patients.isdisjoint(test_patients), "Patient leakage between balanced train and test!"
    assert val_patients.isdisjoint(test_patients), "Patient leakage between val and test!"
    assert train_patients.isdisjoint(val_patients), "Patient leakage between balanced train and val!"
    print("Patient-Disjoint check: PASSED (zero patient leakage between splits)")
    
    balanced_manifest = pd.concat([train_balanced, val_orig, test_orig], ignore_index=True)
    out_manifest = ROOT / "results" / "manifest_balanced.csv"
    balanced_manifest.to_csv(out_manifest, index=False)
    print(f"Saved balanced manifest to: {out_manifest}")
    return balanced_manifest


def train_single_model(model_name, manifest_df, target_device, head_epochs=2, finetune_epochs=6, batch_size=16):
    print("\n" + "="*70)
    print(f"TRAINING BALANCED MODEL: {model_name.upper()}")
    print("="*70)
    
    seed_everything(42)
    
    train_df = manifest_df[manifest_df.split == "train"].copy()
    val_df = manifest_df[manifest_df.split == "val"].copy()
    test_df = manifest_df[manifest_df.split == "test"].copy()
    
    train_loader = DataLoader(XrayDataset(train_df, train=True), batch_size=batch_size, shuffle=True, pin_memory=True, num_workers=0)
    val_loader = DataLoader(XrayDataset(val_df, train=False), batch_size=batch_size, shuffle=False, pin_memory=True, num_workers=0)
    test_loader = DataLoader(XrayDataset(test_df, train=False), batch_size=batch_size, shuffle=False, pin_memory=True, num_workers=0)
    
    # 1:1 balanced loss: pos_weight = 1.0
    criterion = nn.BCEWithLogitsLoss()
    
    model = build_model(model_name, pretrained=True).to(target_device)
    scaler = torch.amp.GradScaler("cuda", enabled=target_device.type == "cuda")
    
    checkpoint_path = ROOT / "checkpoints" / f"best_{model_name}_balanced.pth"
    history_path = ROOT / "logs" / f"history_{model_name}_balanced.csv"
    
    best_val_loss = float("inf")
    history = []
    epoch_counter = 0
    start_time = time.time()
    
    stages = [(1, head_epochs, 1e-3), (2, finetune_epochs, 1e-4)]
    
    for stage, epochs, lr in stages:
        print(f"\n--- Stage {stage}: {'Head Warmup' if stage == 1 else 'Full Fine-Tuning'} (LR = {lr}) ---")
        
        if model_name in ("resnet50_attention", "resnet50_roi_attention"):
            for p in model.backbone.parameters():
                p.requires_grad = (stage == 2)
            for p in model.attention_gate.parameters():
                p.requires_grad = True
            for p in model.classifier.parameters():
                p.requires_grad = True
        else:
            for p in model.parameters():
                p.requires_grad = (stage == 2)
            for p in classifier(model).parameters():
                p.requires_grad = True
                
        optimizer = torch.optim.AdamW(
            filter(lambda p: p.requires_grad, model.parameters()),
            lr=lr,
            weight_decay=1e-4
        )
        scheduler = torch.optim.lr_scheduler.ReduceLROnPlateau(optimizer, mode="min", factor=0.5, patience=2)
        
        for ep in range(1, epochs + 1):
            epoch_counter += 1
            model.train()
            if stage == 1:
                if hasattr(model, "backbone"):
                    model.backbone.eval()
                else:
                    model.eval()
                    classifier(model).train()
                    
            train_loss = 0.0
            train_labels, train_preds = [], []
            
            for imgs, targets, _ in train_loader:
                imgs = imgs.to(target_device)
                targets = targets.to(target_device, dtype=torch.float32)
                
                optimizer.zero_grad(set_to_none=True)
                with torch.autocast(device_type=target_device.type, enabled=scaler.is_enabled()):
                    logits = model(imgs).flatten()
                    loss = criterion(logits, targets)
                    
                scaler.scale(loss).backward()
                scaler.step(optimizer)
                scaler.update()
                
                train_loss += loss.item() * len(targets)
                train_labels.extend(targets.detach().cpu().tolist())
                train_preds.extend(logits.detach().sigmoid().cpu().tolist())
                
            train_m = binary_metrics(train_labels, train_preds, threshold=0.5)
            train_loss /= len(train_df)
            
            # Validation
            model.eval()
            val_loss = 0.0
            val_labels, val_preds = [], []
            with torch.no_grad():
                for imgs, targets, _ in val_loader:
                    imgs = imgs.to(target_device)
                    targets = targets.to(target_device, dtype=torch.float32)
                    with torch.autocast(device_type=target_device.type, enabled=scaler.is_enabled()):
                        logits = model(imgs).flatten()
                        loss = criterion(logits, targets)
                    val_loss += loss.item() * len(targets)
                    val_labels.extend(targets.cpu().tolist())
                    val_preds.extend(logits.sigmoid().cpu().tolist())
                    
            val_m = binary_metrics(val_labels, val_preds, threshold=0.5)
            val_loss /= len(val_df)
            
            scheduler.step(val_loss)
            
            print(f"  Epoch {epoch_counter:2d} (S{stage} E{ep}) | Train Loss: {train_loss:.4f} Acc: {train_m['accuracy']*100:.2f}% | "
                  f"Val Loss: {val_loss:.4f} Acc: {val_m['accuracy']*100:.2f}% Sens: {val_m['recall']*100:.2f}% Spec: {val_m['specificity']*100:.2f}%")
            
            history.append({
                "epoch": epoch_counter,
                "stage": stage,
                "train_loss": train_loss,
                "train_acc": train_m["accuracy"],
                "val_loss": val_loss,
                "val_acc": val_m["accuracy"],
                "val_recall": val_m["recall"],
                "val_specificity": val_m["specificity"],
                "val_f1": val_m["f1"]
            })
            
            # Save best checkpoint
            if val_loss < best_val_loss or (stage == 1 and ep == 1):
                best_val_loss = val_loss
                torch.save({
                    "model": model_name,
                    "trained": True,
                    "balanced": True,
                    "state_dict": model.state_dict(),
                    "epoch": epoch_counter,
                    "val_loss": val_loss,
                    "val_metrics": val_m
                }, checkpoint_path)
                
    pd.DataFrame(history).to_csv(history_path, index=False)
    total_time = time.time() - start_time
    print(f"Training completed in {total_time/60:.2f} minutes. Best Val Loss: {best_val_loss:.4f}")
    
    # Evaluate best checkpoint on Test set
    print(f"\n--- Evaluating Best {model_name.upper()} Balanced on Clean Test Set (618 images) ---")
    best_cp = torch.load(checkpoint_path, map_location=target_device)
    model.load_state_dict(best_cp["state_dict"])
    model.eval()
    
    test_labels, test_preds = [], []
    with torch.no_grad():
        for imgs, targets, _ in test_loader:
            imgs = imgs.to(target_device)
            with torch.autocast(device_type=target_device.type, enabled=scaler.is_enabled()):
                logits = model(imgs).flatten()
            test_preds.extend(logits.sigmoid().cpu().tolist())
            test_labels.extend(targets.tolist())
            
    test_metrics = binary_metrics(test_labels, test_preds, threshold=0.5)
    print(f"Results on Test Set:")
    print(f"  Accuracy    : {test_metrics['accuracy']*100:.2f}%")
    print(f"  Sensitivity : {test_metrics['recall']*100:.2f}% (TP={test_metrics['tp']}, FN={test_metrics['fn']})")
    print(f"  Specificity : {test_metrics['specificity']*100:.2f}% (TN={test_metrics['tn']}, FP={test_metrics['fp']})")
    print(f"  F1-Score    : {test_metrics['f1']*100:.2f}%")
    print(f"  ROC-AUC     : {test_metrics['auc']*100:.2f}%")
    
    return test_metrics


def main():
    target_device = device()
    print(f"Using Compute Device: {target_device} ({torch.cuda.get_device_name(0) if target_device.type == 'cuda' else 'CPU'})")
    
    # 1. Create balanced manifest
    balanced_manifest = create_balanced_manifest()
    
    # 2. Train ResNet50 on balanced dataset
    resnet_balanced_metrics = train_single_model("resnet50", balanced_manifest, target_device, head_epochs=2, finetune_epochs=6, batch_size=16)
    
    # 3. Train ResNet50_Attention on balanced dataset
    attn_balanced_metrics = train_single_model("resnet50_attention", balanced_manifest, target_device, head_epochs=2, finetune_epochs=6, batch_size=16)
    
    # 4. Compare with original imbalanced models
    orig_resnet = json.loads((ROOT / "results" / "resnet50_metrics.json").read_text())
    orig_attn = json.loads((ROOT / "results" / "resnet50_attention_metrics.json").read_text())
    
    comparison_data = [
        {
            "Protocol": "ResNet-50 (Original Imbalanced 1078:2737)",
            "Accuracy": orig_resnet["accuracy"],
            "Sensitivity": orig_resnet["recall"],
            "Specificity": orig_resnet["specificity"],
            "F1": orig_resnet["f1"],
            "AUC": orig_resnet["auc"],
            "TP": orig_resnet["tp"],
            "TN": orig_resnet["tn"],
            "FP": orig_resnet["fp"],
            "FN": orig_resnet["fn"],
        },
        {
            "Protocol": "ResNet-50 (Balanced 1:1 Oversampled 2737:2737)",
            "Accuracy": resnet_balanced_metrics["accuracy"],
            "Sensitivity": resnet_balanced_metrics["recall"],
            "Specificity": resnet_balanced_metrics["specificity"],
            "F1": resnet_balanced_metrics["f1"],
            "AUC": resnet_balanced_metrics["auc"],
            "TP": resnet_balanced_metrics["tp"],
            "TN": resnet_balanced_metrics["tn"],
            "FP": resnet_balanced_metrics["fp"],
            "FN": resnet_balanced_metrics["fn"],
        },
        {
            "Protocol": "ResNet50-Attn (Original Imbalanced 1078:2737)",
            "Accuracy": orig_attn["accuracy"],
            "Sensitivity": orig_attn["recall"],
            "Specificity": orig_attn["specificity"],
            "F1": orig_attn["f1"],
            "AUC": orig_attn["auc"],
            "TP": orig_attn["tp"],
            "TN": orig_attn["tn"],
            "FP": orig_attn["fp"],
            "FN": orig_attn["fn"],
        },
        {
            "Protocol": "ResNet50-Attn (Balanced 1:1 Oversampled 2737:2737)",
            "Accuracy": attn_balanced_metrics["accuracy"],
            "Sensitivity": attn_balanced_metrics["recall"],
            "Specificity": attn_balanced_metrics["specificity"],
            "F1": attn_balanced_metrics["f1"],
            "AUC": attn_balanced_metrics["auc"],
            "TP": attn_balanced_metrics["tp"],
            "TN": attn_balanced_metrics["tn"],
            "FP": attn_balanced_metrics["fp"],
            "FN": attn_balanced_metrics["fn"],
        }
    ]
    
    df_comp = pd.DataFrame(comparison_data)
    df_comp.to_csv(ROOT / "results" / "balanced_model_comparison.csv", index=False)
    
    summary_json = {
        "resnet50_balanced": resnet_balanced_metrics,
        "resnet50_attention_balanced": attn_balanced_metrics,
        "comparison_table": comparison_data
    }
    save_json(ROOT / "results" / "balanced_experiment_results.json", summary_json)
    
    print("\n" + "="*80)
    print("FINAL BALANCED EXPERIMENT COMPARISON SUMMARY")
    print("="*80)
    for row in comparison_data:
        print(f"{row['Protocol']:<48} | Acc: {row['Accuracy']*100:.2f}% | Sens: {row['Sensitivity']*100:.2f}% | Spec: {row['Specificity']*100:.2f}% | FP: {row['FP']:2d} | FN: {row['FN']:2d}")
    print("="*80)

if __name__ == "__main__":
    main()
