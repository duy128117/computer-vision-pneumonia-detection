# -*- coding: utf-8 -*-
"""
STEP 1: Controlled Experiment for Leakage Inflation Gap.
Trains ResNet50 on Kermany's naive original split (with duplicate images & patient overlap)
using the exact same hyperparameters and training protocol.
Then evaluates on:
1. Kermany Naive Test (624 images)
2. Clean Patient-Disjoint Test (618 images)
Quantifies the exact controlled Leakage Inflation Gap.
"""

import os
import sys
import io

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

import json
import time
import hashlib
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import pandas as pd
import torch
import torch.nn as nn
from torch.utils.data import DataLoader

from src.models import build_model, classifier
from src.dataset import XrayDataset, patient_id
from src.metrics import binary_metrics
from src.utils import device, seed_everything, save_json

def build_naive_kermany_manifest():
    root = ROOT / "data" / "chest_xray"
    rows = []
    for split in ["train", "val", "test"]:
        for label, cls in enumerate(["NORMAL", "PNEUMONIA"]):
            folder = root / split / cls
            for p in sorted(folder.iterdir()):
                if p.suffix.lower() in [".jpeg", ".jpg", ".png"]:
                    try:
                        pid = patient_id(p)
                    except Exception:
                        pid = p.stem.split("_")[0].lower()
                    rows.append({
                        "path": str(p),
                        "label": label,
                        "patient": pid,
                        "split": split
                    })
    df = pd.DataFrame(rows)
    out_path = ROOT / "results" / "manifest_kermany_naive.csv"
    df.to_csv(out_path, index=False)
    print(f"Created Naive Kermany Manifest: {len(df)} images at {out_path}")
    print(f"  Train: {(df.split == 'train').sum()} | Val: {(df.split == 'val').sum()} | Test: {(df.split == 'test').sum()}")
    return df

def run_controlled_experiment():
    print("="*75)
    print("STEP 1: RUNNING CONTROLLED EXPERIMENT FOR LEAKAGE INFLATION GAP")
    print("="*75)
    
    target_device = device()
    seed_everything(42)
    
    df_naive = build_naive_kermany_manifest()
    
    # Clean patient-disjoint test set (from certified manifest.csv)
    clean_manifest_path = ROOT / "results" / "manifest.csv"
    df_clean = pd.read_csv(clean_manifest_path)
    df_clean_test = df_clean[df_clean.split == "test"].copy()
    
    # 1. Setup DataLoaders for Naive Kermany Training
    train_df = df_naive[df_naive.split == "train"]
    val_df = df_naive[df_naive.split == "val"]
    naive_test_df = df_naive[df_naive.split == "test"]
    
    counts = train_df.label.value_counts()
    weight = float(counts[0] / counts[1])  # 1341 / 3875 ~ 0.346
    print(f"Naive Train: {counts.to_dict()} -> pos_weight = {weight:.4f}")
    
    criterion = nn.BCEWithLogitsLoss(pos_weight=torch.tensor(weight, device=target_device))
    
    train_loader = DataLoader(XrayDataset(train_df, train=True), batch_size=8, shuffle=True, pin_memory=True)
    val_loader = DataLoader(XrayDataset(val_df, train=False), batch_size=8, shuffle=False, pin_memory=True)
    naive_test_loader = DataLoader(XrayDataset(naive_test_df, train=False), batch_size=8, shuffle=False, pin_memory=True)
    clean_test_loader = DataLoader(XrayDataset(df_clean_test, train=False), batch_size=8, shuffle=False, pin_memory=True)
    
    # 2. Build and train ResNet50 on Naive Split
    model = build_model("resnet50", pretrained=True).to(target_device)
    scaler = torch.amp.GradScaler("cuda", enabled=target_device.type == "cuda")
    
    checkpoint_path = ROOT / "checkpoints" / "best_resnet50_naive_split.pth"
    history_path = ROOT / "logs" / "history_resnet50_naive_split.csv"
    
    head_epochs = 3
    finetune_epochs = 8
    best_val_loss = float("inf")
    history = []
    epoch_counter = 0
    start_time = time.time()
    
    print("\n--- Training ResNet50 on Naive Kermany Split ---")
    for stage, epochs, lr in [(1, head_epochs, 1e-3), (2, finetune_epochs, 1e-4)]:
        print(f"Stage {stage} (LR={lr}):")
        for param in model.parameters():
            param.requires_grad = (stage == 2)
        for param in classifier(model).parameters():
            param.requires_grad = True
            
        optimizer = torch.optim.AdamW(filter(lambda p: p.requires_grad, model.parameters()), lr=lr, weight_decay=1e-4)
        scheduler = torch.optim.lr_scheduler.ReduceLROnPlateau(optimizer, patience=2, factor=0.5)
        
        for ep in range(1, epochs + 1):
            epoch_counter += 1
            model.train()
            if stage == 1:
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
                train_labels.extend(targets.cpu().tolist())
                train_preds.extend(logits.detach().sigmoid().cpu().tolist())
                
            train_m = binary_metrics(train_labels, train_preds)
            train_loss /= len(train_df)
            
            # Val (16 images)
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
            val_m = binary_metrics(val_labels, val_preds)
            val_loss /= len(val_df)
            
            scheduler.step(val_loss)
            
            print(f"  Epoch {epoch_counter:2d} | Train Loss: {train_loss:.4f} Acc: {train_m['accuracy']*100:.2f}% | Val Loss: {val_loss:.4f} Acc: {val_m['accuracy']*100:.2f}%")
            
            history.append({
                "epoch": epoch_counter,
                "stage": stage,
                "train_loss": train_loss,
                "train_acc": train_m["accuracy"],
                "val_loss": val_loss,
                "val_acc": val_m["accuracy"]
            })
            
            if val_loss < best_val_loss or ep == 1:
                best_val_loss = val_loss
                torch.save({
                    "model": "resnet50_naive_split",
                    "trained": True,
                    "state_dict": model.state_dict(),
                    "epoch": epoch_counter,
                    "val_loss": val_loss
                }, checkpoint_path)
                
    pd.DataFrame(history).to_csv(history_path, index=False)
    
    # 3. Load best checkpoint and evaluate on both test sets
    print("\n--- Evaluating ResNet50_Naive on Both Test Sets ---")
    best_cp = torch.load(checkpoint_path, map_location=target_device)
    model.load_state_dict(best_cp["state_dict"])
    model.eval()
    
    def eval_loader(loader):
        probs, labels = [], []
        with torch.no_grad():
            for imgs, lbls, _ in loader:
                imgs = imgs.to(target_device)
                out = model(imgs).flatten().sigmoid().cpu().tolist()
                probs.extend(out)
                labels.extend(lbls.tolist())
        return binary_metrics(labels, probs, threshold=0.5)
        
    naive_test_results = eval_loader(naive_test_loader)
    clean_test_results = eval_loader(clean_test_loader)
    
    # 4. Load baseline ResNet50 trained on Patient-Disjoint split for direct comparison
    resnet_clean_path = ROOT / "results" / "resnet50_metrics.json"
    resnet_clean_results = json.loads(resnet_clean_path.read_text())
    
    controlled_report = {
        "experiment": "Controlled Leakage Inflation Gap Evaluation",
        "description": "Same architecture (ResNet50), same training protocol, evaluated under Naive vs Patient-Disjoint conditions",
        "kermany_naive_model_on_naive_test": {
            "test_images": len(naive_test_df),
            "accuracy": naive_test_results["accuracy"],
            "sensitivity": naive_test_results["recall"],
            "specificity": naive_test_results["specificity"],
            "f1": naive_test_results["f1"],
            "auc": naive_test_results["auc"],
            "fn": naive_test_results["fn"],
            "fp": naive_test_results["fp"]
        },
        "kermany_naive_model_on_clean_disjoint_test": {
            "test_images": len(df_clean_test),
            "accuracy": clean_test_results["accuracy"],
            "sensitivity": clean_test_results["recall"],
            "specificity": clean_test_results["specificity"],
            "f1": clean_test_results["f1"],
            "auc": clean_test_results["auc"],
            "fn": clean_test_results["fn"],
            "fp": clean_test_results["fp"]
        },
        "patient_disjoint_trained_model_on_clean_test": {
            "test_images": len(df_clean_test),
            "accuracy": resnet_clean_results["accuracy"],
            "sensitivity": resnet_clean_results["recall"],
            "specificity": resnet_clean_results["specificity"],
            "f1": resnet_clean_results["f1"],
            "auc": resnet_clean_results["auc"],
            "fn": resnet_clean_results["fn"],
            "fp": resnet_clean_results["fp"]
        },
        "controlled_leakage_inflation_gap": {
            "accuracy_gap": float(naive_test_results["accuracy"] - clean_test_results["accuracy"]),
            "specificity_gap": float(naive_test_results["specificity"] - clean_test_results["specificity"]),
            "auc_gap": float(naive_test_results["auc"] - clean_test_results["auc"])
        }
    }
    
    save_json(ROOT / "results" / "controlled_leakage_experiment.json", controlled_report)
    
    print("\n" + "="*75)
    print("CONTROLLED LEAKAGE INFLATION GAP RESULTS:")
    print(f"1. Naive Model tested on Naive Kermany Test (624 imgs, with leakage):")
    print(f"   Accuracy: {naive_test_results['accuracy']*100:.2f}% | Specificity: {naive_test_results['specificity']*100:.2f}% | Sensitivity: {naive_test_results['recall']*100:.2f}% | AUC: {naive_test_results['auc']*100:.2f}%")
    print(f"2. Naive Model tested on Clean Disjoint Test (618 imgs, unseen patients):")
    print(f"   Accuracy: {clean_test_results['accuracy']*100:.2f}% | Specificity: {clean_test_results['specificity']*100:.2f}% | Sensitivity: {clean_test_results['recall']*100:.2f}% | AUC: {clean_test_results['auc']*100:.2f}%")
    print(f"3. ResNet50 Trained on Clean Disjoint Split tested on Clean Disjoint Test:")
    print(f"   Accuracy: {resnet_clean_results['accuracy']*100:.2f}% | Specificity: {resnet_clean_results['specificity']*100:.2f}% | Sensitivity: {resnet_clean_results['recall']*100:.2f}% | AUC: {resnet_clean_results['auc']*100:.2f}%")
    print(f"--> CONTROLLED INFLATION GAP ON TEST:")
    print(f"   Accuracy Gap:    +{controlled_report['controlled_leakage_inflation_gap']['accuracy_gap']*100:.2f}% (False Boost)")
    print(f"   Specificity Gap: +{controlled_report['controlled_leakage_inflation_gap']['specificity_gap']*100:.2f}% (False Boost)")
    print(f"   AUC Gap:         +{controlled_report['controlled_leakage_inflation_gap']['auc_gap']*100:.2f}%")
    print("="*75)
    print(f"Completed in {time.time() - start_time:.1f}s")
    return controlled_report

if __name__ == "__main__":
    run_controlled_experiment()
