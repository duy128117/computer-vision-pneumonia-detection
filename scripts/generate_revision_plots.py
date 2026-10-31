"""
Generate publication-quality charts for:
1. Leakage Inflation Gap Comparison
2. 5-Fold StratifiedGroupKFold Metrics with Error Bars
"""

from pathlib import Path
import json
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
RESULTS = ROOT / "results"
PLOTS = RESULTS / "plots"
PLOTS.mkdir(parents=True, exist_ok=True)

# Set global aesthetic style
plt.style.use('seaborn-v0_8-whitegrid' if 'seaborn-v0_8-whitegrid' in plt.style.available else 'default')
plt.rcParams['font.family'] = 'DejaVu Sans'
plt.rcParams['font.size'] = 11

def plot_leakage_gap():
    data_path = RESULTS / "leakage_inflation_gap.json"
    with open(data_path, "r", encoding="utf-8") as f:
        data = json.load(f)
        
    labels = ['Accuracy (%)', 'Specificity (%)', 'AUC-ROC (%)']
    naive_vals = [
        data['literature_naive_split']['accuracy'] * 100,
        data['literature_naive_split']['specificity'] * 100,
        data['literature_naive_split']['auc'] * 100
    ]
    clean_vals = [
        data['clean_patient_disjoint_split']['accuracy'] * 100,
        data['clean_patient_disjoint_split']['specificity'] * 100,
        data['clean_patient_disjoint_split']['auc'] * 100
    ]
    
    x = np.arange(len(labels))
    width = 0.35
    
    fig, ax = plt.subplots(figsize=(8, 5), dpi=300)
    rects1 = ax.bar(x - width/2, naive_vals, width, label='Bộ chia gốc Kermany (Có rò rỉ bệnh nhân & Trùng ảnh)', color='#E74C3C', alpha=0.9)
    rects2 = ax.bar(x + width/2, clean_vals, width, label='Bộ chia sạch Patient-Disjoint (Đề tài đề xuất)', color='#2980B9', alpha=0.9)
    
    ax.set_ylabel('Giá trị chỉ số (%)', fontweight='bold')
    ax.set_title('Định lượng Mức độ Thổi phồng Độ chính xác do Rò rỉ Dữ liệu\n(Leakage Inflation Gap)', fontsize=13, fontweight='bold', pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontweight='bold')
    ax.set_ylim(50, 105)
    ax.legend(loc='lower left', frameon=True)
    
    # Annotate gap
    for i in range(len(labels)):
        gap = naive_vals[i] - clean_vals[i]
        ax.annotate(f'Chênh lệch:\n+{gap:.2f}% (Ảo)',
                    xy=(x[i], max(naive_vals[i], clean_vals[i]) + 1.5),
                    ha='center', va='bottom', fontsize=10, fontweight='bold', color='#C0392B',
                    bbox=dict(boxstyle="round,pad=0.3", fc="#FDEDEC", ec="#E74C3C", lw=1))
        
    fig.tight_layout()
    out_path = PLOTS / "leakage_inflation_gap.png"
    plt.savefig(out_path)
    plt.close()
    print(f"Saved {out_path}")

def plot_kfold_summary():
    data_path = RESULTS / "kfold_metrics.csv"
    df = pd.read_csv(data_path)
    
    metrics = ['accuracy', 'recall', 'specificity', 'f1', 'auc']
    display_names = ['Accuracy', 'Sensitivity', 'Specificity', 'F1-Score', 'AUC-ROC']
    
    means = [df[m].mean() * 100 for m in metrics]
    stds = [df[m].std() * 100 for m in metrics]
    
    fig, ax = plt.subplots(figsize=(8, 4.8), dpi=300)
    x = np.arange(len(metrics))
    
    bars = ax.bar(x, means, yerr=stds, capsize=6, color='#27AE60', alpha=0.85, edgecolor='#1E8449', linewidth=1.5)
    
    ax.set_ylabel('Điểm số (%)', fontweight='bold')
    ax.set_title('Độ tin cậy Thống kê qua 5-Fold StratifiedGroupKFold\n(Bảo toàn độc lập bệnh nhân, Mean ± Std)', fontsize=13, fontweight='bold', pad=15)
    ax.set_xticks(x)
    ax.set_xticklabels(display_names, fontweight='bold')
    ax.set_ylim(90, 102)
    
    for bar, mean, std in zip(bars, means, stds):
        yval = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2.0, yval - 4.5, f'{mean:.2f}%\n± {std:.2f}%', ha='center', va='bottom', color='white', fontweight='bold', fontsize=10)
        
    fig.tight_layout()
    out_path = PLOTS / "kfold_summary_barchart.png"
    plt.savefig(out_path)
    plt.close()
    print(f"Saved {out_path}")

if __name__ == '__main__':
    plot_leakage_gap()
    plot_kfold_summary()
