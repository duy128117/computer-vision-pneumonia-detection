# -*- coding: utf-8 -*-
"""
STEP 5: Create Benchmark Release Data Artifacts
Exports:
1. benchmark_release/patient_disjoint_manifest.csv (certified patient-disjoint split)
2. benchmark_release/duplicate_and_corrupt_sha256.csv (SHA-256 duplicate list & exclusion audit)
3. benchmark_release/evaluation_protocol.json (standardized metrics and benchmark definitions)
4. benchmark_release/README.md (comprehensive release documentation for the research community)
"""

import os
import sys
import io
import json
import shutil
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding='utf-8', errors='replace')

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

import pandas as pd

def run_step5():
    print("=" * 75)
    print("STEP 5: GENERATING BENCHMARK RELEASE PACKAGE")
    print("=" * 75)
    
    release_dir = ROOT / "benchmark_release"
    release_dir.mkdir(parents=True, exist_ok=True)
    
    # 1. Patient-disjoint manifest
    src_manifest = ROOT / "results" / "manifest.csv"
    if src_manifest.exists():
        df_manifest = pd.read_csv(src_manifest)
        # Convert absolute paths to relative paths for portability
        rel_paths = []
        for p in df_manifest["path"]:
            p_obj = Path(p)
            try:
                rel = p_obj.relative_to(ROOT)
                rel_paths.append(str(rel).replace("\\", "/"))
            except ValueError:
                # If not relative to ROOT, find from 'data/'
                idx = str(p_obj).find("data")
                if idx != -1:
                    rel_paths.append(str(p_obj)[idx:].replace("\\", "/"))
                else:
                    rel_paths.append(p_obj.name)
        df_manifest_clean = df_manifest.copy()
        df_manifest_clean["path"] = rel_paths
        
        dest_manifest = release_dir / "patient_disjoint_manifest.csv"
        df_manifest_clean.to_csv(dest_manifest, index=False)
        print(f"[OK] Exported {len(df_manifest_clean)} rows to {dest_manifest}")
        
    # 2. Duplicate & Corrupt SHA-256 List
    src_excluded = ROOT / "results" / "excluded_images.csv"
    if src_excluded.exists():
        df_exc = pd.read_csv(src_excluded)
        rel_exc_paths = []
        for p in df_exc["path"]:
            p_obj = Path(p)
            idx = str(p_obj).find("data")
            if idx != -1:
                rel_exc_paths.append(str(p_obj)[idx:].replace("\\", "/"))
            else:
                rel_exc_paths.append(p_obj.name)
        df_exc_clean = df_exc.copy()
        df_exc_clean["path"] = rel_exc_paths
        
        dest_exc = release_dir / "duplicate_and_corrupt_sha256.csv"
        df_exc_clean.to_csv(dest_exc, index=False)
        print(f"[OK] Exported {len(df_exc_clean)} excluded instances to {dest_exc}")
        
    # 3. Benchmark Protocol JSON
    protocol = {
        "benchmark_name": "Kermany-Clean-Disjoint Pediatric Pneumonia Benchmark (KCD-Pneumonia)",
        "version": "2.0",
        "description": "Certified patient-disjoint, deduplicated evaluation benchmark derived from Kermany chest X-ray dataset.",
        "splits": {
            "train": {"samples": 3815, "patients": "disjoint_pool_train", "label_ratio_normal_pneumonia": "942 / 2873"},
            "val": {"samples": 938, "patients": "disjoint_pool_val", "label_ratio_normal_pneumonia": "241 / 697"},
            "test": {"samples": 618, "patients": "disjoint_pool_test_strictly_held_out", "label_ratio_normal_pneumonia": "231 / 387"}
        },
        "evaluation_standard": {
            "decision_threshold": 0.5,
            "metrics": ["Accuracy", "Sensitivity (Recall)", "Specificity", "F1-Score", "ROC-AUC", "Exact McNemar p-value"],
            "bootstrap_ci": "95% BCa / Percentile with N=1000 resamples",
            "cross_validation_clarification": "5-Fold CV is reported solely on development cohort (4,753 images); test set (618 images) is strictly frozen."
        }
    }
    with open(release_dir / "evaluation_protocol.json", "w", encoding="utf-8") as f:
        json.dump(protocol, f, indent=2)
    print(f"[OK] Exported protocol to {release_dir / 'evaluation_protocol.json'}")
    
    # 4. README.md
    readme_content = """# Kermany-Clean-Disjoint (KCD) Pediatric Pneumonia Benchmark

## 1. Overview & Motivation
The popular Kermany chest X-ray dataset (Kermany et al., 2018) is widely utilized in medical imaging literature. However, rigorous audit reveals critical methodological vulnerabilities:
1. **Identical Image Duplication**: Exact bit-level duplicate images exist across splits.
2. **Patient Overlap (Data Leakage)**: Multiple X-ray acquisitions of the same pediatric patient are dispersed across training and test splits, allowing deep networks to memorize patient-specific anatomical shortcuts rather than pathology.
3. **Severe Validation Starvation**: The naive split contains only 16 images in `val/`, creating noisy early stopping.

This package provides a standardized, verified, **patient-disjoint benchmark split** along with full exclusion manifests to enable fair, reproducible, and leak-free evaluation.

## 2. Dataset Structure
- Total verified unique images: **5,371**
- **Train Split**: 3,815 images (Patient-disjoint)
- **Validation Split**: 938 images (Patient-disjoint)
- **Test Split**: 618 images (Held-out prospective cohort, strictly disjoint patient IDs)

## 3. Contents of This Package
1. `patient_disjoint_manifest.csv`: Image path, binary label (0=Normal, 1=Pneumonia), anonymized patient identifier, and assigned split.
2. `duplicate_and_corrupt_sha256.csv`: Audit log of all identical duplicates, SHA-256 hashes, and corrupt files excluded from the certified benchmark.
3. `evaluation_protocol.json`: Formal evaluation protocol, metric definitions, and confidence interval guidelines.

## 4. Usage in Python
```python
import pandas as pd

# Load certified manifest
manifest = pd.read_csv("patient_disjoint_manifest.csv")
train_df = manifest[manifest.split == "train"]
val_df = manifest[manifest.split == "val"]
test_df = manifest[manifest.split == "test"]

# Verify patient disjointness
assert len(set(train_df.patient).intersection(set(test_df.patient))) == 0
assert len(set(val_df.patient).intersection(set(test_df.patient))) == 0
print(f"Verified patient-disjoint test set: {len(test_df)} images.")
```

## 5. Citation & License
Please cite this benchmark when reporting experimental results on leak-free pediatric pneumonia detection. Released under Creative Commons Attribution 4.0 International (CC BY 4.0).
"""
    with open(release_dir / "README.md", "w", encoding="utf-8") as f:
        f.write(readme_content)
    print(f"[OK] Created documentation at {release_dir / 'README.md'}")
    print("=" * 75)

if __name__ == "__main__":
    run_step5()
