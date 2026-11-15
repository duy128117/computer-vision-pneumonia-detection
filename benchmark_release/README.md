# Kermany-Clean-Disjoint (KCD) Pediatric Pneumonia Benchmark

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
