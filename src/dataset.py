import argparse
import hashlib
import re
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
from PIL import Image
from sklearn.model_selection import StratifiedGroupKFold
from torch.utils.data import Dataset

from .transforms import image_transform
from .utils import save_json


def patient_id(path):
    name = Path(path).stem
    match = re.match(r"(person\d+)_", name, re.I)
    if match:
        return match.group(1).lower()
    match = re.match(r"(IM-\d+|NORMAL2-IM-\d+)-", name, re.I)
    if match:
        return match.group(1).lower()
    raise ValueError(
        f"Không suy ra patient ID: {name}. Cần bổ sung quy tắc tên hoặc metadata trước khi chia."
    )


def prepare(root, output="results", seed=42):
    root, output = Path(root).resolve(), Path(output)
    rows = []
    for split in ("train", "val", "test"):
        for label, cls in enumerate(("NORMAL", "PNEUMONIA")):
            folder = root / split / cls
            if not folder.is_dir():
                raise FileNotFoundError(folder)
            for path in sorted(folder.iterdir()):
                if path.suffix.lower() not in (".jpeg", ".jpg", ".png"):
                    continue
                with Image.open(path) as image:
                    gray = image.convert("L")
                    digest = hashlib.sha256(str(gray.size).encode() + gray.tobytes()).hexdigest()
                rows.append(
                    dict(
                        path=str(path),
                        label=label,
                        patient=patient_id(path),
                        original_split=split,
                        sha256=digest,
                    )
                )
                if len(rows) % 1000 == 0:
                    print(f"Validated {len(rows)} source images", flush=True)
    frame = pd.DataFrame(rows)
    if frame.empty:
        raise ValueError("Không có ảnh.")
    if frame.groupby("sha256").label.nunique().max() > 1:
        raise ValueError("Ảnh trùng có nhãn mâu thuẫn.")
    # Keep official test intact. Exclude all matching patients and pixels from development.
    test = frame[frame.original_split == "test"].drop_duplicates("sha256").copy()
    dev = frame[frame.original_split != "test"].copy()
    dev = dev[~dev.patient.isin(test.patient) & ~dev.sha256.isin(test.sha256)].drop_duplicates(
        "sha256"
    )
    # A duplicate pixel under different IDs is removed above before splitting.
    train_idx, val_idx = next(
        StratifiedGroupKFold(n_splits=5, shuffle=True, random_state=seed).split(
            dev, dev.label, dev.patient
        )
    )
    train, val = dev.iloc[train_idx].copy(), dev.iloc[val_idx].copy()
    for split, part in [("train", train), ("val", val), ("test", test)]:
        part["split"] = split
        if set(part.label) != {0, 1}:
            raise ValueError(f"{split} phải có cả hai lớp.")
    result = pd.concat([train, val, test], ignore_index=True)
    validate_manifest(result)
    output.mkdir(parents=True, exist_ok=True)
    pd.crosstab(frame.original_split, frame.label).rename(
        columns={0: "NORMAL", 1: "PNEUMONIA"}
    ).to_csv(output / "dataset_counts_original.csv")
    result.to_csv(output / "manifest.csv", index=False)
    counts = pd.crosstab(result.split, result.label).reindex(["train", "val", "test"])
    counts.columns = ["NORMAL", "PNEUMONIA"]
    counts.to_csv(output / "dataset_counts.csv")
    counts.plot.bar(rot=0, ylabel="Images", title="Dataset distribution after leakage checks")
    plt.tight_layout()
    plt.savefig(output / "dataset_distribution.png", dpi=160)
    plt.close()
    save_json(
        output / "dataset_audit.json",
        dict(
            source_images=len(frame),
            retained_images=len(result),
            excluded_images=len(frame) - len(result),
            seed=seed,
            strategy="Official test; deduplicated pixels; filename-derived patient groups; 20% development validation",
            limitation="Patient IDs inferred from filenames, not independently verified clinical metadata.",
            manifest_sha256=hashlib.sha256((output / "manifest.csv").read_bytes()).hexdigest(),
        ),
    )
    return result


def validate_manifest(frame):
    required = {"path", "label", "patient", "sha256", "split"}
    if not required.issubset(frame.columns):
        raise ValueError("Manifest thiếu trường bắt buộc.")
    if set(frame.split) != {"train", "val", "test"} or set(frame.label) != {0, 1}:
        raise ValueError("Manifest phải có train/val/test và nhãn 0/1.")
    if frame[list(required)].isna().any().any() or frame.path.duplicated().any():
        raise ValueError("Manifest có giá trị trống hoặc đường dẫn trùng.")
    for key in ("patient", "sha256"):
        if frame.groupby(key).split.nunique().max() > 1:
            raise ValueError(f"Data leakage giữa các split: {key}")
    for split, part in frame.groupby("split"):
        if set(part.label) != {0, 1}:
            raise ValueError(f"{split} thiếu một lớp.")


class XrayDataset(Dataset):
    def __init__(self, frame, train=False):
        self.frame = frame.reset_index(drop=True)
        self.transform = image_transform(train)

    def __len__(self):
        return len(self.frame)

    def __getitem__(self, index):
        row = self.frame.iloc[index]
        with Image.open(row.path) as image:
            tensor = self.transform(image)
        return tensor, float(row.label), row.path


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--data", default="data/chest_xray")
    parser.add_argument("--output", default="results")
    args = parser.parse_args()
    prepare(args.data, args.output)
