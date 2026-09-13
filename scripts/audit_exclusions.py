"""Explain exclusions using actual source files and the committed split manifest."""

import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import pandas as pd
from PIL import Image
from src.dataset import patient_id


def main():
    root = Path("data/chest_xray").resolve()
    retained = pd.read_csv("results/manifest.csv")
    kept = set(retained.path)
    test = retained[retained.split == "test"]
    test_patients = set(test.patient)
    test_hashes = set(test.sha256)
    original = []
    excluded = []
    for split in ("train", "val", "test"):
        for label in ("NORMAL", "PNEUMONIA"):
            for path in sorted((root / split / label).glob("*.jpeg")):
                original.append(dict(split=split, label=label))
                if str(path) in kept:
                    continue
                patient = patient_id(path)
                with Image.open(path) as image:
                    gray = image.convert("L")
                    digest = hashlib.sha256(str(gray.size).encode() + gray.tobytes()).hexdigest()
                if split == "test":
                    reason = "duplicate_pixels_within_test"
                elif patient in test_patients:
                    reason = "development_patient_overlaps_test"
                elif digest in test_hashes:
                    reason = "development_pixels_overlap_test"
                else:
                    reason = "duplicate_pixels_within_development"
                excluded.append(
                    dict(
                        path=str(path),
                        original_split=split,
                        label=label,
                        patient=patient,
                        sha256=digest,
                        reason=reason,
                    )
                )
    frame = pd.DataFrame(excluded)
    frame.to_csv("results/excluded_images.csv", index=False)
    source = pd.DataFrame(original)
    pd.crosstab(source.split, source.label).reindex(["train", "val", "test"]).to_csv(
        "results/dataset_counts_original.csv"
    )
    audit_path = Path("results/dataset_audit.json")
    audit = json.loads(audit_path.read_text(encoding="utf-8"))
    audit["exclusion_reasons"] = {k: int(v) for k, v in frame.reason.value_counts().items()}
    assert len(frame) == audit["excluded_images"]
    audit_path.write_text(json.dumps(audit, indent=2, ensure_ascii=False), encoding="utf-8")
    print(audit["exclusion_reasons"])


if __name__ == "__main__":
    main()
