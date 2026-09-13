"""Audit completed real-run artifacts without inventing or replacing metrics."""

import hashlib
import json
from pathlib import Path
import re
import sys
import zipfile

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import numpy as np
import pandas as pd
import torch
from src.metrics import binary_metrics
from src.dataset import validate_manifest


def main():
    root = Path(__file__).resolve().parents[1]
    results = root / "results"
    manifest_path = results / "manifest.csv"
    manifest = pd.read_csv(manifest_path)
    validate_manifest(manifest)
    test = manifest[manifest.split == "test"].set_index("path")
    required = ["loss", "accuracy", "precision", "recall", "f1", "auc"]
    for name in ("resnet50", "densenet121"):
        checkpoint = torch.load(
            root / "checkpoints" / f"best_{name}.pth", map_location="cpu", weights_only=True
        )
        assert checkpoint["trained"] and checkpoint["model"] == name
        assert (
            checkpoint["manifest_sha256"] == hashlib.sha256(manifest_path.read_bytes()).hexdigest()
        )
        history = pd.read_csv(root / "logs" / f"history_{name}.csv")
        assert set(history.stage) == {1, 2}
        assert all(
            f"{split}_{metric}" in history for split in ("train", "val") for metric in required
        )
        best = history.loc[history.val_loss.idxmin()]
        assert checkpoint["epoch"] == int(best.epoch)
        assert np.isclose(checkpoint["val_metrics"]["loss"], best.val_loss)
        assert (root / "logs" / f"run_{name}.json").is_file()
        predictions = pd.read_csv(results / "predictions" / f"{name}.csv").set_index("path")
        assert predictions.index.is_unique and set(predictions.index) == set(test.index)
        assert (predictions.label == test.loc[predictions.index].label).all()
        expected = binary_metrics(predictions.label, predictions.probability)
        saved = json.loads((results / f"{name}_metrics.json").read_text())
        for key, value in expected.items():
            assert np.isclose(value, saved[key]), (name, key)
        assert sum(saved[k] for k in ("tp", "tn", "fp", "fn")) == len(test)
        print(f"{name}: checkpoint, history, test coverage and recomputed metrics OK")
    report = (results / "report.md").read_text(encoding="utf-8")
    assert "| Model | Accuracy" in report
    assert "| Model | Accuracy | Precision | Sensitivity | Specificity | F1 | AUC | FN |\n|---|" in report
    for path in re.findall(r"!\[[^\]]*\]\(([^)]+)\)", report):
        assert (results / path).is_file(), path
    with zipfile.ZipFile(results / "presentation.pptx") as deck:
        assert deck.testzip() is None
        print(
            "Slides:",
            len([p for p in deck.namelist() if re.fullmatch(r"ppt/slides/slide\d+\.xml", p)]),
        )
    print("Artifact audit: PASS")


if __name__ == "__main__":
    main()
