import argparse
import hashlib
import shutil
from pathlib import Path
import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
import torch
from PIL import Image
from sklearn.metrics import ConfusionMatrixDisplay, roc_curve
from torch.utils.data import DataLoader
from .dataset import XrayDataset
from .models import load_checkpoint
from .metrics import binary_metrics
from .inference import predict
from .utils import device, save_json


def training_plots(history_path, name, output):
    history = pd.read_csv(history_path)
    groups = {
        "loss_curve": ["train_loss", "val_loss"],
        "accuracy_curve": ["train_accuracy", "val_accuracy"],
        "precision_recall": ["val_precision", "val_recall"],
        "f1_curve": ["val_f1"],
        "auc_curve": ["val_auc"],
    }
    for suffix, columns in groups.items():
        axis = history.plot(x="epoch", y=columns, title=name, grid=True)
        axis.set_ylabel(suffix)
        plt.tight_layout()
        plt.savefig(output / f"{name}_{suffix}.png", dpi=160)
        plt.close()


def evaluate(checkpoint, manifest="results/manifest.csv", output="."):
    target_device = device()
    torch.set_num_threads(4)
    model, metadata = load_checkpoint(checkpoint, target_device)
    if hashlib.sha256(Path(manifest).read_bytes()).hexdigest() != metadata["manifest_sha256"]:
        raise ValueError("Manifest does not match training checkpoint.")
    name, root = metadata["model"], Path(output)
    results = root / "results"
    for folder in ("plots", "confusion_matrix", "roc", "gradcam", "predictions", "errors"):
        (results / folder).mkdir(parents=True, exist_ok=True)
    frame = pd.read_csv(manifest)
    frame = frame[frame.split == "test"]
    rows = []
    with torch.no_grad():
        for images, labels, paths in DataLoader(XrayDataset(frame), batch_size=8):
            probabilities = model(images.to(target_device)).flatten().sigmoid().cpu().tolist()
            for path, label, probability in zip(paths, labels.tolist(), probabilities):
                predicted = int(probability >= metadata["threshold"])
                case = {(0, 0): "TN", (0, 1): "FP", (1, 0): "FN", (1, 1): "TP"}[
                    (int(label), predicted)
                ]
                rows.append(
                    dict(
                        path=path,
                        label=int(label),
                        probability=probability,
                        predicted=predicted,
                        case=case,
                    )
                )
    predictions = pd.DataFrame(rows)
    predictions.to_csv(results / "predictions" / f"{name}.csv", index=False)
    metrics = binary_metrics(predictions.label, predictions.probability, metadata["threshold"])
    save_json(
        results / f"{name}_metrics.json",
        dict(
            model=name,
            **metrics,
            checkpoint_sha256=hashlib.sha256(Path(checkpoint).read_bytes()).hexdigest(),
            validation=metadata["val_metrics"],
            epoch=metadata["epoch"],
        ),
    )
    ConfusionMatrixDisplay.from_predictions(
        predictions.label,
        predictions.predicted,
        labels=[0, 1],
        display_labels=["NORMAL", "PNEUMONIA"],
    )
    plt.tight_layout()
    plt.savefig(results / "confusion_matrix" / f"{name}_confusion_matrix.png", dpi=160)
    plt.close()
    fpr, tpr, _ = roc_curve(predictions.label, predictions.probability)
    plt.plot(fpr, tpr, label=f'{name} AUC={metrics["auc"]:.4f}')
    plt.plot([0, 1], [0, 1], "--", label="Random")
    plt.xlabel("False positive rate")
    plt.ylabel("True positive rate")
    plt.legend()
    plt.savefig(results / "roc" / f"{name}_roc.png", dpi=160)
    plt.close()
    training_plots(root / "logs" / f"history_{name}.csv", name, results / "plots")
    cases = {}
    for case in ("TP", "TN", "FP", "FN"):
        subset = predictions[predictions.case == case]
        cases[case] = len(subset)
        for index, row in enumerate(subset.head(4).itertuples(), 1):
            with Image.open(row.path) as image:
                _, images = predict(model, metadata, image)
            folder = results / "gradcam" / name
            folder.mkdir(parents=True, exist_ok=True)
            for suffix, image in zip(("original", "heatmap", "overlay"), images):
                image.save(folder / f"{case}_{index:03}_{suffix}.png")
        if case in ("FP", "FN"):
            folder = (
                results / "errors" / name / ("false_positive" if case == "FP" else "false_negative")
            )
            folder.mkdir(parents=True, exist_ok=True)
            for row in subset.itertuples():
                shutil.copy2(row.path, folder / Path(row.path).name)
    save_json(results / "gradcam" / f"{name}_cases.json", cases)
    print(metrics, flush=True)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--checkpoint", required=True)
    parser.add_argument("--manifest", default="results/manifest.csv")
    parser.add_argument("--output", default=".")
    args = parser.parse_args()
    evaluate(args.checkpoint, args.manifest, args.output)
