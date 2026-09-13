import argparse
import hashlib
import time
from pathlib import Path
import pandas as pd
import torch
import yaml
from torch.utils.data import DataLoader
from .dataset import XrayDataset, validate_manifest
from .metrics import binary_metrics
from .models import build_model, classifier
from .utils import device, seed_everything, save_json


def run_epoch(model, loader, criterion, target_device, optimizer=None, frozen=False, scaler=None):
    model.train(optimizer is not None)
    if frozen:
        model.eval()
        classifier(model).train()
    total, labels, probabilities = 0.0, [], []
    for batch_index, (images, targets, _) in enumerate(loader, 1):
        images = images.to(target_device)
        targets = targets.to(target_device, dtype=torch.float32)
        with torch.set_grad_enabled(optimizer is not None):
            with torch.autocast(
                device_type=target_device.type, enabled=scaler is not None and scaler.is_enabled()
            ):
                logits = model(images).flatten()
                loss = criterion(logits, targets)
            if not torch.isfinite(loss):
                raise RuntimeError("Non-finite loss")
            if optimizer is not None:
                optimizer.zero_grad(set_to_none=True)
                if scaler is not None and scaler.is_enabled():
                    scaler.scale(loss).backward()
                    scaler.step(optimizer)
                    scaler.update()
                else:
                    loss.backward()
                    optimizer.step()
        total += loss.item() * len(targets)
        labels.extend(targets.detach().cpu().tolist())
        probabilities.extend(logits.detach().sigmoid().cpu().tolist())
        if optimizer is not None and batch_index % 100 == 0:
            print(f"  batch {batch_index}/{len(loader)}", flush=True)
    return dict(loss=total / len(loader.dataset), **binary_metrics(labels, probabilities))


def train(config_path):
    config = yaml.safe_load(Path(config_path).read_text(encoding="utf-8"))
    seed_everything(config["seed"])
    torch.set_num_threads(config.get("threads", 4))
    target_device = device()
    output = Path(config.get("output", "."))
    name = config["model"]
    for folder in ("checkpoints", "logs", "results"):
        (output / folder).mkdir(parents=True, exist_ok=True)
    checkpoint_path = output / "checkpoints" / f"best_{name}.pth"
    history_path = output / "logs" / f"history_{name}.csv"
    if checkpoint_path.exists() or history_path.exists():
        raise FileExistsError("Choose a new output directory to preserve existing run.")
    manifest = Path(config.get("manifest", "results/manifest.csv"))
    frame = pd.read_csv(manifest)
    validate_manifest(frame)
    if config["head_epochs"] < 1 or config["finetune_epochs"] < 1:
        raise ValueError("Both training stages require at least one epoch.")
    loaders = {
        split: DataLoader(
            XrayDataset(frame[frame.split == split], train=split == "train"),
            batch_size=config["batch_size"],
            shuffle=split == "train",
            num_workers=config.get("workers", 0),
            pin_memory=target_device.type == "cuda",
        )
        for split in ("train", "val")
    }
    counts = frame[frame.split == "train"].label.value_counts()
    weight = float(counts[0] / counts[1]) if config["weighted_loss"] else 1.0
    criterion = torch.nn.BCEWithLogitsLoss(pos_weight=torch.tensor(weight, device=target_device))
    model = build_model(name).to(target_device)
    scaler = torch.amp.GradScaler(
        "cuda", enabled=target_device.type == "cuda" and config.get("amp", True)
    )
    history, best, epoch = [], float("inf"), 0
    start = time.time()
    for stage, epochs, lr in [
        (1, config["head_epochs"], config["head_lr"]),
        (2, config["finetune_epochs"], config["finetune_lr"]),
    ]:
        if stage == 2 and checkpoint_path.exists():
            model.load_state_dict(
                torch.load(checkpoint_path, map_location=target_device, weights_only=True)[
                    "state_dict"
                ]
            )
        for parameter in model.parameters():
            parameter.requires_grad_(stage == 2)
        for parameter in classifier(model).parameters():
            parameter.requires_grad_(True)
        optimizer = torch.optim.AdamW(
            filter(lambda p: p.requires_grad, model.parameters()),
            lr=lr,
            weight_decay=config["weight_decay"],
        )
        scheduler = torch.optim.lr_scheduler.ReduceLROnPlateau(optimizer, patience=2, factor=0.5)
        stale = 0
        for _ in range(epochs):
            epoch += 1
            train_metrics = run_epoch(
                model,
                loaders["train"],
                criterion,
                target_device,
                optimizer,
                frozen=stage == 1,
                scaler=scaler,
            )
            val_metrics = run_epoch(model, loaders["val"], criterion, target_device, scaler=scaler)
            history.append(
                dict(
                    epoch=epoch,
                    stage=stage,
                    learning_rate=optimizer.param_groups[0]["lr"],
                    **{"train_" + k: v for k, v in train_metrics.items()},
                    **{"val_" + k: v for k, v in val_metrics.items()},
                )
            )
            pd.DataFrame(history).to_csv(history_path, index=False)
            print(
                f'{name} stage={stage} epoch={epoch} train_loss={train_metrics["loss"]:.4f} val_loss={val_metrics["loss"]:.4f} val_auc={val_metrics["auc"]:.4f}',
                flush=True,
            )
            if val_metrics["loss"] < best:
                best, stale = val_metrics["loss"], 0
                torch.save(
                    dict(
                        model=name,
                        trained=True,
                        state_dict=model.state_dict(),
                        epoch=epoch,
                        stage=stage,
                        threshold=0.5,
                        config=config,
                        val_metrics=val_metrics,
                        manifest_sha256=hashlib.sha256(manifest.read_bytes()).hexdigest(),
                        pretrained="ImageNet DEFAULT torchvision",
                        pos_weight=weight,
                    ),
                    checkpoint_path,
                )
            else:
                stale += 1
            scheduler.step(val_metrics["loss"])
            if stale >= config["patience"]:
                break
    save_json(
        output / "logs" / f"run_{name}.json",
        dict(
            config=config,
            seconds=time.time() - start,
            device=str(target_device),
            torch=str(torch.__version__),
            epochs=epoch,
            pos_weight=weight,
        ),
    )
    return checkpoint_path


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--config", required=True)
    train(parser.parse_args().config)
