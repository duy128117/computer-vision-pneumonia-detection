import os
from pathlib import Path
import torch
from torchvision.models import resnet50, densenet121, ResNet50_Weights, DenseNet121_Weights

# Keep large pretrained downloads off the user's system drive by default.
torch.hub.set_dir(
    str(Path(os.environ.get("TORCH_HOME", Path(__file__).resolve().parents[1] / ".torch")) / "hub")
)


def build_model(name, pretrained=True):
    if name == "resnet50":
        model = resnet50(weights=ResNet50_Weights.DEFAULT if pretrained else None)
        model.fc = torch.nn.Linear(model.fc.in_features, 1)
    elif name == "densenet121":
        model = densenet121(weights=DenseNet121_Weights.DEFAULT if pretrained else None)
        model.classifier = torch.nn.Linear(model.classifier.in_features, 1)
    else:
        raise ValueError(f"Unknown architecture: {name}")
    return model


def classifier(model):
    return model.fc if hasattr(model, "fc") else model.classifier


def load_checkpoint(path, device):
    checkpoint = torch.load(path, map_location=device, weights_only=True)
    if checkpoint.get("trained") is not True:
        raise ValueError("Checkpoint chưa được huấn luyện bởi pipeline này.")
    model = build_model(checkpoint["model"], pretrained=False).to(device)
    model.load_state_dict(checkpoint["state_dict"])
    model.eval()
    return model, checkpoint
