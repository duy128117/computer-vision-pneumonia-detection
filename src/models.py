import os
from pathlib import Path
import torch
import torch.nn as nn
from torchvision.models import resnet50, densenet121, ResNet50_Weights, DenseNet121_Weights

# Keep large pretrained downloads off the user's system drive by default.
torch.hub.set_dir(
    str(Path(os.environ.get("TORCH_HOME", Path(__file__).resolve().parents[1] / ".torch")) / "hub")
)


class CustomPneumoniaModel(nn.Module):
    def __init__(self, pretrained=True):
        super().__init__()
        self.backbone = densenet121(weights=DenseNet121_Weights.DEFAULT if pretrained else None)
        in_features = self.backbone.classifier.in_features
        self.backbone.classifier = nn.Identity()
        
        # Tùy biến kiến trúc: Thêm classification head với Dropout, BatchNorm và các lớp Linear
        self.custom_classifier = nn.Sequential(
            nn.BatchNorm1d(in_features),
            nn.Dropout(0.5),
            nn.Linear(in_features, 512),
            nn.ReLU(),
            nn.BatchNorm1d(512),
            nn.Dropout(0.3),
            nn.Linear(512, 1)
        )
        
    def forward(self, x):
        features = self.backbone(x)
        return self.custom_classifier(features)

def build_model(name, pretrained=True):
    if name == "resnet50":
        model = resnet50(weights=ResNet50_Weights.DEFAULT if pretrained else None)
        model.fc = torch.nn.Linear(model.fc.in_features, 1)
    elif name == "densenet121":
        model = densenet121(weights=DenseNet121_Weights.DEFAULT if pretrained else None)
        model.classifier = torch.nn.Linear(model.classifier.in_features, 1)
    elif name == "custom_model":
        model = CustomPneumoniaModel(pretrained=pretrained)
    else:
        raise ValueError(f"Unknown architecture: {name}")
    return model


def classifier(model):
    if hasattr(model, "custom_classifier"):
        return model.custom_classifier
    return model.fc if hasattr(model, "fc") else model.classifier


def load_checkpoint(path, device):
    checkpoint = torch.load(path, map_location=device, weights_only=True)
    if checkpoint.get("trained") is not True:
        raise ValueError("Checkpoint chưa được huấn luyện bởi pipeline này.")
    model = build_model(checkpoint["model"], pretrained=False).to(device)
    model.load_state_dict(checkpoint["state_dict"])
    model.eval()
    return model, checkpoint
