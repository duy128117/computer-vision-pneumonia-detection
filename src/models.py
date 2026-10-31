import os
from pathlib import Path
import torch
import torch.nn as nn
from torchvision.models import (
    DenseNet121_Weights,
    Inception_V3_Weights,
    ResNet50_Weights,
    densenet121,
    inception_v3,
    resnet50,
)

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


class SpatialLungRoIAttention(nn.Module):
    """
    Spatial Lung-RoI Guided Attention Gate.
    Applies channel-pooling (avg + max) -> 7x7 Conv -> Sigmoid -> Spatial Attention Map Ms.
    Refines feature map via residual scaling: F_out = F * Ms + F.
    Actively suppresses non-pulmonary background shortcuts (clavicle, neck, boundaries).
    """
    def __init__(self, in_channels=2048, kernel_size=7):
        super().__init__()
        assert kernel_size in (3, 7), "Kernel size must be 3 or 7"
        padding = 3 if kernel_size == 7 else 1
        self.conv = nn.Conv2d(2, 1, kernel_size=kernel_size, padding=padding, bias=False)
        self.sigmoid = nn.Sigmoid()
        self.attention_map = None

    def forward(self, x):
        avg_out = torch.mean(x, dim=1, keepdim=True)
        max_out, _ = torch.max(x, dim=1, keepdim=True)
        scale = torch.cat([avg_out, max_out], dim=1)
        att = self.sigmoid(self.conv(scale))
        self.attention_map = att
        return x * att + x


class ResNet50_LungRoI_Attention(nn.Module):
    """
    ResNet50 integrated with Spatial Lung-RoI Guided Attention Gate.
    Designed specifically to eliminate clavicle/shoulder shortcut learning.
    """
    def __init__(self, pretrained=True):
        super().__init__()
        self.backbone = resnet50(weights=ResNet50_Weights.DEFAULT if pretrained else None)
        in_features = self.backbone.fc.in_features
        self.backbone.fc = nn.Identity()
        
        # Spatial Lung-RoI Attention Gate
        self.attention_gate = SpatialLungRoIAttention(in_channels=in_features, kernel_size=7)
        self.global_pool = nn.AdaptiveAvgPool2d((1, 1))
        
        # Classification Head with regularizing Dropout
        self.classifier = nn.Sequential(
            nn.Dropout(0.3),
            nn.Linear(in_features, 1)
        )

    def forward(self, x):
        x = self.backbone.conv1(x)
        x = self.backbone.bn1(x)
        x = self.backbone.relu(x)
        x = self.backbone.maxpool(x)

        x = self.backbone.layer1(x)
        x = self.backbone.layer2(x)
        x = self.backbone.layer3(x)
        x = self.backbone.layer4(x)

        x = self.attention_gate(x)

        x = self.global_pool(x)
        x = torch.flatten(x, 1)
        x = self.classifier(x)
        return x


def build_model(name, pretrained=True):
    if name == "resnet50":
        model = resnet50(weights=ResNet50_Weights.DEFAULT if pretrained else None)
        model.fc = torch.nn.Linear(model.fc.in_features, 1)
    elif name in ("resnet50_attention", "resnet50_roi_attention"):
        model = ResNet50_LungRoI_Attention(pretrained=pretrained)
    elif name == "densenet121":
        model = densenet121(weights=DenseNet121_Weights.DEFAULT if pretrained else None)
        model.classifier = torch.nn.Linear(model.classifier.in_features, 1)
    elif name == "inceptionv3":
        # Disable the auxiliary classifier so every architecture exposes the
        # same single-logit interface expected by the binary training loop.
        # Torchvision requires AuxLogits while loading its pretrained state,
        # so remove that branch immediately after the weights are loaded.
        model = inception_v3(
            weights=Inception_V3_Weights.DEFAULT if pretrained else None,
            aux_logits=pretrained,
            init_weights=False,
            # Keep checkpoint inference identical to pretrained training.
            # Torchvision enables this automatically only when weights are
            # supplied, while load_checkpoint builds without downloading them.
            transform_input=True,
        )
        model.aux_logits = False
        model.AuxLogits = None
        model.fc = torch.nn.Linear(model.fc.in_features, 1)
    elif name == "custom_model":
        model = CustomPneumoniaModel(pretrained=pretrained)
    else:
        raise ValueError(f"Unknown architecture: {name}")
    return model


def classifier(model):
    if hasattr(model, "custom_classifier"):
        return model.custom_classifier
    if hasattr(model, "attention_gate"):
        return model.classifier
    return model.fc if hasattr(model, "fc") else model.classifier


def load_checkpoint(path, device):
    checkpoint = torch.load(path, map_location=device, weights_only=True)
    if checkpoint.get("trained") is not True:
        raise ValueError("Checkpoint chưa được huấn luyện bởi pipeline này.")
    model = build_model(checkpoint["model"], pretrained=False).to(device)
    model.load_state_dict(checkpoint["state_dict"])
    model.eval()
    return model, checkpoint
