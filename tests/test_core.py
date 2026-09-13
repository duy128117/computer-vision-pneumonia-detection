import numpy as np
import pytest
import torch
from PIL import Image
from src.metrics import binary_metrics
from src.dataset import patient_id
from src.transforms import image_transform
from src.gradcam import gradcam
from src.models import build_model
import pandas as pd
from src.dataset import validate_manifest
from src.train import run_epoch


def test_metrics_orientation():
    result = binary_metrics([0, 0, 1, 1], [0.1, 0.8, 0.3, 0.9])
    assert [result[k] for k in ("tn", "fp", "fn", "tp")] == [1, 1, 1, 1]
    assert result["specificity"] == result["recall"] == 0.5
    assert result["auc"] == 0.75


def test_patient_grouping():
    assert patient_id("person23_bacteria_11.jpeg") == patient_id("person23_virus_22.jpeg")
    assert patient_id("NORMAL2-IM-0123-0001.jpeg") == patient_id("NORMAL2-IM-0123-0002.jpeg")
    with pytest.raises(ValueError):
        patient_id("unknown.jpg")


def test_validation_deterministic():
    image = Image.fromarray(np.arange(4096, dtype=np.uint8).reshape(64, 64))
    transform = image_transform()
    assert torch.equal(transform(image), transform(image))
    assert transform(image).shape == (3, 224, 224)


def test_manifest_rejects_leakage():
    rows = [
        dict(
            path=f"{split}-{label}",
            label=label,
            patient=f"{split}-{label}",
            sha256=f"{split}-{label}",
            split=split,
        )
        for split in ("train", "val", "test")
        for label in (0, 1)
    ]
    frame = pd.DataFrame(rows)
    validate_manifest(frame)
    frame.loc[4, "patient"] = frame.loc[0, "patient"]
    with pytest.raises(ValueError, match="leakage"):
        validate_manifest(frame)


class TinyModel(torch.nn.Module):
    def __init__(self):
        super().__init__()
        self.layer4 = torch.nn.Conv2d(3, 1, 1, bias=False)
        self.bn = torch.nn.BatchNorm2d(1)
        self.fc = torch.nn.Linear(1, 1)

    def forward(self, x):
        return self.fc(self.bn(self.layer4(x)).mean((2, 3)))


def test_frozen_backbone_includes_batchnorm():
    model = TinyModel()
    for parameter in model.parameters():
        parameter.requires_grad_(False)
    for parameter in model.fc.parameters():
        parameter.requires_grad_(True)
    before_bn = model.bn.running_mean.clone()
    before_conv = model.layer4.weight.clone()
    before_head = model.fc.weight.clone()
    samples = [(torch.ones(3, 8, 8), float(i % 2), str(i)) for i in range(4)]
    loader = torch.utils.data.DataLoader(samples, batch_size=2)
    optimizer = torch.optim.SGD(model.fc.parameters(), lr=0.1)
    result = run_epoch(
        model, loader, torch.nn.BCEWithLogitsLoss(), torch.device("cpu"), optimizer, frozen=True
    )
    assert torch.equal(before_bn, model.bn.running_mean)
    assert torch.equal(before_conv, model.layer4.weight)
    assert not torch.equal(before_head, model.fc.weight)
    assert np.isfinite(result["loss"])


def test_gradcam_binary_target_sign():
    model = TinyModel().eval()
    with torch.no_grad():
        model.layer4.weight.fill_(1)
        model.fc.weight.fill_(1)
        model.fc.bias.zero_()
    x = torch.ones(1, 3, 8, 8)
    _, positive, _ = gradcam(model, x, 1)
    _, negative, _ = gradcam(model, x, 0)
    assert positive.max() > 0
    assert negative.max() == 0


@pytest.mark.parametrize("name", ["resnet50", "densenet121"])
def test_gradcam_architectures(name):
    torch.set_num_threads(2)
    model = build_model(name, pretrained=False)
    for target in (0, 1):
        probability, heat, returned = gradcam(model, torch.randn(1, 3, 64, 64), target)
        assert 0 <= probability <= 1
        assert np.isfinite(heat).all() and heat.min() >= 0 and heat.max() <= 1
        assert returned == target
    layer = model.layer4 if name == "resnet50" else model.features.norm5
    assert not layer._forward_hooks
